"""
검사성적서 자동 생성 — 사내용 웹앱 (미리보기 + 엑셀 다운로드)
"""
import io, os, re, shutil, tempfile, zipfile
from datetime import datetime

import pandas as pd
import streamlit as st
import fitz  # PyMuPDF

from generate_reports import (compute_stats, fill_template, xlsx_to_pdf, merge_pdfs,
                                 list_printers, get_default_printer, print_pdf)

import json
import statistics

HERE = os.path.dirname(os.path.abspath(__file__))

# 거래처 설정 로드
VENDORS_PATH = os.path.join(HERE, "vendors.json")
with open(VENDORS_PATH, "r", encoding="utf-8") as _f:
    VENDORS = json.load(_f)

st.set_page_config(page_title="검사성적서 자동 생성", page_icon="📋", layout="wide")
st.title("📋 검사성적서 자동 생성")
st.caption("ZIP 업로드 → 자동 통계 → 미리보기 → PDF·엑셀 다운로드. 부적합은 화면에 별도 강조됩니다.")

with st.sidebar:
    st.header("🏢 거래처")
    vendor_keys = list(VENDORS.keys())
    vendor_key = st.selectbox(
        "거래처 선택",
        vendor_keys,
        format_func=lambda k: VENDORS[k]["name"],
        key="vendor_key",
    )
    vendor   = VENDORS[vendor_key]
    spec_d   = vendor["spec"]
    template_path = os.path.join(HERE, vendor["template"])
    conc_count    = vendor["concentricity_count"]
    has_equip_no  = vendor.get("has_equip_no", True)
    st.caption(f"동심도 샘플 {conc_count}개 · 장비번호 {'사용' if has_equip_no else '미사용'}")

    st.divider()
    st.header("⚙ 판정 기준")
    spec_size_lo = st.number_input("Size(직경) 하한",  value=spec_d["size_lo"], format="%.3f", step=0.001, key=f"size_lo_{vendor_key}")
    spec_size_hi = st.number_input("Size(직경) 상한",  value=spec_d["size_hi"], format="%.3f", step=0.001, key=f"size_hi_{vendor_key}")
    spec_round_max = st.number_input("Roundness(진원도) MAX ≤", value=spec_d["round_max"], format="%.4f", step=0.0001, key=f"round_max_{vendor_key}")
    if spec_d.get("round_avg") is not None:
        spec_round_avg = st.number_input("Roundness 평균 ≤", value=spec_d["round_avg"], format="%.4f", step=0.0001, key=f"round_avg_{vendor_key}")
    else:
        spec_round_avg = None
    spec_pos_max = st.number_input("Position(위치도) MAX ≤", value=spec_d["pos_max"], format="%.3f", step=0.001, key=f"pos_max_{vendor_key}")
    spec_conc_max = st.number_input("Concentricity(동심도) MAX ≤", value=spec_d["conc_max"], format="%.3f", step=0.001, key=f"conc_max_{vendor_key}")
    if spec_d.get("conc_avg") is not None:
        spec_conc_avg = st.number_input("Concentricity 평균 ≤", value=spec_d["conc_avg"], format="%.3f", step=0.001, key=f"conc_avg_{vendor_key}")
    else:
        spec_conc_avg = None

    st.divider()
    st.header("👤 서명자 기본값")
    default_insp = st.text_input("검사자", value="박성진")
    default_conf = st.text_input("확인자", value="김민지")
    default_appr = st.text_input("승인자", value="안예슬")
    st.divider()
    preview_dpi = st.slider("미리보기 해상도 (DPI)", 80, 200, 120, 10)

    st.divider()
    st.header("🖨 프린터 설정")
    _printers = list_printers()
    if _printers:
        _default = get_default_printer() or _printers[0]
        try:
            _idx = _printers.index(_default)
        except ValueError:
            _idx = 0
        printer_name = st.selectbox("프린터", _printers, index=_idx)
        print_color  = st.radio("색상", ["컬러", "흑백"], horizontal=True) == "컬러"
        print_duplex = st.radio("인쇄 면", ["단면", "양면(긴 쪽)"], horizontal=True) == "양면(긴 쪽)"
        print_copies = st.number_input("부수", min_value=1, max_value=20, value=1, step=1)
    else:
        st.caption("Windows 에서만 프린터 출력 가능. 현재 프린터 목록 없음.")
        printer_name = None
        print_color = True
        print_duplex = False
        print_copies = 1

st.header("1️⃣ 측정 데이터 ZIP 업로드")
st.markdown(
    "**압축 구조 예시**\n```\n내_데이터.zip\n├─ 004 #11/  (004 A.xls, 004 B.xls)\n├─ 005 #4/   (005 A.xls, 005 B.xls)\n└─ ...\n```"
)
up = st.file_uploader("ZIP 파일", type=["zip"], accept_multiple_files=False)

if not up:
    st.info("👆 위에 ZIP 파일을 올려 주세요.")
    st.stop()

@st.cache_data(show_spinner=False)
def unzip_and_scan(zip_bytes: bytes):
    work = tempfile.mkdtemp(prefix="insp_")
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        z.extractall(work)
    folders = []
    for root, dirs, files in os.walk(work):
        a = [f for f in files if f.lower().endswith("a.xls")]
        b = [f for f in files if f.lower().endswith("b.xls")]
        if a and b:
            folders.append(root)
    return work, sorted(folders, key=lambda p: os.path.basename(p))

work, candidate_folders = unzip_and_scan(up.getvalue())

if not candidate_folders:
    st.error("❌ A.xls + B.xls 가 함께 있는 폴더를 찾지 못했습니다.")
    st.stop()

st.success(f"✅ {len(candidate_folders)}개 폴더 발견")

st.header("2️⃣ 검사성적서 메타데이터")
st.caption("폴더명에서 자동 추출되는 항목은 미리 채워둡니다. 빈칸을 채우거나 수정하세요.")

today_kr = datetime.now().strftime("%Y.%-m.%-d") if os.name != "nt" else datetime.now().strftime("%Y.%#m.%#d")

def _sanitize_filename(name: str) -> str:
    """Windows 파일명 금지 문자 제거 + 앞뒤 공백/점 정리. 빈 문자열은 그대로 둔다."""
    if name is None:
        return ""
    s = str(name).strip()
    s = re.sub(r'[\\/:*?"<>|]', "", s)
    s = re.sub(r"\.(pdf|xlsx|zip)$", "", s, flags=re.I)
    return s.strip(" .")

default_rows = []
# 거래처별 기본값 (제품명·관리번호 양식)
v_defaults = vendor.get("defaults", {})
default_product_name = v_defaults.get("product_name", "")
mgmt_no_template     = v_defaults.get("mgmt_no_template", "")  # "1660-{yymm}-{folder_num}" 같은 형식

for folder in candidate_folders:
    name = os.path.basename(folder)
    m_eq = re.search(r"#(\d+)", name)
    m_no = re.match(r"(\d+)", name)
    folder_num = m_no.group(1) if m_no else ""
    yymm = datetime.now().strftime("%y%m")
    default_filename = name.replace("#", "no").replace(" ", "_")
    if mgmt_no_template and folder_num:
        default_mgmt = mgmt_no_template.format(yymm=yymm, folder_num=folder_num)
    else:
        default_mgmt = ""
    default_rows.append({
        "folder": name,
        "product_name":    default_product_name,
        "processing_date": today_kr,
        "mgmt_no":         default_mgmt,
        "inspection_date": today_kr,
        "equip_no":        f"{m_eq.group(1)}호기" if m_eq else "",
        "tool_sn":         "",
        "insp_name":       default_insp,
        "conf_name":       default_conf,
        "appr_name":       default_appr,
        "pdf_filename":    default_filename,
    })

df = pd.DataFrame(default_rows)
edited = st.data_editor(
    df, num_rows="fixed", use_container_width=True, hide_index=True,
    column_config={
        "folder":          st.column_config.TextColumn("폴더", disabled=True),
        "product_name":    st.column_config.TextColumn("제품명"),
        "processing_date": st.column_config.TextColumn("가공일"),
        "mgmt_no":         st.column_config.TextColumn("관리번호"),
        "inspection_date": st.column_config.TextColumn("검사일"),
        "equip_no":        st.column_config.TextColumn("장비번호"),
        "tool_sn":         st.column_config.TextColumn("공구 S/N"),
        "insp_name":       st.column_config.TextColumn("검사자"),
        "conf_name":       st.column_config.TextColumn("확인자"),
        "appr_name":       st.column_config.TextColumn("승인자"),
        "pdf_filename":    st.column_config.TextColumn(
            "📄 파일명 (확장자 제외)",
            help="비우면 폴더명을 자동 변환한 기본값 사용. PDF·Excel 모두 동일 이름.",
        ),
    },
)

st.header("3️⃣ 검사성적서 생성")

zip_basename = os.path.splitext(up.name)[0]

st.markdown("**📁 합본 파일명** _(확장자 제외, 비우면 기본값 사용)_")
col_m, col_x = st.columns(2)
merged_filename = col_m.text_input(
    "합본 PDF",
    value=f"검사성적서_{zip_basename}",
    help="모든 검사성적서를 하나로 합친 PDF 파일의 이름",
)
xlsx_zip_filename = col_x.text_input(
    "엑셀 모음 ZIP",
    value=f"검사성적서_엑셀_{zip_basename}",
    help="개별 검사성적서 엑셀 파일들을 모은 ZIP 의 이름",
)

col_pv, col_go = st.columns(2)
preview = col_pv.button("🔍 미리보기 (첫 1건만)", use_container_width=True,
                          help="첫 폴더 한 건만 빠르게 PDF로 만들어 미리보기. 양식·수치 확인용.")
go = col_go.button("🚀 PDF 전체 생성", type="primary", use_container_width=True,
                     help="모든 폴더를 처리하고 합본 PDF 까지 생성합니다.")

if not preview and not go:
    st.info("👆 미리보기로 양식부터 확인해 보세요. 괜찮으면 옆의 PDF 전체 생성 클릭.")
    st.stop()

is_preview = preview and not go
records = edited.to_dict("records")
if is_preview:
    records = records[:1]
    st.info("🔍 미리보기 모드 — 첫 폴더 1건만 빠르게 처리합니다.")

out_dir = os.path.join(work, "out")
os.makedirs(out_dir, exist_ok=True)

results = []
prog = st.progress(0.0, text="처리 시작...")
for i, row in enumerate(records):
    folder_path = next((f for f in candidate_folders
                        if os.path.basename(f) == row["folder"]), None)
    if not folder_path:
        results.append({"folder": row["folder"], "ok": False, "error": "폴더 없음"})
        continue
    prog.progress((i + 0.3) / len(edited), text=f"통계 계산 중: {row['folder']}")
    try:
        stats = compute_stats(folder_path)
        stats["S1_SIZE_PF"] = "PASS" if (spec_size_lo <= stats["Size"]["MIN"]
                                          and stats["Size"]["MAX"] <= spec_size_hi) else "FAIL"
        rnd_ok = stats["Round"]["MAX"] <= spec_round_max
        if spec_round_avg is not None:
            rnd_ok = rnd_ok and stats["Round"]["AVG"] <= spec_round_avg
        stats["S1_RND_PF"]  = "PASS" if rnd_ok else "FAIL"
        stats["S2_POS_PF"]  = "PASS" if stats["Pos"]["MAX"] <= spec_pos_max else "FAIL"
        conc_used = [c for c in stats["Conc"][:conc_count] if c is not None]
        if conc_used:
            conc_ok = max(conc_used) <= spec_conc_max
            if spec_conc_avg is not None:
                conc_ok = conc_ok and statistics.mean(conc_used) <= spec_conc_avg
        else:
            conc_ok = False
        stats["S3_CONC_PF"] = "PASS" if conc_ok else "FAIL"

        prog.progress((i + 0.6) / len(edited), text=f"PDF 생성 중: {row['folder']}")
        user_filename = _sanitize_filename(row.get("pdf_filename", ""))
        if not user_filename:
            user_filename = row["folder"].replace("#", "no").replace(" ", "_")
        safe = user_filename
        xlsx_out = os.path.join(out_dir, f"{safe}.xlsx")
        fill_template(template_path, xlsx_out, stats, row)
        pdf_out = xlsx_to_pdf(xlsx_out, out_dir)
        results.append({"folder": row["folder"], "ok": True, "safe": safe,
                        "stats": stats, "pdf": pdf_out, "xlsx": xlsx_out})
    except Exception as e:
        results.append({"folder": row["folder"], "ok": False, "error": str(e)})
    prog.progress((i + 1) / len(edited), text=f"{row['folder']} 완료")
prog.empty()

for fn in os.listdir(out_dir):
    if fn.startswith(".~lock") or fn.endswith(".tmp"):
        try: os.remove(os.path.join(out_dir, fn))
        except OSError: pass

pdfs = [r["pdf"] for r in results if r.get("ok") and r.get("pdf")]
date_tag = datetime.now().strftime("%Y-%m-%d")

_m = _sanitize_filename(merged_filename) or f"검사성적서_{zip_basename}"
_x = _sanitize_filename(xlsx_zip_filename) or f"검사성적서_엑셀_{zip_basename}"
merged_name   = f"{_m}.pdf"
xlsx_zip_name = f"{_x}.zip"

merged = os.path.join(out_dir, merged_name)
if len(pdfs) > 1:
    merge_pdfs(pdfs, merged)
elif len(pdfs) == 1:
    shutil.copyfile(pdfs[0], merged)

all_xlsx_zip = os.path.join(out_dir, xlsx_zip_name)
if not is_preview:
    with zipfile.ZipFile(all_xlsx_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in results:
            if r.get("ok") and r.get("xlsx"):
                zf.write(r["xlsx"], arcname=os.path.basename(r["xlsx"]))

st.header("4️⃣ 결과")

failures = []
for r in results:
    if not r["ok"]:
        failures.append((r["folder"], "오류", r.get("error", "")))
        continue
    s = r["stats"]
    items = [
        ("Size",          s["S1_SIZE_PF"], f"MIN={s['Size']['MIN']:.3f}, MAX={s['Size']['MAX']:.3f}"),
        ("Roundness",     s["S1_RND_PF"],  f"MAX={s['Round']['MAX']:.4f}"),
        ("Position",      s["S2_POS_PF"],  f"MAX={s['Pos']['MAX']:.3f}"),
        ("Concentricity", s["S3_CONC_PF"], "8개 중 한도 초과"),
    ]
    for label, pf, detail in items:
        if pf != "PASS":
            failures.append((r["folder"], label, detail))

if failures:
    st.error(f"🚨 부적합 {len(failures)}건 발견 — 거래처 발송 전 확인 필요!")
    st.dataframe(pd.DataFrame(failures, columns=["폴더", "항목", "내용"]),
                 use_container_width=True, hide_index=True)
else:
    ok_count = sum(1 for r in results if r["ok"])
    st.success(f"✅ 전체 {ok_count}건 모두 PASS")

summary_rows = []
for r in results:
    if not r["ok"]:
        summary_rows.append({"폴더": r["folder"], "홀수": "-",
            "Size": "❌", "Roundness": "❌", "Position": "❌",
            "Concentricity": "❌", "비고": r.get("error", "오류")})
        continue
    s = r["stats"]
    pf = lambda p: "✅" if p == "PASS" else "❌"
    summary_rows.append({
        "폴더": r["folder"], "홀수": s["_count"],
        "Size":          pf(s["S1_SIZE_PF"]),
        "Roundness":     pf(s["S1_RND_PF"]),
        "Position":      pf(s["S2_POS_PF"]),
        "Concentricity": pf(s["S3_CONC_PF"]),
        "비고": "",
    })
st.subheader("📊 요약")
st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

if is_preview:
    st.success("✅ 미리보기 완료. 양식·수치를 확인하시고 위의 **🚀 PDF 전체 생성** 으로 본 작업을 시작하세요.")
st.subheader("📥 일괄 다운로드 / 인쇄" + (" (전체 생성 후 활성)" if is_preview else ""))
c1, c2, c3 = st.columns(3)
if os.path.exists(merged):
    with open(merged, "rb") as f:
        c1.download_button("📦 합본 PDF", data=f.read(),
            file_name=os.path.basename(merged), mime="application/pdf",
            use_container_width=True)
if (not is_preview) and os.path.exists(all_xlsx_zip):
    with open(all_xlsx_zip, "rb") as f:
        c2.download_button("📊 엑셀 모음 (ZIP)", data=f.read(),
            file_name=os.path.basename(all_xlsx_zip),
            mime="application/zip", use_container_width=True)
elif is_preview:
    c2.caption("미리보기 모드 — 전체 생성 시 활성화")
if printer_name:
    if c3.button("🖨 합본 인쇄", use_container_width=True):
        try:
            print_pdf(merged, printer_name=printer_name,
                      color=print_color, duplex=print_duplex, copies=print_copies)
            c3.success(f"✅ {printer_name} 로 전송 완료 ({print_copies}부)")
        except Exception as e:
            c3.error(f"인쇄 실패: {e}")
else:
    c3.caption("프린터 미감지 (Windows + 프린터 설치 필요)")

st.subheader("🔍 개별 검사성적서 — 미리보기 & 다운로드")

@st.cache_data(show_spinner=False)
def render_pdf_first_page(pdf_path: str, dpi: int, mtime: float):
    """mtime 을 캐시 키에 포함해, PDF 가 새로 생성되면 자동으로 재렌더링."""
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    pix = page.get_pixmap(dpi=dpi)
    return pix.tobytes("png")

for r in results:
    if not r.get("ok"):
        with st.expander(f"❌ {r['folder']} — 오류"):
            st.error(r.get("error", "처리 중 오류"))
        continue

    s = r["stats"]
    all_pass = all(s[k] == "PASS" for k in ("S1_SIZE_PF","S1_RND_PF","S2_POS_PF","S3_CONC_PF"))
    icon = "✅" if all_pass else "🚨"
    label = f"{icon}  {r['folder']}  —  Size {s['S1_SIZE_PF']} / Round {s['S1_RND_PF']} / Pos {s['S2_POS_PF']} / Conc {s['S3_CONC_PF']}"

    with st.expander(label, expanded=not all_pass):
        col_prev, col_btn = st.columns([3, 1])
        try:
            png = render_pdf_first_page(r["pdf"], preview_dpi, os.path.getmtime(r["pdf"]))
            col_prev.image(png, caption=f"{r['folder']}", use_container_width=True)
        except Exception as e:
            col_prev.warning(f"미리보기 실패: {e}")

        safe = r.get("safe") or r["folder"].replace("#", "no").replace(" ", "_")
        with open(r["pdf"], "rb") as f:
            col_btn.download_button("📥 PDF", data=f.read(),
                file_name=f"{safe}.pdf", mime="application/pdf",
                key=f"pdf_{r['folder']}", use_container_width=True)
        with open(r["xlsx"], "rb") as f:
            col_btn.download_button("📊 Excel", data=f.read(),
                file_name=f"{safe}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"xlsx_{r['folder']}", use_container_width=True)
        if printer_name:
            if col_btn.button("🖨 인쇄", key=f"print_{r['folder']}", use_container_width=True):
                try:
                    print_pdf(r["pdf"], printer_name=printer_name,
                              color=print_color, duplex=print_duplex, copies=print_copies)
                    col_btn.success("✅ 전송됨")
                except Exception as e:
                    col_btn.error(f"실패: {e}")

        col_btn.metric("Size MAX", f"{s['Size']['MAX']:.3f}")
        col_btn.metric("Roundness MAX", f"{s['Round']['MAX']:.4f}")
        col_btn.metric("Position MAX", f"{s['Pos']['MAX']:.3f}")
