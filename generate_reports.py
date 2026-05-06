"""
검사성적서 자동 생성 스크립트
============================

사용법
------
1) 동일 폴더의 metadata.csv 를 채운다
   (folder, product_name, processing_date, mgmt_no, equip_no, tool_sn,
    insp_name, conf_name, appr_name)
2) python generate_reports.py [--inspection-date 2026-05-04]
3) 결과
   - 각 폴더별 PDF: ./out/<folder>.pdf
   - 합본 PDF:      ./out/_합본_<inspection-date>.pdf
   - 채워진 엑셀 사본: ./out/<folder>.xlsx
"""

import argparse
import csv
import datetime as dt
import os
import re
import shutil
import statistics
import subprocess
import sys

import xlrd
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

TEMPLATE = os.path.join(HERE, "template.xlsx")
META_CSV = os.path.join(HERE, "metadata.csv")
OUT_DIR  = os.path.join(HERE, "out")


def find_soffice():
    env = os.environ.get("SOFFICE")
    if env and os.path.exists(env):
        return env
    for cmd in ("soffice", "libreoffice"):
        for p in os.environ.get("PATH", "").split(os.pathsep):
            full = os.path.join(p, cmd + (".exe" if os.name == "nt" else ""))
            if os.path.exists(full):
                return full
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    raise FileNotFoundError(
        "LibreOffice (soffice) 를 찾을 수 없습니다.\n"
        "  - Windows: https://www.libreoffice.org/download/download/ 에서 설치 후 재실행\n"
        "  - 환경변수 SOFFICE 에 soffice 실행 파일 경로를 지정해도 됩니다."
    )


SOFFICE = None


def load_xls_rows(path):
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    rows = []
    for r in range(1, sh.nrows):
        cid  = sh.cell_value(r, 1)
        prop = sh.cell_value(r, 2)
        nom  = sh.cell_value(r, 3)
        meas = sh.cell_value(r, 6)
        if isinstance(meas, (int, float)):
            rows.append((cid, prop, nom, meas))
    return rows


def compute_stats(folder_path):
    a_files = sorted(f for f in os.listdir(folder_path) if f.lower().endswith("a.xls"))
    b_files = sorted(f for f in os.listdir(folder_path) if f.lower().endswith("b.xls"))
    if not a_files or not b_files:
        raise FileNotFoundError(f"A.xls/B.xls 누락: {folder_path}")

    rows = []
    for fn in a_files + b_files:
        rows += load_xls_rows(os.path.join(folder_path, fn))

    def vals(prop):
        return [r[3] for r in rows if r[1] == prop]

    D    = vals("D")
    Circ = vals("Circularity")
    Pos  = vals("TRPOS")
    Conc = vals("Concentricity")

    if not D or not Circ or not Pos:
        raise ValueError(f"항목 누락: {folder_path}")

    def stat(arr):
        return dict(
            MIN=min(arr), MAX=max(arr),
            AVG=statistics.mean(arr),
            STDEV=statistics.stdev(arr) if len(arr) > 1 else 0.0,
        )

    out = {
        "Size":  stat(D),
        "Round": stat(Circ),
        "Pos":   stat(Pos),
        "Conc":  Conc[:8] + [None] * max(0, 8 - len(Conc)),
        "_count": len(D),
    }
    out["S1_SIZE_PF"] = "PASS" if (0.44 <= out["Size"]["MIN"] and out["Size"]["MAX"] <= 0.46) else "FAIL"
    out["S1_RND_PF"]  = "PASS" if out["Round"]["MAX"] < 0.02 else "FAIL"
    out["S2_POS_PF"]  = "PASS" if out["Pos"]["MAX"]   < 0.10 else "FAIL"
    out["S3_CONC_PF"] = "PASS" if all((c is not None and c < 0.10) for c in out["Conc"][:8]) else "FAIL"
    return out


def fmt_size(v):  return "" if v is None else f"{v:.3f}"
def fmt_round(v): return "" if v is None else f"{v:.4f}"
def fmt_pos(v):   return "" if v is None else f"{v:.3f}"
def fmt_conc(v):  return "" if v is None else f"{v:.3f}"


def fill_template(template_path, out_xlsx, stats, meta):
    shutil.copyfile(template_path, out_xlsx)
    wb = load_workbook(out_xlsx)
    ws = wb.active

    repl = {
        "{{PRODUCT_NAME}}":  meta["product_name"],
        "{{PROCESS_DATE}}":  meta["processing_date"],
        "{{MGMT_NO}}":       meta["mgmt_no"],
        "{{INSPECT_DATE}}":  meta["inspection_date"],
        "{{EQUIP_NO}}":      meta["equip_no"],
        "{{TOOL_SN}}":       meta["tool_sn"],
        "{{S1_SIZE_MIN}}":   fmt_size(stats["Size"]["MIN"]),
        "{{S1_SIZE_MAX}}":   fmt_size(stats["Size"]["MAX"]),
        "{{S1_SIZE_AVG}}":   fmt_size(stats["Size"]["AVG"]),
        "{{S1_SIZE_STDEV}}": fmt_size(stats["Size"]["STDEV"]),
        "{{S1_SIZE_PF}}":    stats["S1_SIZE_PF"],
        "{{S1_RND_MIN}}":    fmt_round(stats["Round"]["MIN"]),
        "{{S1_RND_MAX}}":    fmt_round(stats["Round"]["MAX"]),
        "{{S1_RND_AVG}}":    fmt_round(stats["Round"]["AVG"]),
        "{{S1_RND_STDEV}}":  fmt_round(stats["Round"]["STDEV"]),
        "{{S1_RND_PF}}":     stats["S1_RND_PF"],
        "{{S2_POS_MIN}}":    fmt_pos(stats["Pos"]["MIN"]),
        "{{S2_POS_MAX}}":    fmt_pos(stats["Pos"]["MAX"]),
        "{{S2_POS_AVG}}":    fmt_pos(stats["Pos"]["AVG"]),
        "{{S2_POS_STDEV}}":  fmt_pos(stats["Pos"]["STDEV"]),
        "{{S2_POS_PF}}":     stats["S2_POS_PF"],
        **{f"{{{{C{i+1}}}}}": fmt_conc(stats["Conc"][i]) for i in range(8)},
        "{{S3_CONC_PF}}":    stats["S3_CONC_PF"],
        "{{INSP_NAME}}":     meta["insp_name"],
        "{{CONF_NAME}}":     meta["conf_name"],
        "{{APPR_NAME}}":     meta["appr_name"],
    }

    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str):
                new_v = v
                for k, rep in repl.items():
                    if k in new_v:
                        new_v = new_v.replace(k, rep)
                if new_v != v:
                    cell.value = new_v

    wb.save(out_xlsx)


def _xlsx_to_pdf_via_libreoffice(xlsx_path, out_dir):
    """LibreOffice (soffice) 로 PDF 변환."""
    global SOFFICE
    if SOFFICE is None:
        SOFFICE = find_soffice()
    cmd = [SOFFICE, "--headless", "--convert-to", "pdf",
           "--outdir", out_dir, xlsx_path]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(f"LibreOffice 변환 실패: {r.stderr}")
    base = os.path.splitext(os.path.basename(xlsx_path))[0]
    return os.path.join(out_dir, base + ".pdf")


def _xlsx_to_pdf_via_excel(xlsx_path, out_dir):
    """MS Excel (win32com) 로 PDF 변환 — Windows 전용."""
    import pythoncom
    import win32com.client as win32
    base = os.path.splitext(os.path.basename(xlsx_path))[0]
    pdf_out = os.path.abspath(os.path.join(out_dir, base + ".pdf"))

    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path),
                                   ReadOnly=True, UpdateLinks=0)
        # 0 = xlTypePDF
        wb.ExportAsFixedFormat(0, pdf_out)
        wb.Close(SaveChanges=False)
        return pdf_out
    finally:
        try: excel.Quit()
        except Exception: pass
        pythoncom.CoUninitialize()


def xlsx_to_pdf(xlsx_path, out_dir):
    """xlsx → pdf. LibreOffice 우선, 없으면 MS Excel(Windows) fallback."""
    # 1) LibreOffice 시도
    try:
        return _xlsx_to_pdf_via_libreoffice(xlsx_path, out_dir)
    except FileNotFoundError:
        pass  # LibreOffice 없음 → Excel 시도
    except Exception as e_lo:
        # LibreOffice 가 있는데 변환 실패 → 그대로 알림
        if os.name != "nt":
            raise
        last_lo_err = str(e_lo)
    else:
        last_lo_err = None

    # 2) Windows + Excel fallback
    if os.name == "nt":
        try:
            return _xlsx_to_pdf_via_excel(xlsx_path, out_dir)
        except ImportError:
            raise RuntimeError(
                "PDF 변환 엔진을 찾을 수 없습니다.\n"
                "  옵션 A: LibreOffice 설치  (https://www.libreoffice.org/download/download/)\n"
                "  옵션 B: pip install pywin32  (MS Excel 이 설치돼 있으면 자동으로 사용)"
            )
        except Exception as e_xl:
            msg = f"Excel 변환 실패: {e_xl}"
            if last_lo_err:
                msg = f"LibreOffice 변환 실패: {last_lo_err}\n{msg}"
            raise RuntimeError(msg)

    # 3) 비-Windows 이고 LibreOffice 없음
    raise FileNotFoundError(
        "LibreOffice (soffice) 를 찾을 수 없습니다. 설치 후 다시 실행하세요."
    )


def merge_pdfs(pdfs, out_path):
    try:
        from pypdf import PdfWriter
    except ImportError:
        from PyPDF2 import PdfWriter
    writer = PdfWriter()
    for p in pdfs:
        writer.append(p)
    with open(out_path, "wb") as f:
        writer.write(f)




# ─────────────────── 사내 프린터 출력 (Windows 전용) ───────────────────
def list_printers():
    """Windows 에 등록된 프린터 목록 반환. 비-Windows 면 빈 리스트."""
    if os.name != "nt":
        return []
    try:
        import win32print
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        return [p[2] for p in win32print.EnumPrinters(flags)]
    except Exception:
        return []


def get_default_printer():
    if os.name != "nt":
        return None
    try:
        import win32print
        return win32print.GetDefaultPrinter()
    except Exception:
        return None


def print_pdf(pdf_path, printer_name=None, color=True, duplex=False, copies=1):
    """PDF 를 지정 프린터로 직접 출력 (Windows 전용).
    - color: True=컬러, False=흑백
    - duplex: True=양면(긴쪽), False=단면
    - copies: 부수
    """
    if os.name != "nt":
        raise RuntimeError("프린터 출력은 Windows 에서만 지원됩니다.")

    import time
    import win32api
    import win32print

    if not printer_name:
        printer_name = win32print.GetDefaultPrinter()

    # 프린터 DEVMODE 수정 (color / duplex / copies)
    handle = win32print.OpenPrinter(printer_name, {"DesiredAccess": win32print.PRINTER_ALL_ACCESS})
    props = win32print.GetPrinter(handle, 2)
    devmode = props["pDevMode"]
    orig = (devmode.Color, devmode.Duplex, devmode.Copies)

    devmode.Color   = 2 if color  else 1   # 2=DMCOLOR_COLOR, 1=DMCOLOR_MONOCHROME
    devmode.Duplex  = 2 if duplex else 1   # 2=long edge, 1=simplex
    devmode.Copies  = max(1, int(copies))
    props["pDevMode"] = devmode
    win32print.SetPrinter(handle, 2, props, 0)

    try:
        # ShellExecute "printto" — PDF 뷰어가 받아 처리 (Edge / Adobe 등)
        rc = win32api.ShellExecute(0, "printto", pdf_path, f'"{printer_name}"', ".", 0)
        if rc <= 32:
            raise RuntimeError(f"인쇄 시작 실패 (코드 {rc}). PDF 뷰어가 설치되어 있는지 확인.")
        # 작업이 큐에 들어갈 시간 확보
        time.sleep(2)
    finally:
        # 원래 설정 복원
        devmode.Color, devmode.Duplex, devmode.Copies = orig
        props["pDevMode"] = devmode
        try:
            win32print.SetPrinter(handle, 2, props, 0)
        except Exception:
            pass
        win32print.ClosePrinter(handle)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--inspection-date", default=None)
    parser.add_argument("--data-folder", default=None)
    parser.add_argument("--folder", default=None)
    args = parser.parse_args()

    if not os.path.exists(META_CSV):
        print(f"ERROR: {META_CSV} 가 없습니다.")
        sys.exit(1)

    with open(META_CSV, encoding="utf-8-sig") as f:
        meta_list = list(csv.DictReader(f))

    if args.folder:
        meta_list = [m for m in meta_list if m["folder"] == args.folder]

    insp_date = args.inspection_date
    if insp_date:
        for m in meta_list:
            m["inspection_date"] = insp_date.replace("-", ".")

    data_root = args.data_folder
    if not data_root and meta_list:
        d = meta_list[0]["inspection_date"]
        d_norm = re.sub(r"[.\-]", "-", d)
        try:
            parsed = dt.datetime.strptime(d_norm, "%Y-%m-%d")
        except ValueError:
            parts = d_norm.split("-")
            parsed = dt.datetime(int(parts[0]), int(parts[1]), int(parts[2]))
        data_root = os.path.join(ROOT, parsed.strftime("%Y-%m-%d"))
    print(f"데이터 루트: {data_root}")

    os.makedirs(OUT_DIR, exist_ok=True)
    pdfs = []

    for m in meta_list:
        folder = m["folder"]
        folder_path = os.path.join(data_root, folder)
        if not os.path.isdir(folder_path):
            print(f"  [skip] 폴더 없음: {folder_path}")
            continue
        print(f"\n▶ {folder}")
        stats = compute_stats(folder_path)
        print(f"  · 홀 갯수: {stats['_count']}")
        print(f"  · 판정: Size={stats['S1_SIZE_PF']} Round={stats['S1_RND_PF']} "
              f"Pos={stats['S2_POS_PF']} Conc={stats['S3_CONC_PF']}")

        safe_name = folder.replace("#", "no").replace(" ", "_")
        xlsx_out = os.path.join(OUT_DIR, f"{safe_name}.xlsx")
        fill_template(TEMPLATE, xlsx_out, stats, m)
        pdf_out = xlsx_to_pdf(xlsx_out, OUT_DIR)
        pdfs.append(pdf_out)
        print(f"  → {pdf_out}")

    for fn in os.listdir(OUT_DIR):
        if fn.startswith(".~lock") or fn.endswith(".tmp"):
            try:
                os.remove(os.path.join(OUT_DIR, fn))
            except OSError:
                pass

    if len(pdfs) > 1:
        date_tag = meta_list[0]["inspection_date"].replace(".", "-")
        merged = os.path.join(OUT_DIR, f"_합본_{date_tag}.pdf")
        merge_pdfs(pdfs, merged)
        print(f"\n✅ 합본 PDF → {merged}")
    elif pdfs:
        print(f"\n✅ {len(pdfs)} 개 PDF 생성 완료")


if __name__ == "__main__":
    main()
