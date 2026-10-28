"""Bat Excel tinh lai toan bo cong thuc roi doi chieu file ket qua voi 5 file raw.

    py -3.13 verify_report.py "<duong dan file ket qua>"
"""
import os
import re
import sys
import time
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_debt_report import TT_KHONG_NO, don_goc  # noqa: E402

ERRS = {"#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NULL!", "#NUM!"}
ROW_DATA = 12
ACC2SHEET = {"131112": "12", "131113": "13", "131114": "14", "131115": "15", "0051": "0051"}
C_TIEN_RAW, C_TIEN_OUT, C_QH_OUT, C_NHOM_OUT = 23, 22, 23, 34   # 0-based
C_SO_DON_OUT, C_GOC_OUT, C_TT_OUT = 4, 33, 35                   # sheet 0051: E / AH / AJ


def num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def kiem_dong_con(ws, log):
    """Dong con (ten dai ly, thut le) cong lai phai bang dong khach hang ngay tren."""
    nhom = []
    for row in ws.iter_rows():
        a = row[0]
        if a.value in (None, ""):
            continue
        if a.alignment is not None and (a.alignment.indent or 0) > 0:
            if nhom:
                nhom[-1][1].append(row)
        else:
            nhom.append((row, []))
    ok = True
    for cha, con in nhom:
        if not con:
            continue
        lech = [j for j in range(1, len(cha)) if isinstance(cha[j].value, (int, float))
                and abs(sum(num(c[j].value) for c in con) - cha[j].value) >= 0.01]
        ok &= not lech
        log(f"  {ws.title}: {str(cha[0].value)[:38]} = {len(con)} dong dai ly"
            + ("  OK" if not lech else "  <-- LECH cot " + ",".join(str(j + 1) for j in lech)))
    return ok


def recalc(path, log=print, lan=2):
    """Mo Excel tinh lai cong thuc. Tra ve False neu may khong co Excel.

    Thu lai 1 lan: chay nhieu tien trinh lien tiep hay gap 'Server execution failed'
    (-2146959355) do Excel chua kip khoi dong lai.
    """
    try:
        import win32com.client as win32
    except ImportError:
        log("  ! khong co pywin32, bo qua buoc tinh lai bang Excel")
        return False
    path = Path(path)
    tam = path.with_name("_recalc_" + path.name)       # khong bat dau bang "BC c" -> gui khong nham
    for i in range(lan):
        app, ban_sao = None, False
        try:
            app = win32.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            t0 = path.stat().st_mtime
            wb = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
            app.Application.CalculateFullRebuild()
            if not wb.ReadOnly:
                wb.Save()
            if wb.ReadOnly or path.stat().st_mtime == t0:
                # Save() khong ghi gi ma KHONG bao loi -> doc lai ra toan o trong, bao LECH oan.
                # Gap 14/09/2026: Excel mo moi file o che do chi doc, SaveAs .xlsx im lang bo qua
                # nhung SaveCopyAs van ghi duoc. Luu ban sao, dong Excel roi moi thay file goc.
                tam.unlink(missing_ok=True)
                wb.SaveCopyAs(str(tam))
                ban_sao = True
            wb.Close(SaveChanges=False)
        except Exception as e:                               # noqa: BLE001
            if i + 1 < lan:
                log(f"  ! Excel chua san sang ({e}), thu lai sau 5 giay")
                time.sleep(5)
                continue
            log(f"  ! khong goi duoc Excel ({e}), cong thuc se tu tinh khi mo file")
            return False
        finally:
            if app is not None:
                try:
                    app.Quit()
                except Exception:                            # noqa: BLE001
                    pass
        if ban_sao:
            if not tam.exists():
                log("  ! Excel khong luu duoc file (ca Save lan SaveCopyAs), cong thuc se tu "
                    "tinh khi mo file")
                return False
            for _ in range(10):                  # Excel vua thoat co the chua nha file
                try:
                    os.replace(tam, path)
                    break
                except PermissionError:
                    time.sleep(1)
            else:
                log(f"  ! khong ghi de duoc {path.name} (dang mo?), ban da tinh nam o {tam.name}")
                return False
        log(f"  Excel da tinh lai va luu: {path.name}" + (" (qua ban sao)" if ban_sao else ""))
        return True


def read_sheet(path, sheet, first_col=1):
    """Doc vung du lieu cua 1 sheet (tu ROW_DATA) -> list cac dong."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    out = []
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i < ROW_DATA:
            continue
        if row is None or len(row) <= first_col or row[first_col] in (None, ""):
            continue
        out.append(row)
    wb.close()
    return out


def read_raw_dir(src):
    raw = {}
    for f in sorted(Path(src).glob("BCPBHTT_*.xlsx")):
        wb = load_workbook(f, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        acc, rows = None, []
        for i, r in enumerate(ws.iter_rows(values_only=True), 1):
            if i == 6:
                m = re.search(r":\s*([0-9]{4,6})", " ".join(str(x) for x in r if x))
                acc = m.group(1) if m else None
            elif i > 13 and r and len(r) > 1 and r[1] not in (None, ""):
                rows.append(r)
        wb.close()
        if acc and (acc not in raw or len(rows) > len(raw[acc])):
            raw[acc] = rows
    return raw


def verify(path, src_dir=None, log=print):
    """src_dir = thu muc chua 5 file raw; bo trong thi tim canh file ket qua."""
    path = Path(path).resolve()
    log("Kiem tra file ket qua")
    # Khong goi duoc Excel thi cong thuc chua co gia tri, doc lai bang openpyxl se ra None
    # -> moi so tinh tu cot cong thuc deu bang 0 va bao LECH oan. Bo qua phan do, dung ket
    # luan la sai lech.
    da_tinh = recalc(path, log)

    bad = set()
    if da_tinh:
        wb = load_workbook(path, data_only=True)
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for v in row:
                    if isinstance(v, str) and v.strip() in ERRS:
                        bad.add(ws.title)
        wb.close()
        log(f"\nLoi cong thuc: {'KHONG CO' if not bad else sorted(bad)}")

    # bao cao chay le (chi 0051 / chi 13111) thi chi kiem phan tai khoan co trong file
    wb = load_workbook(path, read_only=True)
    co13, co51 = "BC 13111" in wb.sheetnames, "0051" in wb.sheetnames
    wb.close()
    raw = read_raw_dir(src_dir or path.parent)
    ok = True
    log("\n=== doi chieu tung sheet voi file raw ===")
    log(f"{'sheet':<10}{'dong raw':>10}{'dong out':>10}{'Σ con lai raw':>22}{'Σ con lai out':>22}")
    for acc, sh in ACC2SHEET.items():
        if (acc == "0051" and not co51) or (acc != "0051" and not co13):
            continue
        if acc not in raw:
            log(f"{sh:<10}  (khong tim thay file raw cua TK {acc})")
            ok = False
            continue
        d = read_sheet(path, sh)
        n_out, s_out = len(d), sum(num(r[C_TIEN_OUT]) for r in d)
        n_raw = len(raw[acc])
        s_raw = sum(num(r[C_TIEN_RAW]) for r in raw[acc])
        good = n_out == n_raw and abs(s_out - s_raw) < 0.01
        ok &= good
        log(f"{sh:<10}{n_raw:>10}{n_out:>10}{s_raw:>22,.2f}{s_out:>22,.2f}"
            f"{'' if good else '   <-- LECH'}")

    bc = read_sheet(path, "BC 13111") if co13 else []
    if co13:
        n_u = sum(len(raw[a]) for a in ("131112", "131113", "131114", "131115") if a in raw)
        s_u = sum(num(r[C_TIEN_RAW]) for a in ("131112", "131113", "131114", "131115")
                  if a in raw for r in raw[a])
        s_bc = sum(num(r[C_TIEN_OUT]) for r in bc)
        good = len(bc) == n_u and abs(s_bc - s_u) < 0.01
        ok &= good
        log(f"{'BC 13111':<10}{n_u:>10}{len(bc):>10}{s_u:>22,.2f}{s_bc:>22,.2f}"
            f"{'' if good else '   <-- LECH'}")

    if not da_tinh:
        log("\n! Chua goi duoc Excel nen KHONG kiem tra duoc phan cong thuc (3 khoi "
            "'Theo ma KH', dong KIEM TRA, o loi). So dong va so tien deu khop file raw;"
            "\n  mo file bang Excel de no tu tinh lai, hoac dong bot Excel dang chay roi "
            "bam Dung bao cao lan nua.")
        log("\n" + ("SO LIEU KHOP FILE RAW." if ok else "CO SAI LECH, xem cac dong danh dau."))
        return ok

    if co13:
        log("\n=== 3 khoi 'Theo ma KH' vs tinh lai tu BC 13111 ===")
    tinh = {}
    for r in bc:
        g = str(r[C_NHOM_OUT] or "").strip().upper()
        v, q = tinh.get(g, (0.0, 0.0))
        tinh[g] = (v + num(r[C_TIEN_OUT]), q + num(r[C_QH_OUT]))

    wb = load_workbook(path, data_only=True)
    rows = list(wb["Theo mã KH"].iter_rows(values_only=True)) if co13 else []
    log("\n=== dong con theo dai ly = dong khach hang ngay tren ===")
    for sn in ("Theo mã KH", "Quá hạn 13111", "Quá hạn 0051"):
        if sn in wb.sheetnames:
            ok &= kiem_dong_con(wb[sn], log)
    for sn in ("Quá hạn 13111", "Quá hạn 0051"):
        if sn not in wb.sheetnames:
            continue
        r = next((r for r in wb[sn].iter_rows(values_only=True)
                  if str(r[0] or "").startswith("KIỂM TRA")), None)
        v = [num(x) for x in r[1:]] if r else []
        good = r is not None and all(abs(x) < 0.01 for x in v)
        ok &= good
        log(f"  {sn}: dong KIEM TRA (phai = 0) " + ("OK" if good else f"<-- LECH {v}"))
    wb.close()

    log("\n=== sheet 0051: so don goc (bo SDBS) va tinh trang ===")
    s51 = read_sheet(path, "0051") if co51 else []
    lech = [r for r in s51 if str(r[C_GOC_OUT] or "").strip() != don_goc(r[C_SO_DON_OUT])]
    tong = defaultdict(float)
    for r in s51:
        tong[don_goc(r[C_SO_DON_OUT])] += num(r[C_TIEN_OUT])
    can = sum(1 for r in s51 if abs(tong[don_goc(r[C_SO_DON_OUT])]) < 0.005)
    co = sum(1 for r in s51 if r[C_TT_OUT] == TT_KHONG_NO)
    good = not lech and can == co
    ok &= good
    log(f"  cot so don goc khac cach tinh Python: {len(lech)} dong"
        + (f" (vd {lech[0][C_SO_DON_OUT]} -> {lech[0][C_GOC_OUT]})" if lech else ""))
    log(f"  dong '{TT_KHONG_NO}': file={co:,}  tinh lai={can:,}  {'OK' if good else '<-- LECH'}")
    titles = [i for i, r in enumerate(rows) if str(r[0] or "").startswith("CÔNG NỢ QUÁ HẠN")]
    totals = [i for i, r in enumerate(rows) if str(r[0] or "").strip() == "Tổng cộng"]
    for t, tr in zip(titles, totals):
        name = str(rows[t][0]).replace("CÔNG NỢ QUÁ HẠN ", "").strip()
        fv, fq = num(rows[tr][1]), num(rows[tr][2])
        cv, cq = tinh.get(name.upper(), (float("nan"), float("nan")))
        good = abs(fv - cv) < 0.01 and abs(fq - cq) < 0.01
        ok &= good
        log(f"  {name:<18} file={fv:>20,.2f} /{fq:>16,.2f}   tinh lai={cv:>20,.2f} /{cq:>16,.2f}"
            f"  {'OK' if good else '<-- LECH'}")
    chk = [i for i, r in enumerate(rows) if str(r[0] or "").startswith("KIỂM TRA")]
    if chk:
        v = [num(rows[chk[0]][j]) for j in (1, 2)]
        ok &= all(abs(x) < 0.01 for x in v)
        log(f"  dong KIEM TRA (phai = 0): {v[0]:,.2f} / {v[1]:,.2f}")

    ok = ok and not bad
    log("\n" + ("TAT CA KHOP." if ok else "CO SAI LECH, xem cac dong danh dau."))
    return ok


if __name__ == "__main__":
    sys.exit(0 if verify(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None) else 1)
