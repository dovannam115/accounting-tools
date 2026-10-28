"""Bang ke chi tiet dinh kem ban xac nhan cong no gui P.KD.

Moi phong 2 sheet — "<Phong> 13111" va "<Phong> 0051" — la phan chi tiet cua dong
"Hoa don theo bang ke dinh kem" tren ban xac nhan. Luat do chi Ngoc chot:
  - lay tu cot STT den cot cong no qua han CUOI CUNG CON SO (cac bucket rong thi cat bo),
  - sheet 0051 chi giu dong co Check + - khac 0 (cac so don GOC triet tieu nhau thi bo han,
    ke ca SDBS /E01, /E02... cua cung don goc; loc o build_all cho ca ban xac nhan),
  - sheet nao toan VND thi bo 3 cot nguyen te cho chu in ra do nho.
Tong tien cua bang ke phai bang tong tren ban xac nhan — kiem tra tung o (tai khoan x ky).
"""
import sys
from collections import defaultdict
from datetime import date, datetime
from math import ceil
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.properties import PageSetupProperties

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_debt_report import (BASE24, BOX, BUCKET_BOUNDS, BUCKETS,  # noqa: E402
                               F_DATA, F_HEAD, F_TITLE, F_TOTAL, FILL_HEAD, NUMFMT,
                               TNR, WIDTH, don_goc, norm)

# vi tri cot trong sheet bao cao (0-based) — 24 cot dau giong nhau o BC 13111 va 0051
C_SO_DON, C_PHONG, C_NGAY_CAP, C_NGAY_TP, C_TTE, C_CON, C_QH = 4, 7, 13, 14, 15, 22, 23
N_BASE = 24                       # so cot dung chung (BASE24)
I_NTE = (17, 18, 19)              # Nguyen te phai thu / da thu / con lai
I_TIEN = (20, 21, 22, 23)         # So tien phai thu / da thu / con lai / Tong qua han
# chi Ngoc bo (17/08/2026): 3 cot ma — da co ten ben canh roi; rieng 0051 bo them
# 'So ngay qua han' (chinh la cot N_BASE cua sheet 0051)
I_BO_MA = (1, 3, 6)               # Ma khach hang / Ma CBCNV / Ma Phong ban

R_TITLE, R_HEAD, R_DATA = 1, 5, 6
F_SUB = Font(name=TNR, size=9, bold=True)


def _ky(r, gom_nam):
    ng = r[C_NGAY_CAP]
    if not isinstance(ng, (datetime, date)):
        return (0, 0)
    return (ng.year, 0) if gom_nam else (ng.year, ng.month)


def _ky_chu(ky):
    if not ky[0]:
        return "Các đơn BH không xác định ngày cấp đơn"
    return (f"Các đơn BH cấp tháng {ky[1]:02d}/{ky[0]}" if ky[1]
            else f"Các đơn BH cấp năm {ky[0]}")


def _phang(r, tk, ngay):
    """Tra ve list gia tri theo dung thu tu cot cua sheet nguon.

    Sheet 0051 co 'So ngay qua han' + 6 bucket la CONG THUC trong file bao cao — doc
    lai gia tri cache thi phu thuoc Excel co tinh lai hay khong, nen tinh thang o day
    y het cong thuc: so ngay = ngay bao cao - Ngay thu phi, bucket = Tong qua han
    roi vao khoang nao."""
    vals = list(r[:N_BASE])
    if tk == "13111":
        return vals + [r[N_BASE + k] for k in range(6)]
    ngay_tp = r[C_NGAY_TP]
    if isinstance(ngay_tp, datetime):
        ngay_tp = ngay_tp.date()
    so_ngay = (ngay - ngay_tp).days if isinstance(ngay_tp, date) else None
    qh = r[C_QH] if isinstance(r[C_QH], (int, float)) else 0
    buckets = []
    for lo, hi in BUCKET_BOUNDS:
        trong = so_ngay is not None and so_ngay >= lo and (hi is None or so_ngay < hi)
        buckets.append(qh if trong else 0)
    return vals + [so_ngay] + buckets


def _cot_dung(recs, tk):
    """Chon cot se in: cat sau bucket cuoi cung con so, bo 3 cot nguyen te neu toan VND.
    -> (danh sach chi so cot, danh sach tieu de, danh sach kieu)"""
    headers = [h for h, _k, _t in BASE24] + ([] if tk == "13111" else ["Số ngày quá hạn"]) \
        + list(BUCKETS)
    kinds = [t for _h, _k, t in BASE24] + ([] if tk == "13111" else ["int"]) + ["money"] * 6
    b0 = len(kinds) - 6
    cuoi = -1
    for k in range(6):
        if any(isinstance(v[b0 + k], (int, float)) and abs(v[b0 + k]) >= 0.005 for v in recs):
            cuoi = k
    bo = set(I_BO_MA) if tk == "13111" else set(I_BO_MA) | {N_BASE}
    if all(str(v[C_TTE] or "").strip().upper() in ("", "VND") for v in recs):
        bo |= set(I_NTE)
    idx = [i for i in list(range(b0)) + [b0 + k for k in range(cuoi + 1)] if i not in bo]
    return idx, headers, kinds


def _do_rong(recs, idx, kinds):
    """Be rong theo NOI DUNG THAT, khong de nguyen 22 nhu sheet bao cao. Bang ke in ep vua
    be rong 1 trang nen thua bao nhieu la chu in ra nho bay nhieu — vd PTSP 13111 co 21 cot
    ma 11 cot text de 22 thi in ra chi con ~45%."""
    rong = []
    for i in idx:
        k = kinds[i]
        if k == "money":
            m = max((abs(v[i]) for v in recs if isinstance(v[i], (int, float))), default=0)
            rong.append(min(20, max(12, len(f"{m:,.2f}") + 3)))
        elif k == "text":
            m = max((len(str(v[i])) for v in recs if v[i] is not None), default=0)
            rong.append(min(28, max(10, m + 2)))
        else:
            rong.append(WIDTH[k] if k != "int" else 7)
    return rong


def _trang_in(ws, ncol, ten):
    """Bang ke la de IN — ngang, ep vua be rong 1 trang, lap lai dong tieu de."""
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_title_rows = f"{R_HEAD}:{R_HEAD}"
    ws.print_area = f"A1:{L(ncol)}{ws.max_row}"
    for k, v in dict(left=0.25, right=0.25, top=0.4, bottom=0.4, header=0.2, footer=0.2).items():
        setattr(ws.page_margins, k, v)
    ws.oddFooter.left.text = ten
    ws.oddFooter.right.text = "Trang &P/&N"
    ws.oddFooter.left.size = ws.oddFooter.right.size = 8


def _ghi_sheet(wb, ten_sheet, ten_phong, tk, ngay, recs, gom_nam):
    """Ghi 1 sheet bang ke -> {ky: tong con lai} de doi chieu voi ban xac nhan."""
    idx, headers, kinds = _cot_dung(recs, tk)
    ncol = len(idx)
    ws = wb.create_sheet(ten_sheet)
    ws.sheet_view.showGridLines = False

    ws.cell(R_TITLE, 1, "BẢNG KÊ CHI TIẾT CÔNG NỢ PHÍ BẢO HIỂM").font = F_TITLE
    ws.cell(R_TITLE + 1, 1, f"{ten_phong}, Tài khoản {tk}").font = Font(name=TNR, size=11, bold=True)
    ws.cell(R_TITLE + 2, 1, f"Số liệu đến ngày {ngay:%d/%m/%Y}").font = Font(name=TNR, size=10)
    c = ws.cell(R_TITLE + 2, ncol, "Đơn vị tính: đồng")
    c.font = Font(name=TNR, size=10, italic=True)
    c.alignment = Alignment(horizontal="right")

    rong, cao = _do_rong(recs, idx, kinds), 30
    for j, i in enumerate(idx, 1):
        c = ws.cell(R_HEAD, j, headers[i])
        c.font, c.fill, c.border = F_HEAD, FILL_HEAD, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[L(j)].width = rong[j - 1]
        # cot hep lai thi tieu de phai xuong nhieu dong hon — tinh chieu cao cho khoi cut chu
        cao = max(cao, ceil(len(headers[i]) / max(6, rong[j - 1] * 10.8 / 9)) * 9 * 1.45)
    ws.row_dimensions[R_HEAD].height = round(cao, 1)

    # cot duoc cong: so tien + bucket qua han. KHONG cong cot nguyen te vi mot sheet
    # co the lan nhieu loai tien.
    b0 = len(kinds) - 6
    cong = {j for j, i in enumerate(idx, 1) if i in I_TIEN or i >= b0}

    recs = sorted(recs, key=lambda v: (_ky(v, gom_nam),
                                       v[C_NGAY_CAP] if isinstance(v[C_NGAY_CAP], (datetime, date))
                                       else date(1900, 1, 1),
                                       str(v[C_SO_DON] or "")))
    r, stt, dong_cong, tong_ky = R_DATA, 0, [], {}
    for ky, nhom in _theo_ky(recs, gom_nam):
        dau = r
        for v in nhom:
            stt += 1
            for j, i in enumerate(idx, 1):
                c = ws.cell(r, j)
                c.font, c.number_format = F_DATA, NUMFMT[kinds[i]]
                c.value = stt if i == 0 else v[i]
            r += 1
        lb = _ky_chu(ky)
        ws.cell(r, 1, "Cộng " + lb[0].lower() + lb[1:]).font = F_SUB
        for j in cong:
            c = ws.cell(r, j, f"=SUM({L(j)}{dau}:{L(j)}{r - 1})")
            c.font, c.number_format, c.border = F_SUB, NUMFMT["money"], BOX
        tong_ky[ky] = sum(v[C_CON] for v in nhom if isinstance(v[C_CON], (int, float)))
        dong_cong.append(r)
        r += 2

    ws.cell(r, 1, "TỔNG CỘNG").font = F_TOTAL
    for j in cong:
        c = ws.cell(r, j, "=" + "+".join(f"{L(j)}{d}" for d in dong_cong))
        c.font, c.number_format, c.border = F_TOTAL, NUMFMT["money"], BOX
    ws.freeze_panes = f"A{R_DATA}"
    ws.auto_filter.ref = f"A{R_HEAD}:{L(ncol)}{r - 1}"
    _trang_in(ws, ncol, f"{ten_phong}, TK {tk}, đến {ngay:%d/%m/%Y}")
    return tong_ky, len(recs)


def _theo_ky(recs, gom_nam):
    nhom = defaultdict(list)
    for v in recs:
        nhom[_ky(v, gom_nam)].append(v)
    return sorted(nhom.items())


def loc_0051(rows):
    """Giu dong co Check + - khac 0. Check + - = SUMIF theo So don BH GOC (bo SDBS) cua cot
    'So tien con lai phai thu' — cap dong triet tieu nhau (vd P000650/E04 va /E05, hay don goc
    va /E01 huy no) cong ra 0 nen khong con la cong no, in vao bang ke chi lam P.KD roi."""
    tong = defaultdict(float)
    for r in rows:
        if isinstance(r[C_CON], (int, float)):
            tong[don_goc(r[C_SO_DON])] += r[C_CON]
    giu = [r for r in rows if abs(tong[don_goc(r[C_SO_DON])]) >= 0.005]
    return giu, len(rows) - len(giu), sum(1 for v in tong.values() if abs(v) < 0.005)


def build_annex(bang, phong, gom, thu_tu, ngay, out_dir, log=print):
    """bang  : {"BC 13111": rows, "0051": rows da qua loc_0051} doc tu file bao cao
       phong : config sheet 'Phong KD'
       gom   : {sheet: {(tk, ky): tien}} do build_pkd tra ve — de doi chieu tong
       thu_tu: danh sach sheet da ghi trong ban xac nhan, giu dung thu tu do"""
    log_rows = []
    ten_theo_sheet, gom_nam = {}, {}
    theo_sheet = defaultdict(lambda: {"13111": [], "0051": []})
    for tk, rows in (("13111", bang.get("BC 13111", [])), ("0051", bang.get("0051", []))):
        for r in rows:
            p = phong.get(norm(r[C_PHONG]))
            if not p or not p["sheet"]:
                continue
            theo_sheet[p["sheet"]][tk].append(r)
            ten_theo_sheet[p["sheet"]] = p["ten"] or p["pias"]
            gom_nam[p["sheet"]] = p.get("gom", "").startswith("n")

    wb = Workbook()
    wb.remove(wb.active)
    thu_tu = [s for s in thu_tu if s.strip() in theo_sheet] + \
             [s for s in theo_sheet if s not in {t.strip() for t in thu_tu}]
    n_sheet = 0
    for sn in thu_tu:
        sn = sn.strip()
        for tk in ("13111", "0051"):
            rows = theo_sheet[sn][tk]
            if not rows:
                continue
            recs = [_phang(r, tk, ngay) for r in rows]
            tong_ky, n = _ghi_sheet(wb, f"{sn} {tk}", ten_theo_sheet[sn], tk, ngay,
                                    recs, gom_nam[sn])
            n_sheet += 1
            # tong bang ke phai trung tung o cua ban xac nhan
            xn = {k: v for (t, k), v in gom.get(sn, {}).items() if t == tk}
            for ky in sorted(set(tong_ky) | set(xn)):
                lech = tong_ky.get(ky, 0.0) - xn.get(ky, 0.0)
                if abs(lech) >= 0.005:
                    log_rows.append((f"Bảng kê lệch bản xác nhận, {_ky_chu(ky)}", "", "",
                                     f"{sn} {tk}", "bảng kê P.KD", ten_theo_sheet[sn],
                                     xn.get(ky, 0.0), tong_ky.get(ky, 0.0)))
            log(f"    {sn} {tk}: {n:,} dong, {sum(tong_ky.values()):,.2f}")

    duoi = "" if len(bang) == 2 else " (%s)" % ("0051" if "0051" in bang else "13111")
    out = Path(out_dir) / f"BẢNG KÊ ĐÍNH KÈM BẢN XÁC NHẬN P.KD T{ngay:%m.%Y}{duoi}.xlsx"
    wb.save(out)
    log(f"  bang ke dinh kem: {n_sheet} sheet -> {out.name}")
    return out, log_rows
