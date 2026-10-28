"""Doc hoa don cho tung so don bao hiem.

Nguon tien, uu tien tu tren xuong:
  1. 'Hóa đơn sạch <ngay>.xlsx' — file trung gian do build_invoice_master.py ghep
     BK_HDVAT (tien) voi List hoa don (so don BH).
  2. BK_HDVAT*.xlsx + List hóa đơn.xlsx — tu ghep trong bo nho roi luu lai file (1)
     de lan sau khoi ghep lai.
  3. Chi co List hóa đơn.xlsx — luat cu, tien = tong 'Tien phi'+'Tien VAT' cua don.
     Voi don NHIEU KY PHI so nay bi thieu vi 1 hoa don xuat cho ca hai ky,
     vd BA-00005103 = 295.084.660 nhung List chi gan duoc 162.296.563.

Luat chon hoa don cho 1 so don (giu nguyen tu ban lam tay cua ke toan):
  gom theo NGAY, lay ngay MOI NHAT, cong het hoa don trong ngay do
  — 26/45/06/CHBB/CN00026 co 4 chung tu cung ngay 27/07 -> cong ca 4;
    26/45/06/CHBB/P000046 co 2 ngay khac nhau -> chi lay ngay moi nhat (xuat lai).
"""
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

import build_invoice_master as M
from build_debt_report import DIR_IN, DIR_OUT

TEN_FILE = "List hóa đơn.xlsx"
MAU_SACH, MAU_BK, MAU_LIST = "Hóa đơn sạch*.xlsx", "BK_HDVAT*.xlsx", "List hóa đơn*.xlsx"
SHEET_CAP = "Đơn ↔ Hóa đơn"

# vi tri cot trong file List hoa don (0-based) — dung cho luat cu
C_SO_CTU, C_NGAY_CTU, C_SO_DON, C_KY_HIEU, C_TIEN_PHI, C_TIEN_VAT = 1, 2, 3, 17, 26, 28


def _ngay(v):
    """Cat gio ra khoi 'Ngay lap HD'. BK ghi ca gio phat hanh nen 4 chung tu cung ngay
    27/07 cua 26/45/06/CHBB/CN00026 se thanh 4 moc khac nhau, luat 'lay ngay moi nhat'
    chi con cong duoc 2 chung tu -> hut 1,6 ty."""
    return v.date() if isinstance(v, datetime) else v


def _quet(thu_muc, mau):
    """File moi nhat khop mau, tim ca trong 2 thu muc con 'Output' va 'Input'."""
    ra = []
    for d in thu_muc:
        if not d:
            continue
        d = Path(d)
        for cho in (d, d / DIR_OUT, d / DIR_IN):
            ra += [p for p in cho.glob(mau) if p.is_file()]
    return max(ra, key=lambda p: p.stat().st_mtime) if ra else None


def tim_file(*thu_muc):
    """File hoa don de dung: ban sach neu con moi hon nguon, khong thi List hoa don."""
    sach = _quet(thu_muc, MAU_SACH)
    if _con_moi(sach, _quet(thu_muc, MAU_BK), _quet(thu_muc, MAU_LIST)):
        return sach
    return _quet(thu_muc, MAU_LIST) or sach


def _con_moi(sach, *nguon):
    """File sach chi dung duoc khi moi hon moi file nguon."""
    if not sach:
        return False
    t = sach.stat().st_mtime
    return all(t >= p.stat().st_mtime for p in nguon if p and Path(p).exists())


def _doc_sach(path, log=print):
    """Doc sheet 'Don <-> Hoa don' -> {so don: {ngay: {(ky hieu, so hd): tien}}}"""
    wb = load_workbook(path, read_only=True, data_only=True)
    if SHEET_CAP not in wb.sheetnames:
        wb.close()
        return None
    ws = wb[SHEET_CAP]
    it = ws.iter_rows(min_row=2, values_only=True)
    hdr = [str(v or "").split("\n")[0].strip() for v in next(it)]
    ix = {t: hdr.index(t) for t in ("Số đơn BH", "Ký hiệu", "Số hóa đơn", "Ngày lập HĐ",
                                    "Ngày ctừ (List)", "Phí+VAT của đơn", "Σ Phí+VAT",
                                    "Tổng tiền HĐ", "Số đơn")}
    idx = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    n = 0
    for r in it:
        don = str(r[ix["Số đơn BH"]] or "").strip()
        if not don:
            continue
        # cot cong thuc trong file khong co gia tri san -> tinh lai bang dung cong thuc do
        tien = M.phan_bo(r[ix["Phí+VAT của đơn"]] or 0, r[ix["Σ Phí+VAT"]] or 0,
                         r[ix["Tổng tiền HĐ"]], r[ix["Số đơn"]])
        ngay = _ngay(r[ix["Ngày lập HĐ"]] or r[ix["Ngày ctừ (List)"]])
        idx[don][ngay][(str(r[ix["Ký hiệu"]] or ""), str(r[ix["Số hóa đơn"]] or ""))] += tien
        n += 1
    wb.close()
    log(f"  {path.name}: {n:,} cặp đơn×hóa đơn, {len(idx):,} số đơn BH")
    return idx


def _tu_ghep(f_bk, f_list, out_dir, log=print):
    """Ghep thang tu BK + List, luu lai file sach roi tra ve index."""
    bk, ky = M.doc_bk(f_bk, log)
    li = M.doc_list(f_list, log)
    idx = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    for k, dons in li.items():
        b = bk.get(k)
        tong = sum(d["tien"] for d in dons.values())
        for don, d in dons.items():
            tien = M.phan_bo(d["tien"], tong, b["tong"] if b else None, len(dons))
            idx[don][_ngay((b["ngay"] if b else None) or d["ngay"])][k] += tien
    try:
        nhan = M.ky_thanh_nhan(ky)
        out = Path(out_dir) / f"Hóa đơn sạch {nhan}.xlsx"
        out.parent.mkdir(parents=True, exist_ok=True)
        M.xay(bk, li, ky, out, log)
    except Exception as e:                                   # noqa: BLE001
        log(f"  ! khong ghi duoc file hoa don sach ({e}), van chay tiep")
    return idx


def _doc_list_cu(path, log=print):
    """Luat cu: chi co List hoa don, tien = Tien phi + Tien VAT cua chinh don do."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    idx = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    n = 0
    for i, r in enumerate(ws.iter_rows(values_only=True), 1):
        if i == 1 or not r or len(r) <= C_TIEN_VAT:
            continue
        don = str(r[C_SO_DON] or "").strip()
        if not don:
            continue
        phi = r[C_TIEN_PHI] if isinstance(r[C_TIEN_PHI], (int, float)) else 0
        vat = r[C_TIEN_VAT] if isinstance(r[C_TIEN_VAT], (int, float)) else 0
        k = (str(r[C_KY_HIEU] or "").strip(), M.so_hd(r[C_SO_CTU]))
        idx[don][_ngay(r[C_NGAY_CTU])][k] += phi + vat
        n += 1
    wb.close()
    log(f"  {path.name}: {n:,} dòng, {len(idx):,} số đơn BH (tiền theo List)")
    return idx


def doc_hoa_don(path, log=print, thu_muc=(), f_bk=None):
    """-> {so don BH: {ngay: {(ky hieu, so hoa don): tien}}}"""
    path = Path(path)
    thu_muc = tuple(thu_muc) + (path.parent,)
    f_bk = Path(f_bk) if f_bk else None
    if path.match(MAU_SACH):
        idx = _doc_sach(path, log)
        if idx is not None:
            return idx
        log(f"  ! {path.name} khong co sheet '{SHEET_CAP}', doc nhu List hoa don")

    f_list = path if path.match(MAU_LIST) else _quet(thu_muc, MAU_LIST)
    f_bk = f_bk if (f_bk and f_bk.exists()) else _quet(thu_muc, MAU_BK)
    if not f_bk:
        log("  ! khong thay BK_HDVAT*.xlsx, tien hoa don lay theo List (don nhieu ky phi se thieu)")
        return _doc_list_cu(f_list or path, log)

    sach = _quet(thu_muc, MAU_SACH)
    if _con_moi(sach, f_bk, f_list):
        idx = _doc_sach(sach, log)
        if idx is not None:
            return idx
    if sach:
        log(f"  {sach.name} cu hon file nguon, ghep lai")
    out = next((Path(d) / DIR_OUT for d in thu_muc if d and (Path(d) / DIR_OUT).is_dir()),
               f_bk.parent / DIR_OUT)
    return _tu_ghep(f_bk, f_list, out, log)


def chon_hoa_don(idx, so_don):
    """-> (chuoi hien thi, tien, ngay, so lan xuat) hoac (None, 0, None, 0)."""
    d = idx.get(str(so_don or "").strip())
    if not d:
        return None, 0.0, None, 0
    ngay = max((k for k in d if k is not None), default=None)
    if ngay is None:
        return None, 0.0, None, 0
    hd = d[ngay]
    tien = sum(hd.values())
    sos = ", ".join(dict.fromkeys(k[1] if isinstance(k, tuple) else k for k in sorted(hd)))
    try:
        nhan = f"{sos} - {ngay:%d/%m/%Y}"
    except (TypeError, ValueError):
        nhan = sos
    return nhan, tien, ngay, len(d)
