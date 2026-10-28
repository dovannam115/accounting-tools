# -*- coding: utf-8 -*-
"""Đối trừ với sổ doanh thu 5111, kèm phần áp bóc tách hoa hồng và phí dịch vụ.

Hai đầu vào:
  1. Sổ 5111 kết xuất từ PIAS (1 sheet, mỗi dòng một bút toán doanh thu).
  2. File chi tiết đơn của đối tác (nhiều sheet: TNDS XM, TNDS OTO, TNCN, TNCN toàn diện ...).

Hai đầu ra:
  1. Detail_Applied_<ngày>.xlsx  chi tiết đã áp: điền số hóa đơn và mã kế toán lấy từ sổ,
     tính hoa hồng đại lý và phí dịch vụ theo biểu tỷ lệ, kèm sheet BBDS và sheet Tỷ lệ.
  2. Offset_5111_<ngày>.xlsx  bảng đối trừ: theo hóa đơn, theo đơn, thiếu trong sổ, thừa trong sổ.

Quy ước đã kiểm bằng dữ liệu thật (kỳ T08/2026, sổ 5111 xuất ngày 28/09/2026):
  - Khoá mức hóa đơn: cột "Hóa đơn" của file chi tiết ghép với "Số ctừ" của sổ. Số ctừ bị lặp
    giữa các ký hiệu hóa đơn nên phải lọc sổ theo mã đại lý; lọc rồi thì 31/33 hóa đơn lệch 0.
  - Khoá mức đơn: "Số đơn BH" ghép với "Mã sp". Đơn xe máy tập thể trong sổ chỉ có một dòng cho
    mỗi mã sản phẩm, còn file chi tiết có một dòng cho mỗi xe, nên chỉ so được tổng theo khoá
    chứ không so được từng dòng.
  - Mọi cột đều dò theo TÊN tiêu đề, không neo theo chỉ số cột. Chữ đ và Đ không tách dấu được
    bằng Unicode NFD nên phải đổi sang d và D trước, nếu không "Số đơn BH" thành "sonbh".
  - Hoa hồng và phí dịch vụ làm tròn từng dòng rồi RẢI phần dư để tổng theo mã sản phẩm khớp
    đúng tổng tính một lần (kế toán đang làm tay bằng cách cộng trừ 1 ở vài dòng).

Chạy ngoài app:  py -3.13 offset_5111.py <sổ 5111.xlsx> <file chi tiết.xlsx> [thư mục ra]
"""
import datetime as dt
import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from decimal import Decimal, ROUND_HALF_UP

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.environ.get("OFFSET_ROOT") or os.path.dirname(HERE)
RATES_NAME = "Offset_Rates.xlsx"
BBDS_NAME = "BBDS_Template.xlsx"   # sheet BBDS: dung lai dung file mau cua ke toan
RATE_SHEET = "Tỷ lệ"       # sheet ty le trong file giao, cong thuc tra sang day

MAX_HEADER_ROW = 12                      # tiêu đề nằm trong mấy dòng đầu
FILL_HEAD = PatternFill("solid", start_color="EEF2F7")
FILL_BAD = PatternFill("solid", start_color="FDE9E9")
FILL_WARN = PatternFill("solid", start_color="FFF6E0")
THIN = Side(style="thin", color="D0D0D0")
RED = Font(name="Arial", size=10, color="C00000")


class ToolError(Exception):
    """Lỗi do đầu vào, trang web hiện thẳng câu này cho người dùng."""


# ----------------------------------------------------------------- tiện ích chung

def plain(s):
    """Chuẩn hoá tên cột: bỏ dấu, bỏ ký tự lạ. Phải đổi đ/Đ trước vì NFD không tách được."""
    s = str(s or "").replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]", "", s.lower())


def dec(v):
    """Số tiền về Decimal, chuỗi rỗng hay chữ thì về None."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float, Decimal)):
        return Decimal(str(v))
    s = str(v).strip().replace(" ", "").replace(",", "")
    if not s or s in {"-", "."}:
        return None
    try:
        return Decimal(s)
    except Exception:
        return None


def r0(x):
    """Làm tròn về số nguyên kiểu Excel ROUND (nửa lên)."""
    return int(Decimal(x).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def txt(v):
    """Ô chữ: ngày về dd/mm/yyyy, số thực nguyên bỏ phần .0, còn lại strip."""
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def norm_policy(v):
    return txt(v).upper().replace(" ", "")


def norm_sp(v):
    """Mã sản phẩm luôn 6 ký tự: Excel hay biến 050201 thành 50201."""
    s = txt(v)
    if re.fullmatch(r"\d+(\.0)?", s):
        s = s.split(".")[0].zfill(6)
    return s


def norm_inv(v):
    """Số hóa đơn và số chứng từ luôn 8 chữ số."""
    s = txt(v)
    if re.fullmatch(r"\d+(\.0)?", s):
        s = s.split(".")[0].zfill(8)
    return s


def as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = txt(v)
    for f in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def month_end(d):
    """Ngày cuối tháng của d. Kế toán chốt kỳ vào cuối tháng nên cột mvv lấy mốc này."""
    return dt.date(d.year + (d.month == 12), d.month % 12 + 1, 1) - dt.timedelta(days=1)


def months_between(a, b):
    """Số tháng trọn vẹn từ a tới b, giống DATEDIF(a,b,"m") của Excel."""
    if not a or not b:
        return None
    m = (b.year - a.year) * 12 + (b.month - a.month)
    if b.day < a.day:
        m -= 1
    return m


# ----------------------------------------------------------------- dò cột theo tên

LEDGER_NAMES = {
    "so_ctu": ["Số ctừ", "Số chứng từ"],
    "ngay_ctu": ["Ngày ctừ", "Ngày chứng từ"],
    "so_don": ["Số đơn Bh", "Số đơn BH"],
    "tk_no": ["Tk nợ"],
    "ma_sp": ["Mã sp"],
    "phi": ["Nguyên tệ(Phí)", "Nguyên tệ phí"],
    "vat": ["Nguyên tệ(VAT)", "Nguyên tệ VAT"],
    "khach": ["Khách"],
    "phong": ["Phòng Kthác", "Phòng K.thác"],
    "dai_ly": ["Đại lý", "Mã Đại lý"],
    "ky_hieu": ["Ký hiệu HĐ"],
    "can_bo": ["Cán bộ"],
    "ngay_dau": ["Ngày đầu"],
    "ngay_cuoi": ["ngay_cuoi", "Ngày cuối"],
    "don_vi": ["Đơn vị"],
}
LEDGER_NEED = ("so_ctu", "so_don", "ma_sp", "phi")

DETAIL_NAMES = {
    "so_don": ["Số đơn BH", "Số đơn Bh"],
    "ma_sp": ["Mã sp"],
    "phi": ["Nguyên tệ phí", "Nguyên tệ(Phí)"],
    "phi_vat": ["Số tiền phí"],
    "hoa_don": ["Hóa đơn", "Hoá đơn"],
    "dai_ly": ["Mã Đại lý", "Đại lý"],
    "ngay_cap": ["Ngày cấp đơn"],
    "ngay_cuoi": ["Ngày cuối", "ngày cuối"],
    "seri": ["Số seri"],
}
DETAIL_NEED = ("so_don", "ma_sp", "phi")

# cột áp bóc tách sẽ điền vào file chi tiết, theo đúng tên kế toán đang dùng
APPLIED = ["Hóa đơn", "13111", "MK", "MÃ CB", "MÃ PB", "ngày đầu", "ngày cuối",
           "HHĐL", "mvv", "pdv", "Đối trừ"]


def map_cols(header, names):
    """{tên logic: chỉ số cột 0-based} từ danh sách tiêu đề đọc được."""
    seen = {}
    for i, h in enumerate(header):
        k = plain(h)
        if k and k not in seen:
            seen[k] = i
    out = {}
    for logical, cands in names.items():
        for c in cands:
            if plain(c) in seen:
                out[logical] = seen[plain(c)]
                break
    return out


def find_header(ws, names, need):
    """(chỉ số dòng tiêu đề 1-based, danh sách tiêu đề, map cột) hoặc None nếu sheet không phải dữ liệu."""
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=MAX_HEADER_ROW, values_only=True), 1):
        header = [txt(v) for v in row]
        cols = map_cols(header, names)
        if all(k in cols for k in need):
            return i, header, cols
    return None


# ----------------------------------------------------------------- đọc file chi tiết

class DetailSheet:
    """Một sheet dữ liệu của file chi tiết."""

    def __init__(self, name, header, cols, head_row):
        self.name = name
        self.header = header
        self.cols = cols
        self.head_row = head_row
        self.rows = []          # [{"src": [giá trị gốc], "key": (đơn, mã sp), "phi": Decimal, ...}]

    def get(self, row, logical):
        i = self.cols.get(logical)
        return row[i] if i is not None and i < len(row) else None


def read_detail(path, log=print):
    """Đọc mọi sheet dữ liệu của file chi tiết. Sheet BBDS và sheet pivot tự bị bỏ qua."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheets, skipped = [], []
    try:
        for ws in wb.worksheets:
            found = find_header(ws, DETAIL_NAMES, DETAIL_NEED)
            if not found:
                skipped.append(ws.title)
                continue
            head_row, header, cols = found
            sh = DetailSheet(ws.title, header, cols, head_row)
            width = len(header)
            for row in ws.iter_rows(min_row=head_row + 1, values_only=True):
                src = list(row[:width]) + [None] * max(0, width - len(row))
                don = norm_policy(sh.get(src, "so_don"))
                if "/" not in don:                      # dòng trống hay dòng SUBTOTAL cuối bảng
                    continue
                phi = dec(sh.get(src, "phi"))
                sh.rows.append({
                    "src": src,
                    "don": don,
                    "sp": norm_sp(sh.get(src, "ma_sp")),
                    "phi": phi if phi is not None else Decimal(0),
                    "hoa_don": norm_inv(sh.get(src, "hoa_don")),
                    "dai_ly": txt(sh.get(src, "dai_ly")),
                    "ngay_cap": as_date(sh.get(src, "ngay_cap")),
                })
            if sh.rows:
                sheets.append(sh)
                log("Detail sheet %r: %d rows, header at row %d" % (sh.name, len(sh.rows), head_row))
            else:
                skipped.append(ws.title)
    finally:
        wb.close()
    if not sheets:
        raise ToolError("No detail sheet found in the detail file. A data sheet needs the columns "
                        "Số đơn BH, Mã sp and Nguyên tệ phí.")
    if skipped:
        log("Sheets without detail rows (skipped): %s" % ", ".join(repr(s) for s in skipped))
    return sheets


# ----------------------------------------------------------------- đọc sổ 5111

def read_ledger(path, policies, invoices, log=print, progress=None):
    """Quét sổ một lượt, chỉ giữ phần liên quan tới file chi tiết.

    policies: tập Số đơn BH có trong file chi tiết.
    invoices: tập số hóa đơn có trong file chi tiết.
    Trả về (per_key, inv_rows, meta):
      per_key[(đơn, mã sp)] = danh sách dòng sổ của khoá đó
      inv_rows               = các dòng sổ có số chứng từ nằm trong invoices
      meta                   = {"rows": tổng dòng, "max_ngay_ctu": ngày, "cols": map cột}
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        found = find_header(ws, LEDGER_NAMES, LEDGER_NEED)
        if not found:
            raise ToolError("This does not look like the 5111 ledger: the columns Số ctừ, Số đơn Bh, "
                            "Mã sp and Nguyên tệ(Phí) were not all found.")
        head_row, header, cols = found
        log("Ledger header at row %d, %d columns" % (head_row, len([h for h in header if h])))
        per_key = defaultdict(list)
        inv_rows = []
        by_ctu = defaultdict(lambda: {"n": 0, "phi": Decimal(0)})     # (số ctừ, đại lý) -> cộng dồn
        n, max_ctu = 0, None
        for row in ws.iter_rows(min_row=head_row + 1, values_only=True):
            g = lambda k: row[cols[k]] if k in cols and cols[k] < len(row) else None  # noqa: E731
            don = norm_policy(g("so_don"))
            ctu = norm_inv(g("so_ctu"))
            if not don and not ctu:
                continue
            n += 1
            if progress and n % 20000 == 0:
                progress(n)
            d = as_date(g("ngay_ctu"))
            if d and (max_ctu is None or d > max_ctu):
                max_ctu = d
            phi = dec(g("phi"))
            if ctu:
                b = by_ctu[(ctu, txt(g("dai_ly")))]
                b["n"] += 1
                b["phi"] += phi if phi is not None else Decimal(0)
            want_key = don in policies
            want_inv = ctu in invoices
            if not (want_key or want_inv):
                continue
            item = {
                "ctu": ctu, "ngay_ctu": d, "don": don, "sp": norm_sp(g("ma_sp")),
                "phi": phi if phi is not None else Decimal(0),
                "vat": dec(g("vat")) or Decimal(0),
                "tk_no": txt(g("tk_no")), "khach": txt(g("khach")), "can_bo": txt(g("can_bo")),
                "phong": txt(g("phong")), "dai_ly": txt(g("dai_ly")), "ky_hieu": txt(g("ky_hieu")),
                "ngay_dau": as_date(g("ngay_dau")), "ngay_cuoi": as_date(g("ngay_cuoi")),
            }
            if want_key:
                per_key[(item["don"], item["sp"])].append(item)
            if want_inv:
                inv_rows.append(item)
        log("Ledger: %d rows read, %d kept by policy, %d kept by invoice, %d voucher numbers"
            % (n, sum(len(v) for v in per_key.values()), len(inv_rows), len({k[0] for k in by_ctu})))
        return dict(per_key), inv_rows, {"rows": n, "max_ngay_ctu": max_ctu, "cols": cols,
                                         "by_ctu": dict(by_ctu)}
    finally:
        wb.close()


# ----------------------------------------------------------------- biểu tỷ lệ

def rates_path():
    return os.path.join(ROOT, RATES_NAME)


def template_path():
    return os.path.join(ROOT, BBDS_NAME)


def read_rates(path=None):
    """{mã sp: {"ten","vat","hh","ho_tro","pdv","lam_tron"}} đọc từ Offset_Rates.xlsx."""
    p = path or rates_path()
    if not os.path.isfile(p):
        raise ToolError("%s is missing next to the tool. Download it from the page, or ask Nam." % RATES_NAME)
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        names = {"ma_sp": ["Mã sp"], "ten": ["Tên sản phẩm"], "vat": ["Tỷ lệ VAT phí bảo hiểm", "Tỷ lệ VAT"],
                 "vat_pdv": ["Tỷ lệ VAT phí dịch vụ"], "hh": ["Tỷ lệ hoa hồng"],
                 "ho_tro": ["Tỷ lệ hỗ trợ đại lý"], "pdv": ["Tỷ lệ phí dịch vụ"], "lam_tron": ["Làm tròn"]}
        found = find_header(ws, names, ("ma_sp", "hh", "pdv"))
        if not found:
            raise ToolError("%s needs the columns Mã sp, Tỷ lệ hoa hồng and Tỷ lệ phí dịch vụ." % RATES_NAME)
        head_row, header, cols = found
        out = {}
        for row in ws.iter_rows(min_row=head_row + 1, values_only=True):
            g = lambda k: row[cols[k]] if k in cols and cols[k] < len(row) else None  # noqa: E731
            sp = norm_sp(g("ma_sp"))
            if not re.fullmatch(r"\d{4,8}", sp):      # dòng trống hoặc dòng ghi chú dưới bảng
                continue
            out[sp] = {
                "ten": txt(g("ten")),
                "vat": dec(g("vat")) or Decimal(0),
                "vat_pdv": dec(g("vat_pdv")) if dec(g("vat_pdv")) is not None else Decimal("0.1"),
                "hh": dec(g("hh")) or Decimal(0),
                "ho_tro": dec(g("ho_tro")) or Decimal(0),
                "pdv": dec(g("pdv")) or Decimal(0),
                "lam_tron": (txt(g("lam_tron")) or "rải dư"),
            }
        if not out:
            raise ToolError("%s has no product row." % RATES_NAME)
        return out, head_row
    finally:
        wb.close()


# ----------------------------------------------------------------- áp bóc tách

def _common(items, field):
    """Giá trị hay gặp nhất của một trường trong các dòng sổ, bỏ ô rỗng."""
    vals = [i[field] for i in items if i.get(field) not in (None, "")]
    return Counter(vals).most_common(1)[0][0] if vals else None


def spread_rounding(rows, amount_of, rate, divisor):
    """Làm tròn từng dòng rồi rải phần dư để tổng khớp đúng tổng tính một lần.

    Trả về {id(dòng): số nguyên}. Kế toán đang làm tay việc này bằng cách cộng trừ 1 ở vài dòng,
    nên tool phải ra đúng con số đó, không được chỉ làm tròn từng dòng.
    """
    if not rows:
        return {}
    exact = [(amount_of(r) * rate) / divisor for r in rows]
    out = {id(r): r0(v) for r, v in zip(rows, exact)}
    target = r0(sum(amount_of(r) for r in rows) * rate / divisor)
    residual = target - sum(out.values())
    if residual:
        step = 1 if residual > 0 else -1
        # ưu tiên dòng có phần thập phân xa số nguyên nhất, cộng trừ mỗi dòng 1 đồng
        order = sorted(zip(rows, exact), key=lambda t: (t[1] - int(t[1])), reverse=residual > 0)
        for r, _ in order[:abs(residual)]:
            out[id(r)] += step
    return out


def apply_rates(sheets, per_key, rates, chot, log=print):
    """Điền số hóa đơn, mã kế toán, hoa hồng, phí dịch vụ vào từng dòng chi tiết.

    chot: ngày chốt kỳ, dùng cho cột mvv (số tháng từ ngày chốt tới ngày cuối của đơn).
    """
    by_policy = defaultdict(list)
    for (don, _sp), items in per_key.items():
        by_policy[don].extend(items)

    missing_rate = set()
    for sh in sheets:
        matched = [r for r in sh.rows if per_key.get((r["don"], r["sp"]))]
        fallback = {f: _common([i for r in matched for i in per_key[(r["don"], r["sp"])]], f)
                    for f in ("tk_no", "khach", "can_bo", "phong")}
        for r in sh.rows:
            items = per_key.get((r["don"], r["sp"])) or by_policy.get(r["don"]) or []
            r["in_ledger"] = bool(per_key.get((r["don"], r["sp"])))
            # Chỉ điền số hóa đơn khi sổ có ĐÚNG MỘT chứng từ cho khoá đó. Đơn tập thể hay có hai
            # chứng từ (ví dụ CK17014 có 00005931 và 00005934), đoán một số là gán lệch cả đống phí
            # sang sai hóa đơn rồi báo lệch oan.
            ctus = sorted({i["ctu"] for i in items if i["ctu"]})
            r["ctu_ledger"] = ctus[0] if len(ctus) == 1 else ""
            r["ctu_many"] = ctus if len(ctus) > 1 else []
            r["hd_lech_so"] = bool(r["hoa_don"] and ctus and r["hoa_don"] not in ctus)
            r["nguon"] = "sổ 5111" if items else "suy ra"
            r["tk_no"] = _common(items, "tk_no") or fallback["tk_no"] or ""
            r["khach"] = _common(items, "khach") or fallback["khach"] or ""
            r["can_bo"] = _common(items, "can_bo") or fallback["can_bo"] or ""
            r["phong"] = _common(items, "phong") or fallback["phong"] or ""
            r["ngay_dau"] = _common(items, "ngay_dau")
            r["ngay_cuoi"] = _common(items, "ngay_cuoi") or as_date(sh.get(r["src"], "ngay_cuoi"))
            r["mvv"] = months_between(chot, r["ngay_cuoi"])
            if r["sp"] not in rates:
                missing_rate.add(r["sp"])

    if missing_rate:
        raise ToolError("No rate for product code(s) %s. Add a row for each in %s (download it on "
                        "the page, add the row in Excel, upload it again)."
                        % (", ".join(sorted(missing_rate)), RATES_NAME))

    # rải dư theo mã sản phẩm, tính trên toàn bộ các sheet vì BBDS cộng theo sản phẩm
    rows_by_sp = defaultdict(list)
    for sh in sheets:
        for r in sh.rows:
            rows_by_sp[r["sp"]].append(r)
    for sp, rows in rows_by_sp.items():
        rt = rates[sp]
        one = Decimal(1)
        hh = spread_rounding(rows, lambda r: r["phi"], rt["hh"], one)
        pdv = spread_rounding(rows, lambda r: r["phi"], rt["pdv"], one + rt["vat_pdv"])
        for r in rows:
            r["hh"] = hh[id(r)]
            r["pdv"] = pdv[id(r)]
            r["hh_adj"] = r["hh"] - r0(r["phi"] * rt["hh"])
            r["pdv_adj"] = r["pdv"] - r0(r["phi"] * rt["pdv"] / (one + rt["vat_pdv"]))
        log("Product %s: %d rows, commission %s, service fee %s"
            % (sp, len(rows), f"{sum(hh.values()):,}", f"{sum(pdv.values()):,}"))
    return rows_by_sp


# ----------------------------------------------------------------- đối trừ

def offset(sheets, per_key, inv_rows, detail_agents, by_ctu, log=print):
    """So khớp hai mức: theo hóa đơn và theo đơn. Trả về dict các bảng kết quả."""
    led_agents = {i["dai_ly"] for items in per_key.values() for i in items if i["dai_ly"]}
    agents = (detail_agents | led_agents) or None
    log("Agent scope: %s" % (", ".join(sorted(agents)) if agents else "every agent in the ledger"))

    # ----- mức đơn
    det_key = defaultdict(lambda: {"n": 0, "phi": Decimal(0), "hd": set(), "sheet": set()})
    for sh in sheets:
        for r in sh.rows:
            k = (r["don"], r["sp"])
            d = det_key[k]
            d["n"] += 1
            d["phi"] += r["phi"]
            d["sheet"].add(sh.name)
            if r["hoa_don"] or r["ctu_ledger"]:
                d["hd"].add(r["hoa_don"] or r["ctu_ledger"])
    by_policy, missing = [], []
    for k, d in sorted(det_key.items()):
        items = [i for i in per_key.get(k, []) if not agents or i["dai_ly"] in agents]
        phi_led = sum((i["phi"] for i in items), Decimal(0))
        row = {"don": k[0], "sp": k[1], "hd": ", ".join(sorted(d["hd"])), "n_det": d["n"],
               "phi_det": d["phi"], "n_led": len(items), "phi_led": phi_led,
               "sheet": ", ".join(sorted(d["sheet"]))}
        by_policy.append(row)
        if not items:
            missing.append(row)

    # ----- mức hóa đơn
    det_inv = defaultdict(lambda: {"n": 0, "phi": Decimal(0)})
    for sh in sheets:
        for r in sh.rows:
            inv = r["hoa_don"] or r["ctu_ledger"]
            if not inv:
                continue
            det_inv[inv]["n"] += 1
            det_inv[inv]["phi"] += r["phi"]
    led_inv = defaultdict(lambda: {"n": 0, "phi": Decimal(0)})
    for (ctu, dai_ly), b in by_ctu.items():
        if ctu not in det_inv or (agents and dai_ly not in agents):
            continue
        led_inv[ctu]["n"] += b["n"]
        led_inv[ctu]["phi"] += b["phi"]
    by_invoice = []
    for inv in sorted(set(det_inv) | set(led_inv)):
        d, l = det_inv.get(inv, {"n": 0, "phi": Decimal(0)}), led_inv.get(inv, {"n": 0, "phi": Decimal(0)})
        by_invoice.append({"hd": inv, "n_det": d["n"], "phi_det": d["phi"],
                           "n_led": l["n"], "phi_led": l["phi"]})

    # ----- thừa trong sổ: dòng sổ thuộc các hóa đơn đang đối trừ mà file chi tiết không có
    det_keys = set(det_key)
    extra = [i for i in inv_rows
             if (not agents or i["dai_ly"] in agents) and (i["don"], i["sp"]) not in det_keys]

    # ----- chưa gán được hóa đơn: file chi tiết để trống mà sổ có nhiều hơn một chứng từ,
    # hoặc khoá chưa có trong sổ. Phí mấy dòng này chỉ đối trừ được ở mức đơn.
    un = defaultdict(lambda: {"n": 0, "phi": Decimal(0), "ctus": set(), "sheet": set()})
    for sh in sheets:
        for r in sh.rows:
            if r["hoa_don"] or r["ctu_ledger"]:
                continue
            u = un[(r["don"], r["sp"])]
            u["n"] += 1
            u["phi"] += r["phi"]
            u["ctus"].update(r["ctu_many"])
            u["sheet"].add(sh.name)
    unassigned = [{"don": k[0], "sp": k[1], "n_det": v["n"], "phi_det": v["phi"],
                   "ctus": ", ".join(sorted(v["ctus"])), "sheet": ", ".join(sorted(v["sheet"]))}
                  for k, v in sorted(un.items())]

    # ----- hóa đơn trong file không nằm trong các chứng từ của đơn đó trong sổ
    mismatch = [{"don": r["don"], "sp": r["sp"], "hd": r["hoa_don"], "phi_det": r["phi"],
                 "ctus": ", ".join(r["ctu_many"] or [i["ctu"] for i in per_key.get((r["don"], r["sp"]), [])]),
                 "sheet": sh.name}
                for sh in sheets for r in sh.rows if r["hd_lech_so"]]

    return {"by_invoice": by_invoice, "by_policy": by_policy, "missing": missing,
            "extra": sorted(extra, key=lambda i: (i["ctu"], i["don"], i["sp"])),
            "unassigned": unassigned, "mismatch": mismatch,
            "agents": sorted(agents) if agents else []}


# ----------------------------------------------------------------- ghi file Excel

def _head(ws, row, labels, widths=None, font="Arial", size=10):
    for j, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=j, value=lab)
        c.font = Font(name=font, size=size, bold=True)
        c.fill = FILL_HEAD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(bottom=THIN)
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[row].height = 30


def _cell(ws, row, col, value, fmt=None, font="Arial", size=10, bold=False, align=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name=font, size=size, bold=bold)
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = Alignment(horizontal=align, vertical="center")
    return c


MONEY = "#,##0;(#,##0);-"
PCT = "0.0%"
DATE = "dd/mm/yyyy"


def write_offset_file(path, res, info, applied=None):
    """Offset_5111_<ngày>.xlsx: Tổng hợp, Theo hóa đơn, Theo đơn, Thiếu trong sổ, Thừa trong sổ."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tổng hợp"

    # ----- Theo hóa đơn
    s1 = wb.create_sheet("Theo hóa đơn")
    _head(s1, 1, ["Hóa đơn", "Số dòng chi tiết", "Phí chi tiết", "Số dòng sổ", "Phí sổ 5111",
                  "Lệch", "Trạng thái"], [14, 14, 16, 12, 16, 14, 14])
    for i, r in enumerate(res["by_invoice"], 2):
        _cell(s1, i, 1, r["hd"], align="center")
        _cell(s1, i, 2, r["n_det"], MONEY)
        _cell(s1, i, 3, float(r["phi_det"]), MONEY)
        _cell(s1, i, 4, r["n_led"], MONEY)
        _cell(s1, i, 5, float(r["phi_led"]), MONEY)
        c = _cell(s1, i, 6, "=E%d-C%d" % (i, i), MONEY)
        _cell(s1, i, 7, '=IF(F%d=0,"Khớp",IF(D%d=0,"Không có trong sổ","Lệch"))' % (i, i), align="center")
        if r["phi_led"] != r["phi_det"]:
            c.font = RED
            for j in range(1, 8):
                s1.cell(row=i, column=j).fill = FILL_BAD if r["n_led"] == 0 else FILL_WARN
    end1 = len(res["by_invoice"]) + 1
    if end1 >= 2:
        _cell(s1, end1 + 1, 1, "TỔNG", bold=True)
        for col in ("B", "C", "D", "E", "F"):
            _cell(s1, end1 + 1, ord(col) - 64, "=SUM(%s2:%s%d)" % (col, col, end1), MONEY, bold=True)
    s1.freeze_panes = "A2"

    # ----- Theo đơn
    s2 = wb.create_sheet("Theo đơn")
    _head(s2, 1, ["Số đơn BH", "Mã sp", "Hóa đơn", "Số dòng chi tiết", "Phí chi tiết",
                  "Số dòng sổ", "Phí sổ 5111", "Lệch", "Trạng thái", "Sheet"],
          [30, 10, 14, 14, 16, 12, 16, 14, 20, 16])
    for i, r in enumerate(res["by_policy"], 2):
        _cell(s2, i, 1, r["don"])
        _cell(s2, i, 2, r["sp"], align="center")
        _cell(s2, i, 3, r["hd"], align="center")
        _cell(s2, i, 4, r["n_det"], MONEY)
        _cell(s2, i, 5, float(r["phi_det"]), MONEY)
        _cell(s2, i, 6, r["n_led"], MONEY)
        _cell(s2, i, 7, float(r["phi_led"]), MONEY)
        c = _cell(s2, i, 8, "=G%d-E%d" % (i, i), MONEY)
        _cell(s2, i, 9, '=IF(F%d=0,"Chưa có trong sổ",IF(H%d=0,"Khớp","Lệch"))' % (i, i), align="center")
        _cell(s2, i, 10, r["sheet"])
        if r["phi_led"] != r["phi_det"]:
            c.font = RED
            for j in range(1, 11):
                s2.cell(row=i, column=j).fill = FILL_BAD if r["n_led"] == 0 else FILL_WARN
    end2 = len(res["by_policy"]) + 1
    if end2 >= 2:
        _cell(s2, end2 + 1, 1, "TỔNG", bold=True)
        for col in ("D", "E", "F", "G", "H"):
            _cell(s2, end2 + 1, ord(col) - 64, "=SUM(%s2:%s%d)" % (col, col, end2), MONEY, bold=True)
    s2.freeze_panes = "A2"

    # ----- Thiếu trong sổ
    s3 = wb.create_sheet("Thiếu trong sổ")
    _head(s3, 1, ["Số đơn BH", "Mã sp", "Hóa đơn", "Số dòng chi tiết", "Phí chi tiết", "Sheet"],
          [30, 10, 14, 14, 16, 16])
    for i, r in enumerate(res["missing"], 2):
        _cell(s3, i, 1, r["don"])
        _cell(s3, i, 2, r["sp"], align="center")
        _cell(s3, i, 3, r["hd"], align="center")
        _cell(s3, i, 4, r["n_det"], MONEY)
        _cell(s3, i, 5, float(r["phi_det"]), MONEY)
        _cell(s3, i, 6, r["sheet"])
    end3 = len(res["missing"]) + 1
    if end3 >= 2:
        _cell(s3, end3 + 1, 1, "TỔNG", bold=True)
        for col in ("D", "E"):
            _cell(s3, end3 + 1, ord(col) - 64, "=SUM(%s2:%s%d)" % (col, col, end3), MONEY, bold=True)
    s3.freeze_panes = "A2"

    # ----- Thừa trong sổ
    s4 = wb.create_sheet("Thừa trong sổ")
    _head(s4, 1, ["Số ctừ", "Ngày ctừ", "Số đơn Bh", "Mã sp", "Phí sổ 5111", "Tk nợ",
                  "Đại lý", "Ký hiệu HĐ"], [14, 13, 30, 10, 16, 10, 16, 12])
    for i, r in enumerate(res["extra"], 2):
        _cell(s4, i, 1, r["ctu"], align="center")
        _cell(s4, i, 2, r["ngay_ctu"], DATE, align="center")
        _cell(s4, i, 3, r["don"])
        _cell(s4, i, 4, r["sp"], align="center")
        _cell(s4, i, 5, float(r["phi"]), MONEY)
        _cell(s4, i, 6, r["tk_no"], align="center")
        _cell(s4, i, 7, r["dai_ly"], align="center")
        _cell(s4, i, 8, r["ky_hieu"], align="center")
    end4 = len(res["extra"]) + 1
    if end4 >= 2:
        _cell(s4, end4 + 1, 1, "TỔNG", bold=True)
        _cell(s4, end4 + 1, 5, "=SUM(E2:E%d)" % end4, MONEY, bold=True)
    s4.freeze_panes = "A2"

    # ----- Hóa đơn chưa gán: chỉ đối trừ được ở mức đơn
    s5 = wb.create_sheet("Hóa đơn chưa gán")
    _head(s5, 1, ["Số đơn BH", "Mã sp", "Số dòng chi tiết", "Phí chi tiết",
                  "Chứng từ của đơn trong sổ", "Sheet", "Lý do"], [30, 10, 14, 16, 30, 16, 34])
    for i, r in enumerate(res["unassigned"], 2):
        _cell(s5, i, 1, r["don"])
        _cell(s5, i, 2, r["sp"], align="center")
        _cell(s5, i, 3, r["n_det"], MONEY)
        _cell(s5, i, 4, float(r["phi_det"]), MONEY)
        _cell(s5, i, 5, r["ctus"], align="center")
        _cell(s5, i, 6, r["sheet"])
        _cell(s5, i, 7, "Sổ có nhiều chứng từ cho đơn này" if r["ctus"] else "Đơn chưa có trong sổ")
    end5 = len(res["unassigned"]) + 1
    if end5 >= 2:
        _cell(s5, end5 + 1, 1, "TỔNG", bold=True)
        for col in ("C", "D"):
            _cell(s5, end5 + 1, ord(col) - 64, "=SUM(%s2:%s%d)" % (col, col, end5), MONEY, bold=True)
    s4.freeze_panes = "A2"

    # ----- Tổng hợp (đọc số từ các sheet trên bằng công thức)
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 46
    _cell(ws, 1, 1, "ĐỐI TRỪ VỚI SỔ 5111", bold=True, size=13)
    rows = [
        ("Kỳ chi tiết", info["ky"], "Theo Ngày cấp đơn trong file chi tiết"),
        ("Ngày chốt (cột mvv)", info["chot"], "Số tháng từ ngày này tới ngày cuối của đơn"),
        ("File chi tiết", info["detail_name"], "%d dòng, %d sheet" % (info["det_rows"], info["det_sheets"])),
        ("Sổ 5111", info["ledger_name"], "%s dòng đã quét" % f"{info['led_rows']:,}"),
        ("Mã đại lý dùng để lọc sổ", ", ".join(res["agents"]) or "mọi đại lý",
         "Lấy từ cột Mã Đại lý của file chi tiết và từ các dòng sổ khớp đơn"),
        (None, None, None),
        ("Phí chi tiết (chưa VAT)", "='Theo đơn'!E%d" % (end2 + 1), "Tổng cột Phí chi tiết"),
        ("Phí trong sổ 5111", "='Theo đơn'!G%d" % (end2 + 1), "Chỉ các dòng sổ khớp khoá đơn"),
        ("Lệch theo đơn", "='Theo đơn'!H%d" % (end2 + 1), "Bằng 0 là khớp hết"),
        (None, None, None),
        ("Số hóa đơn đối trừ", "=COUNTA('Theo hóa đơn'!A2:A%d)" % end1, None),
        ("Hóa đơn khớp", '=COUNTIF(\'Theo hóa đơn\'!G2:G%d,"Khớp")' % end1, None),
        ("Hóa đơn lệch", '=COUNTIF(\'Theo hóa đơn\'!G2:G%d,"Lệch")' % end1, None),
        ("Hóa đơn không có trong sổ", '=COUNTIF(\'Theo hóa đơn\'!G2:G%d,"Không có trong sổ")' % end1, None),
        (None, None, None),
        ("Khoá đơn và sản phẩm", "=COUNTA('Theo đơn'!A2:A%d)" % end2, None),
        ("Khoá chưa có trong sổ", "=COUNTA('Thiếu trong sổ'!A2:A%d)" % end3, "Phí: xem sheet Thiếu trong sổ"),
        ("Phí chưa có trong sổ", "='Thiếu trong sổ'!E%d" % (end3 + 1), None),
        ("Hoa hồng đại lý (chi tiết đã áp)", (applied or {}).get("hh"),
         "Bằng ô Hoa hồng đại lý của BBDS trong file chi tiết đã áp"),
        ("Phí dịch vụ chưa VAT (chi tiết đã áp)", (applied or {}).get("pdv"),
         "BBDS ghi phí dịch vụ ĐÃ gồm VAT, chia (1 + VAT phí dịch vụ) ra số này"),
        (None, None, None),
        ("Dòng sổ thừa so với chi tiết", "=COUNTA('Thừa trong sổ'!A2:A%d)" % end4, None),
        ("Phí thừa trong sổ", "='Thừa trong sổ'!E%d" % (end4 + 1), None),
        (None, None, None),
        ("Khoá chưa gán được hóa đơn", "=COUNTA('Hóa đơn chưa gán'!A2:A%d)" % end5,
         "File chi tiết để trống mà sổ có nhiều chứng từ, hoặc đơn chưa vào sổ"),
        ("Phí chưa gán hóa đơn", "='Hóa đơn chưa gán'!D%d" % (end5 + 1),
         "Phần này chỉ đối trừ ở mức đơn, không vào bảng theo hóa đơn"),
        ("Dòng có hóa đơn khác sổ", len(res["mismatch"]),
         "Số hóa đơn trong file không nằm trong chứng từ của đơn đó trong sổ"),
    ]
    r = 3
    for label, value, note in rows:
        if label is None:
            r += 1
            continue
        _cell(ws, r, 1, label, bold=True)
        c = _cell(ws, r, 2, value)
        if isinstance(value, str) and value.startswith("=") or isinstance(value, (int, float)):
            c.number_format = MONEY
        if note:
            _cell(ws, r, 3, note).font = Font(name="Arial", size=9, italic=True, color="808080")
        r += 1
    for s in wb.worksheets:
        s.sheet_view.showGridLines = False
    wb.save(path)
    return path


def pct_lit(v):
    """0.616 -> "61.6%" để viết thẳng vào công thức, y như file mẫu của kế toán."""
    s = ("%.6f" % (float(v) * 100)).rstrip("0").rstrip(".")
    return (s or "0") + "%"


def num_lit(v):
    s = ("%.6f" % float(v)).rstrip("0").rstrip(".")
    return s or "0"


def _sheet_ref(name):
    return "'%s'" % str(name).replace("'", "''")


def fill_bbds(ws, used, rates, sheet_info, info, rate_rows, log=print):
    """Điền số vào sheet BBDS của template, KHÔNG dựng lại sheet.

    Anh Nam 2026-10-14: "sheet BBDS phải giống hệt format file mẫu". Nên chỗ này chỉ ghi tháng ở
    tiêu đề, kỳ thanh toán, ngày ký, rồi mỗi sản phẩm có số liệu thì ghi số đơn (COUNTIF), phí chưa
    VAT (SUMIF) và 4 ô tỷ lệ. Định dạng, công thức F/H/J/L, khối nghiệm thu, khối chữ ký: của template.
    """
    head = next((r for r in range(1, 30) if txt(ws.cell(row=r, column=1).value) == "STT"), None)
    total = next((r for r in range(head or 1, (head or 1) + 60)
                  if "TỔNG CỘNG" in txt(ws.cell(row=r, column=1).value).upper()), None)
    if not head or not total:
        raise ToolError("The BBDS template has no STT header row or no TỔNG CỘNG row. Rebuild it "
                        "with make_bbds_template.py.")

    for r in range(1, head):                             # tiêu đề và kỳ thanh toán
        for col in (1, 2, 3):
            c = ws.cell(row=r, column=col)
            if isinstance(c, MergedCell):
                continue
            v = txt(c.value)
            if "BIÊN BẢN ĐỐI SOÁT" in v.upper() and "THÁNG" in v.upper():
                c.value = re.sub(r"THÁNG\s*[0-9/]*\s*$", "THÁNG %s" % info["thang"], v)
            elif v.startswith("Kỳ thanh toán") and info["ky"]:
                c.value = "Kỳ thanh toán: %s" % info["ky"]
    for r in range(total, total + 30):                   # dòng ngày ký dưới bảng
        c = ws.cell(row=r, column=1)
        if isinstance(c, MergedCell):
            continue
        v = txt(c.value)
        if v.startswith("Ngày") and "năm" in v:
            # mẫu ghi tháng LẬP biên bản (kỳ T08 nhưng ký tháng 09), nên lấy tháng hiện tại
            today = dt.date.today()
            c.value = "Ngày      tháng %02d năm %d" % (today.month, today.year)
            break

    rows_by_name = {}                                    # tên sản phẩm -> dòng, trùng tên lấy dòng trên
    for r in range(total - 1, head, -1):
        n = plain(ws.cell(row=r, column=2).value)
        if n:
            rows_by_name[n] = r

    hit = {}
    for sp in used:
        r = rows_by_name.get(plain(rates[sp]["ten"]))
        if r is None:
            raise ToolError("The BBDS template has no line for product %s (%s). Add that line to %s, "
                            "or make the product name in %s match the wording in the template."
                            % (sp, rates[sp]["ten"], BBDS_NAME, RATES_NAME))
        hit[r] = sp

    for r in range(head + 1, total):
        sp = hit.get(r)
        cell = lambda col: (None if isinstance(ws.cell(row=r, column=col), MergedCell)
                            else ws.cell(row=r, column=col))          # noqa: E731
        if sp is None:                                   # sản phẩm kỳ này không phát sinh: để trống
            for col in (3, 4):
                if cell(col) is not None:
                    cell(col).value = None
            continue
        rt = rates[sp]
        count = "+".join("COUNTIF(%s!$%s:$%s,\"%s\")" % (_sheet_ref(s["name"]), s["sp"], s["sp"], sp)
                         for s in sheet_info)
        money = "+".join("SUMIF(%s!$%s:$%s,\"%s\",%s!$%s:$%s)"
                         % (_sheet_ref(s["name"]), s["sp"], s["sp"], sp,
                            _sheet_ref(s["name"]), s["phi"], s["phi"]) for s in sheet_info)
        if cell(3) is not None:
            cell(3).value = "=" + count
        if cell(4) is not None:
            cell(4).value = "=" + money
        # 4 ô tỷ lệ TRỎ sang sheet Tỷ lệ (anh Nam 2026-10-14: công thức link với sheet đó),
        # tỷ lệ 0 thì để trống như file mẫu
        for col, letter, key in ((5, "C", "vat"), (7, "D", "hh"), (9, "E", "ho_tro"), (11, "F", "pdv")):
            c = cell(col)
            if c is None:
                continue
            c.value = ("='%s'!$%s$%d" % (RATE_SHEET, letter, rate_rows[sp])) if float(rt[key]) else None
        log("BBDS line %d: %s %s" % (r, sp, rt["ten"]))
    ws.sheet_view.showGridLines = False                  # anh Nam 2026-10-14: BBDS tắt gridline
    return head, total


def write_detail_file(path, sheets, rates, info, res, log=print):
    """Detail_Applied_<ngày>.xlsx: sheet BBDS của template + từng sheet chi tiết đã áp.

    Khuôn sheet chi tiết theo đúng file kế toán đang làm tay: dòng 1 là mốc (ngày chốt + tổng các
    cột tiền), dòng 2 là tiêu đề, dữ liệu từ dòng 3. Không thêm sheet nào khác vào file giao.
    """
    tpl = template_path()
    if not os.path.isfile(tpl):
        raise ToolError("%s is missing next to the tool. Rebuild it with make_bbds_template.py."
                        % BBDS_NAME)
    wb = openpyxl.load_workbook(tpl)
    bbds = wb.worksheets[0]

    used = []
    for sh in sheets:
        for r in sh.rows:
            if r["sp"] not in used:
                used.append(r["sp"])
    used.sort()

    # ---- sheet Tỷ lệ: mọi tỷ lệ và ngày chốt nằm ở đây, công thức chỗ khác tra sang
    rt_ws = wb.create_sheet(RATE_SHEET)
    _cell(rt_ws, 1, 1, "Ngày chốt kỳ", bold=True)
    _cell(rt_ws, 1, 2, info["chot"], DATE, bold=True, align="center")
    HR = 3
    _head(rt_ws, HR, ["Mã sp", "Tên sản phẩm", "VAT phí bảo hiểm", "Tỷ lệ hoa hồng",
                      "Tỷ lệ hỗ trợ đại lý", "Tỷ lệ phí dịch vụ", "VAT phí dịch vụ"],
          [10, 52, 15, 14, 16, 16, 15])
    rate_rows = {}
    for i, sp in enumerate(used, HR + 1):
        d = rates[sp]
        rate_rows[sp] = i
        _cell(rt_ws, i, 1, sp, align="center")
        _cell(rt_ws, i, 2, d["ten"])
        for col, key in ((3, "vat"), (4, "hh"), (5, "ho_tro"), (6, "pdv"), (7, "vat_pdv")):
            _cell(rt_ws, i, col, float(d[key]), PCT, align="center")
    rt_ws.sheet_view.showGridLines = False
    RANGE = "'%s'!$A$%d:$G$%d" % (RATE_SHEET, HR + 1, HR + len(used))
    CHOT = "'%s'!$B$1" % RATE_SHEET

    sheet_info = []
    for sh in sheets:
        ws = wb.create_sheet(sh.name[:31])
        header = list(sh.header)
        width = len(header)
        pos, existed = {}, {}                      # tên cột áp bóc tách -> chỉ số cột 1-based
        for name in APPLIED:
            j = next((k for k, h in enumerate(header) if plain(h) == plain(name)), None)
            existed[name] = j is not None and j < width
            if j is None:
                header.append(name)
                j = len(header) - 1
            pos[name] = j + 1
        HDR, FIRST = 2, 3                          # dòng 1 để mốc, dòng 2 tiêu đề, dữ liệu từ dòng 3
        _head(ws, HDR, [h or "" for h in header])
        ws.cell(row=HDR, column=pos["HHĐL"]).comment = Comment(
            "Hoa hồng đại lý = phí × tỷ lệ hoa hồng, làm tròn từng dòng rồi rải phần dư để tổng theo "
            "mã sản phẩm khớp đúng BBDS. Dòng nào có +1 hay -1 là dòng nhận phần dư.", "Offset 5111")
        ws.cell(row=HDR, column=pos["pdv"]).comment = Comment(
            "Phí dịch vụ CHƯA VAT = phí × tỷ lệ phí dịch vụ / (1 + VAT phí dịch vụ), làm tròn từng "
            "dòng rồi rải phần dư. Tỷ lệ lấy từ %s." % RATES_NAME, "Offset 5111")
        c_phi = sh.cols["phi"] + 1
        c_sp = sh.cols["ma_sp"] + 1
        L_cuoi = get_column_letter(pos["ngày cuối"])
        L_sp = get_column_letter(c_sp)
        # ô Mã sp của file gốc có sheet ghi dạng chữ, có sheet ghi dạng số: VLOOKUP tra chữ nên
        # sheet nào ghi số thì phải bọc TEXT(...,"000000")
        sp_text = all(isinstance(r["src"][sh.cols["ma_sp"]], str) for r in sh.rows)
        for i, r in enumerate(sh.rows, FIRST):
            for j, v in enumerate(r["src"][:width], 1):
                _cell(ws, i, j, v, DATE if isinstance(v, (dt.datetime, dt.date)) else None)
            if isinstance(r["src"][sh.cols["phi"]], (int, float)):
                ws.cell(row=i, column=c_phi).number_format = MONEY
            phi_ref = "%s%d" % (get_column_letter(c_phi), i)

            def put(name, value, fmt=None, row=r, i=i):
                """Điền số liệu lấy từ sổ, nhưng KHÔNG ghi đè ô mà file gốc đã có số liệu."""
                if existed[name] and txt(row["src"][pos[name] - 1]):
                    return
                _cell(ws, i, pos[name], value, fmt, align="center")

            put("Hóa đơn", r["hoa_don"] or r["ctu_ledger"])
            put("13111", r["tk_no"])
            put("MK", r["khach"])
            put("MÃ CB", r["can_bo"])
            put("MÃ PB", r["phong"])
            put("ngày đầu", r["ngay_dau"], DATE)
            put("ngày cuối", r["ngay_cuoi"], DATE)
            sp_ref = ("$%s%d" % (L_sp, i)) if sp_text else ('TEXT($%s%d,"000000")' % (L_sp, i))
            adj = ("%+d" % r["hh_adj"]) if r["hh_adj"] else ""
            _cell(ws, i, pos["HHĐL"], "=ROUND(%s*VLOOKUP(%s,%s,4,0),0)%s"
                  % (phi_ref, sp_ref, RANGE, adj), MONEY)
            adj = ("%+d" % r["pdv_adj"]) if r["pdv_adj"] else ""
            _cell(ws, i, pos["pdv"],
                  "=ROUND(%s*VLOOKUP(%s,%s,6,0)/(1+VLOOKUP(%s,%s,7,0)),0)%s"
                  % (phi_ref, sp_ref, RANGE, sp_ref, RANGE, adj), MONEY)
            # DATEDIF ném #NUM! khi ngày cuối trước ngày chốt (đơn đã hết hạn), nên chặn bằng IF
            _cell(ws, i, pos["mvv"],
                  '=IF(%s%d="","",IF(%s%d<=%s,0,DATEDIF(%s,%s%d,"m")))'
                  % (L_cuoi, i, L_cuoi, i, CHOT, CHOT, L_cuoi, i), align="center")
            if not r["in_ledger"]:
                status, fill = "Chưa có trong sổ", FILL_BAD
            elif r["hd_lech_so"]:
                status, fill = "Hóa đơn khác sổ", FILL_BAD
            elif r["ctu_many"] and not r["hoa_don"]:
                status, fill = "Sổ có nhiều hóa đơn", FILL_WARN
            else:
                status, fill = "Khớp sổ", None
            c = _cell(ws, i, pos["Đối trừ"], status, align="center")
            if fill:
                c.fill = fill

        last = FIRST + len(sh.rows) - 1
        for name in ("HHĐL", "pdv"):
            L = get_column_letter(pos[name])
            _cell(ws, 1, pos[name], "=SUM(%s%d:%s%d)" % (L, FIRST, L, last), MONEY, bold=True)
        for logical in ("phi", "phi_vat"):
            j = sh.cols.get(logical)
            if j is not None:
                L = get_column_letter(j + 1)
                _cell(ws, 1, j + 1, "=SUM(%s%d:%s%d)" % (L, FIRST, L, last), MONEY, bold=True)

        # cột ngày mà hẹp thì Excel hiện ####: nới theo nội dung thật, trần 34 cho khỏi dài quá
        for j in range(1, len(header) + 1):
            longest = len(txt(header[j - 1]))
            for i in range(FIRST, min(last + 1, FIRST + 300)):
                v = ws.cell(row=i, column=j).value
                if isinstance(v, (dt.datetime, dt.date)):
                    longest = max(longest, 10)
                elif isinstance(v, str) and v.startswith("="):
                    longest = max(longest, 12)
                else:
                    longest = max(longest, len(txt(v)))
            ws.column_dimensions[get_column_letter(j)].width = min(max(longest + 2, 9), 34)
        ws.freeze_panes = "A%d" % FIRST
        ws.sheet_view.showGridLines = False
        sheet_info.append({"name": ws.title, "rows": len(sh.rows), "first": FIRST, "last": last,
                           "sp": get_column_letter(c_sp), "phi": get_column_letter(c_phi),
                           "hh": get_column_letter(pos["HHĐL"]), "pdv": get_column_letter(pos["pdv"])})

    fill_bbds(bbds, used, rates, sheet_info, info, rate_rows, log=log)
    wb.save(path)
    return path



# ----------------------------------------------------------------- chạy một lượt

def period_text(sheets):
    """(chuỗi kỳ, 'MM/YYYY', 'YYYY') lấy từ Ngày cấp đơn của file chi tiết."""
    days = sorted(r["ngay_cap"] for sh in sheets for r in sh.rows if r["ngay_cap"])
    if not days:
        today = dt.date.today()
        return "", today.strftime("%m/%Y"), str(today.year)
    a, b = days[0], days[-1]
    return ("Từ ngày %s đến ngày %s" % (a.strftime("%d/%m/%Y"), b.strftime("%d/%m/%Y")),
            a.strftime("%m/%Y"), str(b.year))


def run(ledger_path, detail_path, out_dir, rates_file=None, chot=None, ky=None,
        log=print, progress=None):
    """Chạy đủ một lượt: áp bóc tách rồi đối trừ. Trả về dict kết quả cho trang web."""
    os.makedirs(out_dir, exist_ok=True)
    rates, _rates_head = read_rates(rates_file)
    log("Rates: %d product(s) in %s" % (len(rates), os.path.basename(rates_file or rates_path())))

    sheets = read_detail(detail_path, log=log)
    policies = {r["don"] for sh in sheets for r in sh.rows}
    invoices = {r["hoa_don"] for sh in sheets for r in sh.rows if r["hoa_don"]}
    det_rows = sum(len(sh.rows) for sh in sheets)
    detail_agents = {r["dai_ly"] for sh in sheets for r in sh.rows if r["dai_ly"]}
    log("Detail: %d rows, %d policies, %d invoice number(s)" % (det_rows, len(policies), len(invoices)))

    per_key, inv_rows, meta = read_ledger(ledger_path, policies, invoices, log=log, progress=progress)
    chot = chot or month_end(meta["max_ngay_ctu"] or dt.date.today())
    ky_text, thang, nam = period_text(sheets)
    info = {"ky": ky or ky_text, "thang": thang, "nam": nam, "chot": chot,
            "detail_name": os.path.basename(detail_path), "ledger_name": os.path.basename(ledger_path),
            "det_rows": det_rows, "det_sheets": len(sheets), "led_rows": meta["rows"]}

    apply_rates(sheets, per_key, rates, chot, log=log)

    res = offset(sheets, per_key, inv_rows, detail_agents, meta["by_ctu"], log=log)

    # Hóa đơn nào do tool suy ra từ sổ (sheet TNCN không có cột Hóa đơn) mà lại lệch thì chưa đủ
    # dòng sổ để chỉ ra đơn nào gây lệch: quét sổ thêm một lượt riêng cho mấy hóa đơn đó.
    need = {r["hd"] for r in res["by_invoice"] if r["phi_det"] != r["phi_led"]} - invoices
    if need:
        log("Reading the ledger again for %d invoice number(s) the tool took from the ledger" % len(need))
        _k, more, _m = read_ledger(ledger_path, set(), need, log=lambda *_a: None, progress=progress)
        seen = {(i["ctu"], i["don"], i["sp"], str(i["phi"])) for i in inv_rows}
        inv_rows += [i for i in more if (i["ctu"], i["don"], i["sp"], str(i["phi"])) not in seen]
        res = offset(sheets, per_key, inv_rows, detail_agents, meta["by_ctu"], log=lambda *_a: None)

    stamp = dt.date.today().strftime("%Y%m%d")
    p_det = write_detail_file(os.path.join(out_dir, "Detail_Applied_%s.xlsx" % stamp),
                              sheets, rates, info, res, log=log)
    applied = {"hh": sum(r["hh"] for sh in sheets for r in sh.rows),
               "pdv": sum(r["pdv"] for sh in sheets for r in sh.rows)}
    p_off = write_offset_file(os.path.join(out_dir, "Offset_5111_%s.xlsx" % stamp), res, info, applied)
    log("Written: %s" % os.path.basename(p_det))
    log("Written: %s" % os.path.basename(p_off))

    bad_inv = [r for r in res["by_invoice"] if r["phi_det"] != r["phi_led"]]
    return {
        "outputs": [p_off, p_det], "offset_file": p_off, "detail_file": p_det,
        "info": info, "agents": res["agents"],
        "n_invoice": len(res["by_invoice"]), "n_invoice_bad": len(bad_inv),
        "n_key": len(res["by_policy"]),
        "n_key_bad": len([r for r in res["by_policy"] if r["phi_det"] != r["phi_led"]]),
        "n_missing": len(res["missing"]), "n_extra": len(res["extra"]),
        "n_unassigned": len(res["unassigned"]), "n_mismatch": len(res["mismatch"]),
        "phi_unassigned": float(sum(r["phi_det"] for r in res["unassigned"])),
        "phi_detail": float(sum(r["phi_det"] for r in res["by_policy"])),
        "phi_ledger": float(sum(r["phi_led"] for r in res["by_policy"])),
        "phi_missing": float(sum(r["phi_det"] for r in res["missing"])),
        "bad_invoices": [{"hd": r["hd"], "phi_det": float(r["phi_det"]), "phi_led": float(r["phi_led"]),
                          "lech": float(r["phi_led"] - r["phi_det"])} for r in bad_inv],
        "rows": det_rows, "led_rows": meta["rows"],
    }


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    out = argv[2] if len(argv) > 2 else os.path.join(ROOT, "Output")
    res = run(argv[0], argv[1], out)
    print("\nInvoices: %d, of which off: %d" % (res["n_invoice"], res["n_invoice_bad"]))
    print("Policy keys: %d, off: %d, not in the ledger: %d" % (res["n_key"], res["n_key_bad"], res["n_missing"]))
    print("Detail premium: %s, ledger premium: %s" % (f"{res['phi_detail']:,.0f}", f"{res['phi_ledger']:,.0f}"))
    for p in res["outputs"]:
        print("  ->", p)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
