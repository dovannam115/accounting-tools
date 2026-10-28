"""Doc so tien thanh chu tieng Viet."""

CHU_SO = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]


def _doc3(so, day_du, ngan="nghìn"):
    tram, chuc, dv = so // 100, (so // 10) % 10, so % 10
    out = []
    if tram or day_du:
        out.append(f"{CHU_SO[tram]} trăm")
    if chuc == 0 and dv and (tram or day_du):
        out.append("lẻ")
    elif chuc == 1:
        out.append("mười")
    elif chuc > 1:
        out.append(f"{CHU_SO[chuc]} mươi")
    if dv:
        if dv == 1 and chuc >= 2:
            out.append("mốt")
        elif dv == 5 and chuc >= 1:
            out.append("lăm")
        else:
            out.append(CHU_SO[dv])
    return " ".join(out)


def doc_so(n, ngan="nghìn"):
    """1326220 -> 'Một triệu ba trăm hai mươi sáu nghìn hai trăm hai mươi'."""
    n = int(n or 0)          # cat phan le, khong lam tron — theo dung mau ke toan dang dung
    if n == 0:
        return "Không"
    am = n < 0
    n = abs(n)
    nhom, don_vi = [], ["", ngan, "triệu", "tỷ", f"{ngan} tỷ", "triệu tỷ"]
    while n:
        nhom.append(n % 1000)
        n //= 1000
    phan = []
    for i in range(len(nhom) - 1, -1, -1):
        if nhom[i] == 0:
            continue
        phan.append(_doc3(nhom[i], i != len(nhom) - 1, ngan) + (f" {don_vi[i]}" if don_vi[i] else ""))
    s = " ".join(phan).strip()
    s = ("âm " + s) if am else s
    return s[0].upper() + s[1:]


def doc_tien(n, ngan="nghìn", duoi="đồng."):
    return f"{doc_so(n, ngan)} {duoi}"
