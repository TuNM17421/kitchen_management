"""Thay dữ liệu Excel bằng tối đa 30 dòng giả / sheet, giữ nguyên mọi thứ khác.

Sửa trực tiếp XML của từng sheet nên giữ: tên sheet, dòng tiêu đề, định dạng, độ rộng cột, công thức,
ô gộp, nút bấm / hình vẽ, macro VBA (xlsm). Dữ liệu sau dòng giả thứ 30 bị xoá (giữ định dạng ô).
Giá trị giả sinh theo "hình dạng" (regex) của giá trị gốc cùng cột:
  số / ngày (số seri Excel) -> số ngẫu nhiên trong min..max của cột, cùng số chữ số thập phân
  mã có chữ số (THP00146, PR-T11-NV, 2001234) -> giữ chữ cái và ký tự nối, thay từng chữ số
  email -> nccNN@example.com ; điện thoại, số tài khoản -> thay từng chữ số, giữ số 0 đầu
  danh mục ít giá trị (ĐVT, trạng thái, nhóm) -> chọn lại trong danh mục gốc
  văn bản tự do -> từ giả, cùng số từ, cùng kiểu chữ hoa/thường
Sheet hướng dẫn giữ nguyên; sheet cấu hình (…CauHinh) chỉ thay cột giá trị.
Cách dùng: python3 -I fake_xl.py vao.xlsx ra.xlsx
"""
import sys, re, random, zipfile, posixpath
from lxml import etree

N = 30
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
q = lambda t: "{%s}%s" % (NS, t)
random.seed(20261031)

POOL = ("Rau Cải Bí Cà Đậu Nấm Gạo Bột Thịt Cá Gà Vịt Tôm Mực Trứng Sữa Bánh Mì Bún Phở Hành Tỏi Gừng Chanh "
        "Cam Táo Lê Dưa Ngô Khoai Muối Đường Dầu Nước Mắm Tương Hạt Tiêu Sả Ớt Mướp Su Lạc Vừng Mẫu Thử Xanh Đỏ "
        "Vàng Trắng Tươi Khô Lớn Nhỏ An Bình Minh Phúc Hòa Thịnh Phát Đông Tây Nam Bắc Hồng Lan Mai Đào").split()
# Bảng ẩn danh tuỳ chọn: biến môi trường FAKE_MAP = file JSON {"names": {thật: giả}, "codes": {thật: giả}}
import os, json
_M = json.load(open(os.environ["FAKE_MAP"], encoding="utf8")) if os.environ.get("FAKE_MAP") else {"names": {}, "codes": {}}
_NRE = re.compile("|".join(re.escape(k) for k in sorted(_M["names"], key=len, reverse=True)), re.I) if _M["names"] else None
_LOW = {k.lower(): v for k, v in _M["names"].items()}


def scrub(s):
    if _NRE is not None:
        s = _NRE.sub(lambda m: _LOW[m.group(0).lower()], s)
    for k, v in _M["codes"].items():
        s = s.replace(k, v)
    return s


GUIDE = re.compile(r"hướng dẫn|huongdan|huong dan", re.I)
CONFIG = re.compile(r"cauhinh|cấu hình|cau hinh", re.I)
SEP = re.compile(r"(\s+|[,;:/()\[\]\-–.+&|_]+)")


def col_idx(ref):
    n = 0
    for ch in re.match(r"[A-Z]+", ref).group(0):
        n = n * 26 + ord(ch) - 64
    return n


def case_like(src, w):
    if src.isupper() and len(src) > 1:
        return w.upper()
    if src[:1].isupper():
        return w[:1].upper() + w[1:].lower()
    return w.lower()


def fake_digits(s):
    keep = all(len(m) <= 4 and m.isupper() for m in re.findall(r"[^\W\d_]+", s))  # THP, BS, T, PR: giữ
    out, lead = [], True
    for ch in s:
        if ch.isalpha() and not keep:
            r = random.choice("abcdeghiklmnopqrstuvxy")
            out.append(r.upper() if ch.isupper() else r); lead = False; continue
        if ch.isdigit():
            out.append("0" if lead and ch == "0" else str(random.randint(1 if lead else 0, 9)))
            lead = False
        else:
            out.append(ch)
            lead = not ch.isalnum()
    return "".join(out)


def fake_text(s):
    st = s.strip()
    if re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", st):
        return "ncc%02d@example.com" % random.randint(1, 99)
    if re.match(r"https?://", st):
        return "https://example.com/" + "".join(random.choice("abcdefghijk") for _ in range(10))
    if re.fullmatch(r"[A-Za-z]:\\.*|\\\\.*", st):            # đường dẫn Windows
        b = posixpath.basename(st.replace("\\", "/")) or "File"
        stem, ext = posixpath.splitext(b)
        return "C:\\ViDu\\" + fake_text(stem or "File") + ext
    if re.fullmatch(r"[A-Za-z0-9]{12,}", st) and re.search(r"[a-z]", st) and re.search(r"[A-Z0-9]", st):
        return "".join(random.choice("abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789") for _ in st)  # mã bảo mật
    s = re.sub(r"[^@\s;,]+@[^@\s;,]+", lambda m: "ncc%02d@example.com" % random.randint(1, 99), s)
    out = []
    for t in re.split(r"(ncc\d\d@example\.com)", s):
        if re.fullmatch(r"ncc\d\d@example\.com", t):
            out.append(t); continue
        out.extend(_words(t))
    return "".join(out)


def _words(s):
    out = []
    for t in SEP.split(s):
        if not t or SEP.fullmatch(t):
            out.append(t)
        elif re.search(r"\d", t):
            out.append(fake_digits(t))
        elif re.fullmatch(r"[A-Z]{1,4}", t) and len(s) < 40:
            out.append(t)                                         # ký hiệu ngắn: KG, PR, NV, T11
        else:
            out.append(case_like(t, random.choice(POOL)))
    return out


class Col:
    def __init__(self):
        self.nums, self.strs, self.dec, self.map = [], [], 0, {}

    def add(self, kind, v):
        if kind == "n":
            try:
                x = float(v)
            except ValueError:
                return
            self.nums.append(x)
            if "." in v and "E" not in v.upper():
                self.dec = max(self.dec, min(4, len(v.split(".")[1])))
        elif kind == "s" and v.strip():
            self.strs.append(v)

    def finish(self):
        d = set(self.strs)
        self.cat = (len(self.strs) >= 4 and len(d) <= max(3, min(12, len(self.strs) // 3))
                    and all(len(x) <= 40 for x in d))
        self.cats = sorted(d)

    def num(self, orig):
        lo, hi = (min(self.nums), max(self.nums)) if self.nums else (float(orig),) * 2
        if self.dec == 0:
            return str(random.randint(int(lo), int(hi)) if hi > lo else int(lo))
        return ("%." + str(self.dec) + "f") % random.uniform(lo, hi)

    def text(self, orig):
        if self.cat:
            pick = random.choice(self.cats)
            if len(pick) <= 6:                       # đơn vị tính, ký hiệu ngắn: giữ
                return pick
            if pick not in self.map:                 # danh mục: mỗi giá trị gốc -> 1 giá trị giả cố định
                self.map[pick] = fake_text(pick)
            return self.map[pick]
        return fake_text(random.choice(self.strs) if self.strs else orig)


def cell_val(c, sst):
    t = c.get("t")
    v = c.find(q("v"))
    if t == "s" and v is not None:
        return "s", sst[int(v.text)]
    if t == "str" and v is not None and c.find(q("f")) is None:
        return "s", v.text or ""
    if t == "inlineStr":
        return "s", "".join(x.text or "" for x in c.iter(q("t")))
    if t in (None, "n") and v is not None and c.find(q("f")) is None:
        return "n", v.text
    return None, None


INLINE = False


def set_str(c, s, new_sst):
    for ch in list(c):
        c.remove(ch)
    if INLINE:                                   # file gốc không có bảng chuỗi chung
        c.set("t", "inlineStr")
        t = etree.SubElement(etree.SubElement(c, q("is")), q("t")); t.text = scrub(s)
        return
    c.set("t", "s")
    v = etree.SubElement(c, q("v"))
    new_sst.append(scrub(s))
    v.text = str(len(new_sst) - 1)


def clear(c):
    for ch in list(c):
        c.remove(ch)
    c.attrib.pop("t", None)


def header_row(rows, sst):
    best, br = -1, 1
    for r in rows[:15]:
        vals = [cell_val(c, sst) for c in r.findall(q("c"))]
        ne = [k for k, v in vals if k]
        ss = [v for k, v in vals if k == "s"]
        if len(ne) >= 2 and len(ss) / len(ne) >= .7 and len(ss) > best:
            best, br = len(ss), int(r.get("r"))
    return br


def process(src, dst):
    zin = zipfile.ZipFile(src)
    parts = {n: zin.read(n) for n in zin.namelist()}
    wb = etree.fromstring(parts["xl/workbook.xml"])
    rels = etree.fromstring(parts["xl/_rels/workbook.xml.rels"])
    rmap = {r.get("Id"): r for r in rels}
    sst = []
    if "xl/sharedStrings.xml" in parts:
        for si in etree.fromstring(parts["xl/sharedStrings.xml"]).findall(q("si")):
            sst.append("".join(t.text or "" for t in si.iter(q("t"))))
    global INLINE
    INLINE = "xl/sharedStrings.xml" not in parts
    new_sst, keep_idx = [], {}
    report = []
    sheets = []
    for s in wb.find(q("sheets")):
        tgt = rmap[s.get("{%s}id" % RNS)].get("Target")
        path = tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join("xl", tgt))
        sheets.append((s.get("name"), path))
    for name, path in sheets:
        root = etree.fromstring(parts[path])
        sd = root.find(q("sheetData"))
        rows = sd.findall(q("row"))

        def keep_shared(c):           # ô giữ nguyên: chuyển chỉ số chuỗi sang bảng chuỗi mới (+ ẩn tên/mã NCC)
            if c.get("t") in ("inlineStr", "str") and c.find(q("f")) is None:
                for t in list(c.iter(q("t"))) + list(c.iter(q("v"))):
                    if t.text:
                        t.text = scrub(t.text)
            if c.get("t") is None and c.find(q("f")) is None and c.find(q("v")) is not None:
                c.find(q("v")).text = _M["codes"].get(c.find(q("v")).text, c.find(q("v")).text)
            if c.get("t") == "s" and c.find(q("v")) is not None:
                old = int(c.find(q("v")).text)
                if old not in keep_idx:
                    new_sst.append(scrub(sst[old])); keep_idx[old] = len(new_sst) - 1
                c.find(q("v")).text = str(keep_idx[old])

        if GUIDE.search(name) or not rows:
            for c in root.iter(q("c")):
                keep_shared(c)
            parts[path] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            report.append((name, "giữ nguyên (hướng dẫn)"))
            continue
        cfg = bool(CONFIG.search(name))
        h = header_row(rows, sst)
        if cfg:
            h = next((int(r.get("r")) for r in rows[:6] for c in r.findall(q("c"))
                      if cell_val(c, sst)[1] in ("Gia_tri", "Giá trị")), 1)
        cut = h + N
        cols = {}
        for r in rows:
            if int(r.get("r")) <= h:
                continue
            for c in r.findall(q("c")):
                k, v = cell_val(c, sst)
                if k:
                    cols.setdefault(col_idx(c.get("r")), Col()).add(k, v)
        for col in cols.values():
            col.finish()
        nfake = ndel = 0
        for r in rows:
            rn = int(r.get("r"))
            for c in r.findall(q("c")):
                ci = col_idx(c.get("r"))
                if cfg and ci != 2 and rn > h:
                    k0, v0 = cell_val(c, sst)
                    if k0 == "s" and v0.strip().isupper() and len(v0.strip()) > 3 and "_" not in v0:
                        set_str(c, fake_text(v0), new_sst); continue   # tên NCC trong bảng cấu hình
                if rn <= h or (cfg and ci != 2):
                    keep_shared(c); continue
                if rn > cut:
                    if len(c):
                        clear(c); ndel += 1
                    continue
                f = c.find(q("f"))
                if f is not None:                       # công thức: giữ, bỏ kết quả cũ
                    if f.get("t") == "shared" and f.get("ref"):
                        a, b = f.get("ref").split(":") if ":" in f.get("ref") else (f.get("ref"),) * 2
                        brow = int(re.search(r"\d+", b).group(0))
                        if brow > cut:
                            f.set("ref", a + ":" + re.sub(r"\d+", str(cut), b))
                    v = c.find(q("v"))
                    if v is not None:
                        c.remove(v)
                    if c.get("t") in ("s", "str", "e", "b"):
                        c.attrib.pop("t", None)
                    continue
                k, v = cell_val(c, sst)
                col = cols.get(ci)
                if cfg:                                   # bảng cấu hình: giả theo chính giá trị gốc
                    if k == "s" and v.strip() and len(v.strip()) > 6:
                        set_str(c, fake_text(v), new_sst); nfake += 1
                    elif k == "n":
                        c.find(q("v")).text = fake_digits(v); nfake += 1
                    else:
                        keep_shared(c)
                    continue
                if k == "n" and col:
                    nv = col.num(v)
                    c.find(q("v")).text = _M["codes"].get(nv, nv); nfake += 1
                elif k == "s" and col and v.strip():
                    set_str(c, col.text(v), new_sst); nfake += 1
                else:
                    keep_shared(c)
        hl = root.find(q("hyperlinks"))              # link email/web gốc: bỏ
        if hl is not None:
            root.remove(hl)
            rp = posixpath.join(posixpath.dirname(path), "_rels", posixpath.basename(path) + ".rels")
            if rp in parts:
                rr = etree.fromstring(parts[rp])
                for x in list(rr):
                    if x.get("Type", "").endswith("/hyperlink"):
                        rr.remove(x)
                parts[rp] = etree.tostring(rr, xml_declaration=True, encoding="UTF-8", standalone=True)
        mc = root.find(q("mergeCells"))
        if mc is not None:
            for m in list(mc):
                lastrow = int(re.findall(r"\d+", m.get("ref"))[-1])
                if lastrow > cut and int(re.findall(r"\d+", m.get("ref"))[0]) > h:
                    mc.remove(m)
            if len(mc) == 0:
                root.remove(mc)
            else:
                mc.set("count", str(len(mc)))
        parts[path] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        drows = len([r for r in rows if int(r.get("r")) > h])
        report.append((name, "tiêu đề dòng %d, %d dòng dữ liệu -> giữ %d dòng giả (%d ô thay)"
                       % (h, drows, min(drows, N), nfake)))
    # bảng chuỗi mới: chỉ còn chuỗi tiêu đề + chuỗi giả
    if not INLINE:
        root = etree.Element(q("sst"), nsmap={None: NS})
        root.set("count", str(len(new_sst))); root.set("uniqueCount", str(len(new_sst)))
        for s in new_sst:
            si = etree.SubElement(root, q("si")); t = etree.SubElement(si, q("t")); t.text = s
            if s != s.strip():
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        parts["xl/sharedStrings.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    # bỏ calcChain, bắt Excel tính lại công thức khi mở
    if "xl/calcChain.xml" in parts:
        del parts["xl/calcChain.xml"]
        for r in list(rels):
            if r.get("Target", "").endswith("calcChain.xml"):
                rels.remove(r)
        ct = etree.fromstring(parts["[Content_Types].xml"])
        for o in list(ct):
            if o.get("PartName") == "/xl/calcChain.xml":
                ct.remove(o)
        parts["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8", standalone=True)
        parts["xl/_rels/workbook.xml.rels"] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
    cp = wb.find(q("calcPr"))
    if cp is None:
        cp = etree.Element(q("calcPr"))
        after = [x for x in wb if etree.QName(x).localname in ("oleSize", "customWorkbookViews", "pivotCaches",
                 "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst")]
        (after[0].addprevious(cp) if after else wb.append(cp))
    cp.set("fullCalcOnLoad", "1")
    parts["xl/workbook.xml"] = etree.tostring(wb, xml_declaration=True, encoding="UTF-8", standalone=True)
    for dp in ("docProps/core.xml", "docProps/app.xml"):    # tên người tạo / sửa, công ty
        if dp in parts:
            t = parts[dp].decode("utf8")
            t = re.sub(r"(<(dc:creator|cp:lastModifiedBy|Company|Manager)>)[^<]*(</)", r"\1Người dùng\3", t)
            parts[dp] = t.encode("utf8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for n in zin.namelist():
            if n in parts:
                z.writestr(n, parts[n])
    return report


if __name__ == "__main__":
    for name, msg in process(sys.argv[1], sys.argv[2]):
        print("  %-28s %s" % (name, msg))
