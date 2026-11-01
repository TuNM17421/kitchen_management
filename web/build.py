"""Dựng trang Bảng điều hành Kitchen Management.

    python3 build.py              -> dist/dieu-hanh.html
    python3 build.py --demo       -> dist/dieu-hanh-demo.html (nhúng gói test, tự nạp khi chạy thử)

Nguồn: src/index.html (khung), src/css/*.css, src/js/*.js (ghép theo thứ tự tên file),
data/*.json (bảng giá, ghép mã kho), config/app.config.json (ID Google, không commit).
"""
import glob, json, os, sys

D = os.path.dirname(os.path.abspath(__file__))
REQ = ["dataSheetId", "khoSheetId", "seedFolderId", "xuatFolderId", "packTuan1Id"]


def read(*p):
    with open(os.path.join(D, *p), encoding="utf8") as f:
        return f.read()


def load_config():
    path = os.path.join(D, "config", "app.config.json")
    if not os.path.exists(path):
        print("! Chưa có config/app.config.json: dùng app.config.example.json (chỉ chạy thử được).", file=sys.stderr)
        path = os.path.join(D, "config", "app.config.example.json")
    with open(path, encoding="utf8") as f:
        cfg = json.load(f)
    miss = [k for k in REQ if not cfg.get(k)]
    if miss:
        sys.exit("Thiếu khoá cấu hình: " + ", ".join(miss))
    return {k: cfg[k] for k in REQ}


def build(demo=False):
    cfg = load_config()
    css = "".join(read("src", "css", os.path.basename(p)) for p in sorted(glob.glob(os.path.join(D, "src", "css", "*.css"))))
    js = "".join(read("src", "js", os.path.basename(p)) for p in sorted(glob.glob(os.path.join(D, "src", "js", "*.js"))))
    js = js.replace("__CONFIG__", json.dumps(cfg, ensure_ascii=False))
    s = read("src", "index.html")
    s = s.replace("<!--@CSS-->\n", css).replace("<!--@JS-->\n", js)
    s = s.replace("__DATA_SHEET_ID__", cfg["dataSheetId"])
    s = s.replace("__DATA__", read("data", "data_compact.json"))
    s = s.replace("__KHOMAP__", read("data", "khomap.json").strip())
    if demo:
        pack = read("..", "samples", "test-pack", "GoiTest_T11_2026.json")
        s = s.replace("<title>Bảng điều hành</title>", "<title>Bảng điều hành (demo)</title>", 1)
        s = s.replace('<script type="application/json" id="data">',
                      "<script>window.__PACK_DEMO=" + pack.replace("</", "<\\/") + ";</script>\n" + '<script type="application/json" id="data">', 1)
    for tok in ("__DATA__", "__KHOMAP__", "__CONFIG__", "__DATA_SHEET_ID__", "<!--@CSS-->", "<!--@JS-->"):
        assert tok not in s, "Còn sót " + tok
    os.makedirs(os.path.join(D, "dist"), exist_ok=True)
    out = os.path.join(D, "dist", "dieu-hanh-demo.html" if demo else "dieu-hanh.html")
    with open(out, "w", encoding="utf8") as f:
        f.write(s)
    print("Đã tạo", os.path.relpath(out, D))


if __name__ == "__main__":
    build(demo="--demo" in sys.argv)
