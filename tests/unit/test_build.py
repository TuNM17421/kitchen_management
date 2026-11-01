"""Kiểm tra web/build.py: chạy được, không sót placeholder, chèn đúng cấu hình, JS hợp lệ cú pháp."""
import json, os, re, shutil, subprocess, sys, tempfile, unittest

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
WEB = os.path.join(ROOT, "web")


def build(tmp, demo=False, config=None):
    """Chép web/ + samples/test-pack sang thư mục tạm rồi build, để không đụng dist/ và config thật."""
    shutil.copytree(WEB, os.path.join(tmp, "web"), ignore=shutil.ignore_patterns("dist", "app.config.json"))
    shutil.copytree(os.path.join(ROOT, "samples", "test-pack"), os.path.join(tmp, "samples", "test-pack"))
    if config is not None:
        with open(os.path.join(tmp, "web", "config", "app.config.json"), "w", encoding="utf8") as f:
            json.dump(config, f)
    cmd = [sys.executable, "-I", os.path.join(tmp, "web", "build.py")] + (["--demo"] if demo else [])
    r = subprocess.run(cmd, capture_output=True, text=True)
    out = os.path.join(tmp, "web", "dist", "dieu-hanh-demo.html" if demo else "dieu-hanh.html")
    return r, (open(out, encoding="utf8").read() if r.returncode == 0 else "")


class BuildTest(unittest.TestCase):
    CFG = {"dataSheetId": "SHEET_X", "khoSheetId": "KHO_X", "seedFolderId": "SEED_X", "xuatFolderId": "XUAT_X", "packTuan1Id": "PACK_X"}

    def test_example_config_builds(self):
        with tempfile.TemporaryDirectory() as t:
            r, html = build(t)
            self.assertEqual(r.returncode, 0, r.stderr)
            self.assertIn("app.config.example.json", r.stderr)  # cảnh báo dùng file mẫu
            for tok in ("__DATA__", "__KHOMAP__", "__CONFIG__", "__DATA_SHEET_ID__", "<!--@CSS-->", "<!--@JS-->"):
                self.assertNotIn(tok, html)

    def test_config_injected(self):
        with tempfile.TemporaryDirectory() as t:
            r, html = build(t, config=self.CFG)
            self.assertEqual(r.returncode, 0, r.stderr)
            self.assertIn('"dataSheetId": "SHEET_X"', html)
            self.assertIn("spreadsheets/d/SHEET_X/edit", html)
            self.assertNotIn("1KZTF7JIi0YSEohjlcviMyCXdMU_LL_IOUW6hUNI1B5c", html)  # không còn ID cứng của dự án gốc

    def test_missing_key_fails(self):
        cfg = dict(self.CFG); del cfg["khoSheetId"]
        with tempfile.TemporaryDirectory() as t:
            r, _ = build(t, config=cfg)
            self.assertNotEqual(r.returncode, 0)
            self.assertIn("khoSheetId", r.stderr)

    def test_demo_embeds_pack(self):
        with tempfile.TemporaryDirectory() as t:
            r, html = build(t, demo=True)
            self.assertEqual(r.returncode, 0, r.stderr)
            self.assertIn("window.__PACK_DEMO=", html)

    def test_cdn_has_sri(self):
        with tempfile.TemporaryDirectory() as t:
            _, html = build(t)
            for tag in re.findall(r"<script[^>]+src=[^>]+>", html):
                self.assertIn('integrity="sha384-', tag)
                self.assertIn('crossorigin="anonymous"', tag)

    @unittest.skipUnless(shutil.which("node"), "cần node")
    def test_js_syntax(self):
        with tempfile.TemporaryDirectory() as t:
            _, html = build(t)
            scripts = [m.group(2) for m in re.finditer(r"<script([^>]*)>(.*?)</script>", html, re.S) if "src=" not in m.group(1) and "json" not in m.group(1)]
            js = os.path.join(t, "main.js")
            with open(js, "w", encoding="utf8") as f:
                f.write("\n;\n".join(scripts))
            r = subprocess.run(["node", "--check", js], capture_output=True, text=True)
            self.assertEqual(r.returncode, 0, r.stderr)


if __name__ == "__main__":
    unittest.main()
