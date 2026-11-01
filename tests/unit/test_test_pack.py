"""Kiểm tra gói dữ liệu test: khớp bảng giá, đủ ngày, đủ tuần, sinh lại được giống hệt."""
import json, os, subprocess, sys, tempfile, unittest

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
PACK = os.path.join(ROOT, "samples", "test-pack", "GoiTest_T11_2026.json")
DATA = os.path.join(ROOT, "web", "data", "data_compact.json")


class TestPack(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.P = json.load(open(PACK, encoding="utf8"))
        cls.IT = {x[0] for x in json.load(open(DATA, encoding="utf8"))["items"]}

    def test_codes_in_price_table(self):
        for d in self.P["dinhmuc"].values():
            for n in d["nl"]:
                self.assertIn(n["ma"], self.IT, d["mon"])
                self.assertGreater(n["q"], 0)

    def test_every_dish_has_norm(self):
        mons = {d["mon"] for d in self.P["dinhmuc"].values()}
        for doc in self.P["docs"].values():
            for it in doc["items"]:
                self.assertIn(it["mon"], mons)

    def test_guest_days(self):
        days = self.P["khach"]["days"]
        self.assertEqual(len(days), 30)
        for v in days.values():
            self.assertEqual(len(v), 14)
            self.assertTrue(all(isinstance(x, int) and x >= 0 for x in v))

    def test_weeks(self):
        self.assertEqual(len(self.P["docs"]), 12)  # 6 tuần × 2 nhóm

    def test_regenerates_identically(self):
        with tempfile.TemporaryDirectory() as t:
            out = os.path.join(t, "p.json")
            subprocess.run([sys.executable, "-I", os.path.join(ROOT, "tools", "make_test_pack.py"), DATA, out], check=True, capture_output=True)
            self.assertEqual(json.load(open(out, encoding="utf8")), self.P)


if __name__ == "__main__":
    unittest.main()
