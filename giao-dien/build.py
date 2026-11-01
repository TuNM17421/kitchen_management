"""Dựng trang Bảng điều hành: python3 build.py -> dieu-hanh.html"""
import os
D = os.path.dirname(os.path.abspath(__file__))
s = open(os.path.join(D, "tpl.html"), encoding="utf8").read()
s = s.replace("__DATA__", open(os.path.join(D, "du-lieu", "data_compact.json"), encoding="utf8").read())
s = s.replace("__KHOMAP__", open(os.path.join(D, "du-lieu", "khomap.json"), encoding="utf8").read().strip())
assert "__DATA__" not in s and "__KHOMAP__" not in s
open(os.path.join(D, "dieu-hanh.html"), "w", encoding="utf8").write(s)
print("Đã tạo dieu-hanh.html")
