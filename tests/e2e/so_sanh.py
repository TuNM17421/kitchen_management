"""So kết quả chạy thử e2e mới với kết quả chuẩn, bỏ qua dấu thời gian.

    python3 so_sanh.py ket_qua_chay_thu.txt ket_qua_moi.txt
"""
import difflib, re, sys

TS = re.compile(r"\d{2}/\d{2}(/\d{4})? \d{2}:[0-9A-Za-z]{2,3}")


def norm(p):
    return [TS.sub("<thoi-gian>", l.rstrip("\n")) for l in open(p, encoding="utf8")]


a, b = norm(sys.argv[1]), norm(sys.argv[2])
if a == b:
    print("Kết quả e2e khớp chuẩn (%d dòng)." % len(a))
    sys.exit(0)
sys.stdout.writelines(l + "\n" for l in difflib.unified_diff(a, b, "chuan", "moi", lineterm="", n=1))
sys.exit(1)
