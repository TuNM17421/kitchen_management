"""Tạo samples/google/khoitao.json: dữ liệu khởi tạo cho file PRTruoc_DuLieu (nút "Nạp dữ liệu khởi tạo").

Lấy từ web/data: bảng giá + mã hàng (cauhinh/gia), thông tin NCC (cauhinh/nccinfo),
mã NCC SAP (cauhinh/lifnr), gợi ý ghép mã kho (cauhinh/khomap).
Chạy: python3 tools/make_khoitao.py
"""
import json, os, time
R = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
D = lambda *p: os.path.join(R, *p)
rd = lambda n: json.load(open(D('web', 'data', n), encoding='utf-8'))
ts = int(time.time() * 1000)
gia = rd('data_compact.json')
out = {
    'cauhinh': {
        'gia': {**gia, 'luc': ts, 'nguon': 'web/data/data_compact.json'},
        'nccinfo': rd('nccinfo.json'),
        'lifnr': rd('lifnr.json'),
        'khomap': rd('khomap.json'),
    },
    'nhatky': {'k1': {'ts': ts, 'thang': '', 'buoc': '0', 'viec': f"Nạp dữ liệu khởi tạo: {len(gia['items'])} mã hàng, cấu hình NCC"}},
}
os.makedirs(D('samples', 'google'), exist_ok=True)
p = D('samples', 'google', 'khoitao.json')
json.dump(out, open(p, 'w', encoding='utf-8'), ensure_ascii=False)
print(p, os.path.getsize(p), 'bytes,', len(gia['items']), 'mã hàng')
