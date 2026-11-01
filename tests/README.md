# Kiểm thử

| Loại | Lệnh | Cần | Kiểm tra gì |
|---|---|---|---|
| Unit | `npm run test:unit` (= `python3 -m unittest discover -s tests/unit`) | Python 3.10+, node (cho kiểm tra cú pháp JS) | `build.py` (không sót placeholder, chèn cấu hình, báo thiếu khoá, SRI, cú pháp JS), gói test (mã khớp bảng giá, đủ ngày/tuần, sinh lại giống hệt) |
| Trình duyệt | `npm run build && npm run test:web` | `npm ci`, `npx playwright install chromium` | Bản thường chạy thử không lỗi JS; bản demo tính dự trù đúng 10.042 suất / 684.546.740 đ / 188 mã / 18 NCC và phân bổ PR 350,4 tr / 334,1 tr |
| E2E macro | `cd tests/e2e && /usr/bin/python3 chay_thu_e2e.py > ket_qua_moi.txt && python3 so_sanh.py ket_qua_chay_thu.txt ket_qua_moi.txt` | LibreOffice + `python3-uno` | PR → PO → ME22N → MIGO qua giả lập SAP, so với kết quả chuẩn (bỏ qua dấu thời gian). Xem `e2e/README.md` |

Tất cả chạy tự động trên GitHub Actions (`.github/workflows/ci.yml`) mỗi lần push / PR, thêm 2 việc:
kiểm tra mã SRI của thư viện CDN khớp file thật trên cdnjs, và kiểm tra cú pháp các file PowerShell trên Windows.

Chưa có: chạy macro trên Excel thật và SAP thật (phải làm tay trên máy trạm, xem `sap-station/docs/`).
Số trong kiểm thử trình duyệt phụ thuộc gói `samples/test-pack/GoiTest_T11_2026.json` và bảng giá `web/data/`;
đổi một trong hai thì cập nhật số trong `tests/web/smoke.cjs`.
