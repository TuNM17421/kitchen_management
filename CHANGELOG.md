# Nhật ký thay đổi

## [Chưa phát hành] – migrate sang cấu trúc Kitchen Management (01/11/2026)

Hành vi trang và macro không đổi. Đã kiểm tra:
- `web/dist/dieu-hanh.html` dựng ra giống hệt bản cũ, chỉ khác chỗ ID Google được đọc từ cấu hình.
- Bộ chạy thử `tests/e2e` cho kết quả giống hệt lần chạy 31/10/2026 (bỏ qua dấu thời gian).

Thay đổi:
- Đưa vào git; bản gốc giữ ở tag `v23.1-tram-gl.2`.
- Cấu trúc thư mục chuẩn: `web/`, `sap-station/`, `sap-simulator/`, `tests/`, `samples/`, `tools/`, `docs/`.
- Bỏ phiên bản khỏi tên file (`SAP_Automation_v23_Tram.bas` → `vba/SAP_Automation.bas`, `tram_ui2.js`… → `src/js/93-tram-po-migo.js`…).
- `tpl.html` (2.155 dòng) tách thành `src/index.html` + `src/css/app.css` + 17 file `src/js/*.js`; bỏ thư mục `module/` chép tay.
- ID Google (`FILE`, `KHO`, gói tuần 1, link mở file) chuyển sang `web/config/app.config.json`.
- `TramSAP_API.gs`: `TOKEN` và `XUAT_FOLDER` đọc từ Thuộc tính tập lệnh, từ chối khi chưa cấu hình.
- `TramSAP.ps1`: báo rõ khi thiếu `TramSAP.config.json` hoặc token còn là giá trị mẫu.
- `build.py --demo` dựng bản demo có gói test tháng 11/2026 tự nạp.
- Thêm `samples/test-pack/`, `tools/make_test_pack.py`, `.gitignore`, `.gitattributes` (giữ CRLF cho VBA/PowerShell).

## v23.1-tram-gl.2 – 31/10/2026

Bản QuyTrinhPRTruoc gốc. Chi tiết sửa macro so với v23-Tram: `sap-simulator/HUONG_DAN_GIA_LAP.md` mục 2 và `docs/history/sap_automation_v23_tram.diff`.
