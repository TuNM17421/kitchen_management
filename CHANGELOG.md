# Nhật ký thay đổi

## [Chưa phát hành] – VBA, kiểm thử, bảo mật, triển khai (01/11/2026)

- VBA: `tools/vba/Import-Vba.ps1` / `Export-Vba.ps1` đồng bộ module giữa repo và file `.xlsm`; danh sách module ở `sap-station/vba/modules.json`.
- Kiểm thử: `tests/unit` (build, gói test), `tests/web/smoke.cjs` (Playwright: số dự trù, phân bổ PR trên gói test), `tests/e2e/so_sanh.py`.
- CI GitHub Actions: build, unit, trình duyệt, kiểm tra SRI với cdnjs, cú pháp PowerShell (Windows), e2e LibreOffice.
- Bảo mật:
  - Thư viện `xlsx` từ CDN có `integrity` (SRI) + `crossorigin`.
  - `TramSAP_API.gs`: bắt buộc TOKEN ≥ 32 ký tự, so token theo thời gian không đổi, giới hạn kích thước file tải lên.
  - `ChayTramSAP.bat` chạy với `ExecutionPolicy RemoteSigned` thay cho `Bypass`; `install.ps1` bỏ chặn file tải về.
- Triển khai: `sap-station/install.ps1` (cài / cập nhật C:\TramSAP, tạo cấu hình, nạp VBA), `tools/package_station.py` (gói zip có VERSION).
- Tài liệu, thư viện: `requirements-dev.txt`, `package.json` + lock, `tests/README.md`; hướng dẫn cài trạm thêm cách cài nhanh.

## Migrate sang cấu trúc Kitchen Management (01/11/2026)

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
