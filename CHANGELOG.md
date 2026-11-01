# Nhật ký thay đổi

## Review bước Dự trù – đợt 1 (01/11/2026)

- 1.3: bảng nguyên liệu vừa khung (cột cố định, tên hàng xuống dòng, cột xóa không còn bị cắt); trên điện thoại mỗi nguyên liệu thành thẻ. Danh sách món là nút bấm được bằng bàn phím.
- 1.2: 420 ô số suất có nhãn đọc màn hình ("CN 01/11 · Nhân viên · Trưa"); cột Ngày cố định khi cuộn ngang; hàng tiêu đề thứ hai dính đúng chỗ khi cuộn dọc.
- 1.3: ô mã hàng, định lượng có nhãn; cột xóa có tiêu đề ẩn. axe: 0 lỗi ở cả 5 mục.
- Thông báo nhỏ (toast) tự đóng khi chuyển mục, thời gian hiển thị 2,8 giây.

## Bảng thực đơn 1.1 + thanh cuộn (01/11/2026)

- Sửa lỗi gốc: ô tên hàng / tên món dùng class `wrap` trùng với khung trang nên bị padding 24px và `display:grid`, làm cột tên bị bóp. Khung trang giờ chỉ áp cho `body>.wrap`.
- Ẩn thanh cuộn ở bảng, danh sách, nhật ký và dải bước (vẫn cuộn được); bảng cuộn ngang có bóng mờ ở mép báo còn nội dung.
- Bảng thực đơn 1.1: cột cố định theo nội dung, ô nhập kiểu bảng tính (chỉ hiện viền khi rê chuột hoặc đang nhập), nút xóa chỉ hiện khi rê chuột, dòng ngoài tháng chữ mờ, thêm nhãn cho ô nhập. Trên điện thoại mỗi món thành một thẻ nhỏ, không cuộn ngang.

## Bố cục theo review (01/11/2026)

- Kết quả và thẻ số (KPI) nằm ngay dưới tiêu đề mỗi mục; khối kho chuyển xuống cuối.
- Dải 8 bước thành thanh tiến trình: chấm số nối bằng đường kẻ đi qua tâm chấm, nhãn và trạng thái bên dưới; bước đang chọn đánh dấu bằng vòng sáng quanh chấm và nhãn đậm (không còn khung thẻ). Bỏ nhóm "Trạng thái" trùng ở cột phải. Cuối mỗi mục có nút "Tiếp: …" / "Sang bước N: …".
- Header: trạng thái kết nối và trạm SAP hiển thị dạng chấm màu + chữ, tách khỏi nút; bản demo ẩn "Tải lại dữ liệu" và "Mở file dữ liệu".
- Bảng mã hàng ở 1.4 mặc định chỉ hiện dòng cần xem (có ghi chú, sửa tay, ngoài bảng giá hoặc chưa có giá), chuyển "Cần xem / Tất cả" bằng nút.

## [Chưa phát hành] – Sửa theo review UI/UX + accessibility (01/11/2026)

- Thêm `<!DOCTYPE html>`, `<html lang="vi">`, `<meta charset>` và `<meta name="viewport">` (trước đây thiếu: điện thoại thật hiển thị thu nhỏ như máy tính).
- axe-core (WCAG 2.1 AA): từ 7 loại lỗi (188 ô nhập không nhãn, select không tên, aria sai trên nút chế độ, thiếu lang, thiếu `<main>`, tương phản 3,9:1) về 0 ở mục 1.4, cả sáng / tối / điện thoại.
- Nhãn bước viết thường "Bước N" (bỏ chữ in hoa), dấu ▸ của các mục thu gọn hiện lại, vùng bấm của mục thu gọn cao 32px.
- Kiểm thử trình duyệt chặn Google Fonts để không treo khi mạng chậm.

## Redesign khoảng cách (01/11/2026)

- Hệ khoảng cách chung 4/8/12/16/24/32px (`--s1`…`--s6`); các khối trong khung làm việc cách nhau 24px (điện thoại 16px).
- Thẻ bước, thẻ số, thẻ agent: padding rộng hơn, bo 14px, bóng nhẹ; thẻ bước ghi "Bước N", bước đang chọn tô nền nhạt.
- Bảng thoáng hơn (ô 10×16px), thanh thông báo có chấm trạng thái, cột nhật ký padding 24px.
- CSS tách thành `10-base.css`, `20-gon-responsive.css`, `30-redesign.css`.

## Giao diện gọn + responsive (01/11/2026)

- Mô tả dài đầu mỗi mục thu vào nút "Cách làm"; tuỳ chọn tính dự trù thu vào "Tuỳ chọn"; chú thích nguồn dữ liệu cuối trang thu gọn.
- Dải 8 bước gọn (bỏ dòng mô tả phụ, đưa vào tooltip), cuộn ngang có snap trên màn hình nhỏ, tự cuộn tới bước / mục đang chọn.
- Khung trạng thái / nhật ký bên phải chia nhóm thu gọn được; trên màn hình nhỏ nhật ký mặc định thu gọn.
- Bảng NCC ở 1.4 hiện 6 NCC lớn nhất + "Xem cả N NCC"; điện thoại ẩn cột Kỳ 1 / Kỳ 2.
- Sửa tràn ngang: lưới chính, `.stage-row`, `.two` dùng `minmax(0,1fr)`; mọi bảng nằm trong khung cuộn; header và thanh chế độ trên điện thoại.
- Kiểm thử trình duyệt thêm kiểm tra không tràn ngang ở mọi bước / mục (390px, 820px).

## VBA, kiểm thử, bảo mật, triển khai (01/11/2026)

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
