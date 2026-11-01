# Dự án Quy trình PR trước (tách từ CCN)

Dự án riêng cho Bảng điều hành mua hàng và trạm SAP v23.1-Tram. Mã nguồn gốc lấy từ gói `QuyTrinhPRTruoc_source.zip` (31/10/2026).
Dự án dùng lại bộ **giả lập SAP GUI** của dự án CCN (`XLCongNo_refactor/sap_vba/gia_lap`), đã nâng lên v2.

| Thư mục mới | Nội dung |
|---|---|
| `gia_lap/` | Giả lập SAP v2: ME22N, MIGO từng phần, nhiều cửa sổ, DB dùng chung giữa các lệnh trạm. Đọc `gia_lap/HUONG_DAN_GIA_LAP.md` |
| `kiem-thu/` | Bộ chạy thử trong LibreOffice (PR → PO → ME22N → MIGO qua các hàm `Tram*`) và kết quả lần chạy cuối |

Thay đổi so với gói gốc: `tram-sap/SAP_Automation_v23_Tram.bas` thành bản **v23.1-Tram-gl.2**: có móc nối giả lập, các sửa lỗi của macro CCN v24
(CNRow, chặn SL/giá 0, Save xử lý hộp thoại, PBXX, gPOWarn) và sửa `SuaPO_ME22N` ghi "DA LUU" khi chưa lưu
(chi tiết trong `gia_lap/HUONG_DAN_GIA_LAP.md` mục 2 và `gia_lap/sap_automation_v23_tram.diff`). Khi không cài giả lập, macro vẫn chạy với SAP thật như cũ.
Các file còn lại giữ nguyên, kể cả dữ liệu đã ẩn danh.

---

# Quy trình PR trước – Bếp: mã nguồn

Bộ mã nguồn của Bảng điều hành mua hàng (dự trù → PR → PO → nhận hàng → đối soát → MIGO → ĐNTT)
và trạm SAP chạy trên máy công ty. Phiên bản 31/10/2026.

**Toàn bộ dữ liệu trong gói này là dữ liệu giả.** Tên nhà cung cấp, mã NCC SAP, email, số điện thoại,
mã bảo mật đã được thay. Dữ liệu thật nằm trong file PRTruoc_DuLieu (Google Sheets) và thư mục dự án trên Drive,
không có trong gói.

## Cấu trúc

| Thư mục | Nội dung |
|---|---|
| `giao-dien/tpl.html` | Mã nguồn trang Bảng điều hành (HTML + CSS + JS trong một file) |
| `giao-dien/build.py` | Dựng trang: `python3 build.py` → `dieu-hanh.html` (ghép dữ liệu bảng giá vào mẫu) |
| `giao-dien/du-lieu/` | Bảng giá rút gọn, mã NCC SAP, thông tin NCC, ghép mã kho – đều đã ẩn danh |
| `giao-dien/module/` | Các phần JS đã được ghép vào `tpl.html` (trạm SAP, agent, nạp gói) – để tham khảo, sửa riêng |
| `tram-sap/` | Trạm SAP: `TramSAP.ps1`, `ChayTramSAP.bat`, cổng Google `TramSAP_API.gs`, 3 module VBA `.bas`, hướng dẫn cài |
| `excel-mau/` | Các file Excel giữ nguyên cấu trúc (sheet, tiêu đề, định dạng, công thức, nút bấm, macro) với tối đa 30 dòng dữ liệu giả mỗi sheet |
| `cong-cu/fake_xl.py` | Công cụ tạo dữ liệu giả cho file Excel: `python3 fake_xl.py vao.xlsx ra.xlsx` |

## Dữ liệu giả trong `excel-mau/`

Mỗi sheet giữ dòng tiêu đề và mọi dòng phía trên nó; từ dòng dữ liệu đầu tiên giữ tối đa 30 dòng, phần còn lại xoá trống
(giữ định dạng ô). Sheet ít hơn 30 dòng thì giữ đúng số dòng gốc. Giá trị giả theo đúng mẫu của cột:

| Loại giá trị gốc | Giá trị giả |
|---|---|
| Số, ngày (số seri Excel) | Ngẫu nhiên trong khoảng nhỏ nhất – lớn nhất của cột, cùng số chữ số thập phân |
| Mã có chữ số (`THP00146`, `PR-T11-NV`, `2001234`) | Giữ chữ cái và ký tự nối, thay từng chữ số |
| Email | `nccNN@example.com` |
| Điện thoại, số tài khoản | Thay từng chữ số, giữ số 0 đầu và độ dài |
| Danh mục ít giá trị | Đơn vị tính, ký hiệu ngắn giữ nguyên; giá trị dài hơn đổi 1–1 sang giá trị giả cố định |
| Văn bản tự do | Từ giả, cùng số từ, cùng kiểu chữ hoa / thường |
| Công thức | Giữ nguyên; Excel tính lại khi mở |

Không thay: sheet hướng dẫn (`HƯỚNG DẪN`, `HuongDan`), tên cột khoá của sheet cấu hình (`…CauHinh`).
Link email trong ô đã bỏ.

| File | Nguồn |
|---|---|
| `Template_PR_Tram.xlsm` | File trạm SAP (VBA bản v23-Tram – import lại 3 file `.bas` trong `tram-sap/` để có bản v23.1) |
| `CongNo.xlsx` | Công nợ.xlsx |
| `GiaKhung.xlsx` | Giá khung thực phẩm.xlsx |
| `LuuMau.xlsx` | Lưu mẫu.xlsx |
| `SoluongSANoiTruT9.xlsx` | Số lượng suất ăn nội trú tháng 9 |
| `Bao_cao_Tuan1.xlsx` | Báo cáo tuần 1 – thực đơn, định lượng, đơn hàng |

## Chạy

1. `cd giao-dien && python3 build.py`, mở `dieu-hanh.html` dưới dạng Artifact có quyền Google Sheets / Drive / Gmail.
2. Cài trạm SAP theo `tram-sap/HUONG_DAN_CAI_TRAM_SAP.txt`. Mã bảo mật trong `TramSAP.config.json` và
   `TramSAP_API.gs` đang là `DAN_MA_BAO_MAT_CUA_BAN` – đặt cùng một mã mới ở cả hai file.
3. ID file Google Sheets / thư mục Drive trong `tpl.html` (hằng `FILE`, `KHO`) là của dự án gốc; đổi sang file của bạn.

## Quy tắc nghiệp vụ

- Hàng tươi giá khung theo tháng, đồ khô theo quý; một PO không vượt qua hai kỳ giá.
- Phần nhận vượt PO làm PR bổ sung mới; PR được duyệt trong ngày.
- Mọi bước chạm SAP mặc định chạy thử (không lưu) cho tới khi người dùng chuyển sang lưu thật.
