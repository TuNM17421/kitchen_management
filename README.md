# Kitchen Management

Bảng điều hành mua hàng cho Bếp (dự trù → PR → PO → nhận hàng → đối soát → MIGO → ĐNTT)
và trạm SAP chạy trên máy công ty. Trước đây là dự án **QuyTrinhPRTruoc** (bản 31/10/2026, macro `v23.1-Tram-gl.2`).

**Toàn bộ dữ liệu trong repo là dữ liệu giả.** Dữ liệu thật nằm trong Google Sheets `PRTruoc_DuLieu` và thư mục dự án trên Drive.

## Cấu trúc

| Thư mục | Nội dung |
|---|---|
| `web/src/index.html` | Khung trang (HTML) |
| `web/src/css/*.css` | CSS ghép theo thứ tự tên file: `10-base` (màu, thành phần), `20-gon-responsive`, `30-redesign` (hệ khoảng cách) |
| `web/src/js/*.js` | JS chia theo phần, ghép theo thứ tự tên file: lõi, tính dự trù, từng bước 1–8, trạm SAP, agent, nạp gói, khởi động |
| `web/data/` | Bảng giá rút gọn, ghép mã kho, mã NCC SAP, thông tin NCC (ẩn danh) |
| `web/config/` | `app.config.example.json`: mẫu cấu hình ID Google. File thật `app.config.json` không commit |
| `web/build.py` | Ghép nguồn + dữ liệu + cấu hình thành `web/dist/dieu-hanh.html` |
| `sap-station/` | Trạm SAP: `TramSAP.ps1`, `ChayTramSAP.bat`, `TramSAP.config.example.json`, `install.ps1` (cài / cập nhật) |
| `sap-station/vba/` | 3 module VBA: `SAP_Automation.bas`, `MIGO_TungPhan.bas`, `DNTT_Automation.bas`; `modules.json` liệt kê module của từng file Excel |
| `sap-station/apps-script/` | Cổng Google `TramSAP_API.gs` (token, thư mục đọc từ Thuộc tính tập lệnh) |
| `sap-station/docs/` | Hướng dẫn cài trạm |
| `sap-simulator/` | Giả lập SAP GUI v2 bằng VBA để chạy macro không cần SAP thật |
| `tests/` | `unit/` (Python), `web/` (Playwright), `e2e/` (macro qua giả lập SAP trong LibreOffice). Xem `tests/README.md` |
| `samples/excel/` | File Excel mẫu (tối đa 30 dòng giả / sheet) |
| `samples/test-pack/` | Gói dữ liệu test tháng 11/2026 (thực đơn, định mức, lượng khách) |
| `tools/` | `vba/Import-Vba.ps1`, `vba/Export-Vba.ps1` (đồng bộ VBA ↔ Excel), `package_station.py` (đóng gói trạm), `fake_xl.py`, `make_test_pack.py` |
| `.github/workflows/ci.yml` | CI: build, unit, trình duyệt, SRI, cú pháp PowerShell, e2e |
| `package.json`, `requirements-dev.txt` | Thư viện cho phát triển / kiểm thử |
| `docs/history/` | Tài liệu lịch sử (diff macro trước khi có git) |

Phiên bản không còn ghi trong tên file. Phiên bản macro nằm ở hằng `MACRO_VERSION` trong `SAP_Automation.bas`,
phiên bản phát hành là tag git (xem `CHANGELOG.md`).

## Cấu hình

| Thành phần | File cấu hình (không commit) | Lấy từ mẫu |
|---|---|---|
| Trang web | `web/config/app.config.json`: `dataSheetId`, `khoSheetId`, `seedFolderId`, `xuatFolderId`, `packTuan1Id` | `web/config/app.config.example.json` |
| Trạm SAP | `sap-station/TramSAP.config.json` (đặt cạnh `TramSAP.ps1` trên máy SAP): `url`, `token`, `fileExcel`… | `sap-station/TramSAP.config.example.json` |
| Cổng Google | Apps Script → Cài đặt dự án → Thuộc tính tập lệnh: `TOKEN`, `XUAT_FOLDER` | – |

`token` của trạm và `TOKEN` của Apps Script phải giống nhau. Nên dùng chuỗi ngẫu nhiên ≥ 32 ký tự.

## Chạy

```
cd web
cp config/app.config.example.json config/app.config.json   # rồi điền ID Google
python3 build.py            # -> dist/dieu-hanh.html
python3 build.py --demo     # -> dist/dieu-hanh-demo.html: nhúng gói test, tự nạp khi chạy thử
```

Mở `dist/dieu-hanh.html` dưới dạng Artifact trên claude.ai có quyền Google Sheets / Drive / Gmail.
Không có `app.config.json` thì build dùng file mẫu và trang chỉ chạy ở chế độ chạy thử.

Trạm SAP: làm theo `sap-station/docs/HUONG_DAN_CAI_TRAM_SAP.txt` (cách nhanh: `sap-station/install.ps1`).
Giả lập SAP: `sap-simulator/HUONG_DAN_GIA_LAP.md`. Chạy thử: `tests/e2e/README.md`.

## Phát triển

```
pip install -r requirements-dev.txt
npm ci && npx playwright install chromium
npm test                    # build + unit + kiểm thử trình duyệt
```

**VBA:** file `.bas`/`.cls` trong repo là nguồn gốc, file `.xlsm` chỉ là nơi chạy.
- Nạp vào Excel: `tools\vba\Import-Vba.ps1 -Workbook <file.xlsm> [-WithSimulator]`
- Sửa trong Excel rồi đưa về repo: `tools\vba\Export-Vba.ps1 -Workbook <file.xlsm>`, xem `git diff`, commit.
- Cả hai cần Excel bật "Trust access to the VBA project object model".

**Phát hành trạm SAP:** sửa `MACRO_VERSION` nếu đổi macro, ghi `CHANGELOG.md`, gắn tag git, rồi
`python3 tools/package_station.py` → `dist/TramSAP-<tag>.zip`. Trên máy SAP: giải nén, chạy `sap-station\install.ps1`.

## Quy tắc nghiệp vụ

- Hàng tươi giá khung theo tháng, đồ khô theo quý; một PO không vượt qua hai kỳ giá.
- Phần nhận vượt PO làm PR bổ sung mới; PR được duyệt trong ngày.
- Mọi bước chạm SAP mặc định chạy thử (không lưu) cho tới khi người dùng chuyển sang lưu thật.

## Dữ liệu giả trong `samples/excel/`

Mỗi sheet giữ dòng tiêu đề và mọi dòng phía trên nó; từ dòng dữ liệu đầu tiên giữ tối đa 30 dòng. Giá trị giả theo đúng mẫu của cột:
số/ngày ngẫu nhiên trong khoảng của cột; mã có chữ số (`THP00146`, `PR-T11-NV`, `2001234`) giữ chữ, thay từng chữ số;
email `nccNN@example.com`; điện thoại, số tài khoản thay từng chữ số; văn bản tự do thành từ giả; công thức giữ nguyên.
Tạo lại bằng `python3 tools/fake_xl.py vao.xlsx ra.xlsx`.
