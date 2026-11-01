# SAP GIẢ LẬP v2 cho Quy trình PR trước — hướng dẫn

Bộ giả lập SAP GUI lấy từ dự án CCN (`XLCongNo_refactor/sap_vba/gia_lap`, bản v1 30/10/2026) và mở rộng
cho trạm SAP v23.1-Tram. Macro vẫn gọi `session.findById(...)`, `.Text`, `.press`, `sendVKey`... như với SAP thật.
Các lệnh đó đi vào một "SAP GUI giả" dựng lại đúng màn hình và ID control mà macro dùng.

Chạy được không cần SAP: `CreatePR` (ME51N), `CreatePO` (ME21N), `SuaPO_ME22N` (ME22N),
`PostMIGO_TungPhan` (MIGO từng phần) và toàn bộ các hàm `Tram*` mà `TramSAP.ps1` gọi.

## 1. Cài đặt (một lần)

1. Mở `Template_PR_Tram.xlsm` (bản dành cho trạm), nhấn Alt+F11.
2. Thay module `SAP_Automation` bằng `tram-sap/SAP_Automation_v23_Tram.bas` của dự án này
   (bản **v23.1-Tram-gl.2**, đã có móc nối giả lập và các sửa lỗi của CCN v24). Import `tram-sap/MIGO_TungPhan_v2_Tram.bas` nếu chưa có.
3. File > Import File: 4 file trong thư mục này: `SapGiaLap.bas`, `SapGL_Ctl.cls`, `SapGL_List.cls`, `SapGL_Dong.cls`.
4. Alt+F8 > `SapGiaLap_CaiDat` > Run. Macro tạo 3 sheet `SAP_GiaLap`, `SAP_GiaLap_Log`, `SAP_GiaLap_DB` và **BẬT** giả lập.
5. **Nếu chạy qua trạm:** gõ tên file vào ô **B15** của sheet `SAP_GiaLap`, ví dụ `SAP_GiaLap_DB.xlsx`
   (tên không có thư mục thì file nằm cạnh file Excel). Ctrl+S.
6. Kiểm tra: Alt+F8 > `PhienBan` phải hiện `v23.1-Tram-gl.2 (31/10/2026)`. Trạm vẫn nhận vì bản này bắt đầu bằng `v23.1`.

Quay về SAP thật: chạy `SapGiaLap_Tat` hoặc gõ `TAT` vào ô B3. Nếu xoá hẳn module `SapGiaLap` thì
macro cũng tự dùng SAP thật như cũ.

### Vì sao cần ô B15 khi chạy qua trạm

`TramSAP.ps1` mở `Template_PR_Tram.xlsm` mới cho mỗi lệnh và **không lưu** file đó (chỉ `SaveCopyAs` vào `KetQua\`).
Nếu chứng từ giả chỉ nằm trong sheet `SAP_GiaLap_DB` thì lệnh PO sau không thấy PR vừa tạo,
còn lệnh MIGO thì không thấy PO.

Khi B15 có tên file, giả lập làm như sau:
- Lúc macro kết nối SAP, giả lập chép file DB về sheet `SAP_GiaLap_DB`.
- Sau mỗi lần Save hoặc Post, giả lập ghi các chứng từ vừa đổi ra file DB.
- Nếu một Excel khác đang ghi file, giả lập chờ tối đa 30 giây.

Để trống B15 thì giả lập dùng sheet như bản v1. Cách này phù hợp khi chạy tay bằng Alt+F8 trong cùng một file.

## 2. Macro gốc được sửa những gì (`sap_automation_v23_tram.diff`)

- `MACRO_VERSION` đổi thành `v23.1-Tram-gl.2 (31/10/2026)`.

**Các sửa lỗi lấy từ macro CCN v24** (phần chung PR/PO được thay bằng bản v24, phần `Tram*` giữ nguyên):
- `CNRow`: đọc dòng CongNo theo công thức ở cột C sheet PR. Ai chèn hoặc xoá dòng ở sheet CongNo thì macro vẫn đọc đúng dòng,
  không còn gõ giá 0 vào SAP.
- Chặn SL hoặc giá bằng 0 / không phải số trước khi nhập SAP, báo rõ dòng CongNo nào sai.
- `SavePOWithPopups` / `SaveDocWithPopups`: Save xử lý hộp thoại của SAP, cảnh báo W ở thanh trạng thái thì Enter rồi Save lại
  (tối đa 4 lần), và chỉ coi là xong khi SAP trả số chứng từ.
- PBXX: chỉ thử lại một lần, không lặp.
- `gPOWarn`: cảnh báo khi nhập giá được gom lại và hiện trong hộp hỏi trước khi Save.

**Sửa `SuaPO_ME22N` ghi "DA LUU" khi chưa lưu:** Save dùng `SaveDocWithPopups`. Chỉ ghi `DA LUU` khi SAP báo đã lưu đúng số PO
(hoặc báo `No data was changed`, tức là SAP đã đúng sẵn). Các trường hợp khác ghi `LOI ...` kèm thông báo của SAP.
Khi chạy qua trạm, hộp hỏi "vẫn lưu khi còn dòng lỗi" được trả lời No.

**Móc nối giả lập:**
- `GetSapSession` có thêm một đoạn ở đầu hàm: gọi `Application.Run("SapGiaLap_PhienNeuBat", TRAM_PHIEN)`.
  Nếu giả lập đang BẬT, hàm trả về session giả của cửa sổ số `TRAM_PHIEN`. Nếu không, macro chạy tiếp như cũ.
- `TramDamBaoPhien`: khi giả lập BẬT, hàm trả về đủ n cửa sổ và không gọi `GetObject("SAPGUI")`.
- Thêm `SapSelect`, `SapType`, `SapClose`, giống bản CCN. VBA không cho class tự viết có hàm tên `Select`/`Type`/`Close`,
  nên 5 chỗ được đổi: 3 chỗ `tabAcc.Select`, 1 chỗ `tabC.Select` và 1 chỗ `ch.Type`.
  Với SAP thật, các hàm này vẫn gọi đúng `.Select` và `.Type`.

`MIGO_TungPhan_v2_Tram.bas` và `DNTT_Automation_v6_Tram.bas` **không sửa**.

## 3. Có gì mới ở v2

| Phần | Giả lập làm gì |
|---|---|
| ME22N | `/nme22n`, Shift+F5 (vKey 17) mở hộp `wnd[1]` có ô `MEPO_SELECT-EBELN`, Enter thì mở PO.<br>Ctrl+F1 (vKey 25) chuyển giữa hiển thị và sửa.<br>Sửa được SL/giá dòng có sẵn. Dòng trống thì gõ được số PR + dòng PR, SAP tự lấy mã, SL, giá như ME21N.<br>Save ghi lại DB và báo `Expense PO 45... changed`. Không sửa gì mà Save thì báo `No data was changed`.<br>Giảm SL xuống dưới mức đã nhập kho thì báo W. |
| MIGO header | Có `ctxtGOHEAD-BLDAT`, `ctxtGOHEAD-BUDAT`, `txtGOHEAD-LFSNR` (tab General). Mặc định là ngày hôm nay. Ngày sai dạng dd.mm.yyyy thì báo E. |
| MIGO từng phần | Nhập được một phần SL, bỏ tích các dòng không có trong phiếu.<br>Nhận vượt SL còn mở thì báo `PO quantity exceeded by ...` (E). SL ≤ 0 trên dòng đã tích thì báo E.<br>Nút Check (`tbar[1]/btn[7]`) trả `Document is O.K.` hoặc báo lỗi. |
| Cửa sổ | Có `iconify` / `restore` / `maximize`, chỉ ghi log. |
| Nhiều cửa sổ SAP | Cửa sổ thứ k là `ses[k]`. Số chứng từ của cửa sổ k cộng thêm k × 1.000.000 (cửa sổ 2 có MIGO `5001000001`...), nên các Excel chạy song song không bị trùng số. |
| Kịch bản mới | Dòng 33 `MO_PO / XEM`: PO mở ở chế độ hiển thị.<br>Dòng 34 `MO_PO / KHOA`: PO đang bị user khác sửa.<br>Dòng 35 `LUU ME22N / W`: cảnh báo khi Save.<br>Dòng 36 `LUU ME22N / E`: lỗi khi Save. |

Các kịch bản cũ (dòng 19–32) giữ nguyên như bản v1.
Bảng kịch bản và ý nghĩa từng cột vẫn nằm ở sheet `SAP_GiaLap`.

## 4. Đã chạy thử những gì

Đã chạy **trong LibreOffice ở chế độ VBA**, **chưa chạy trong Excel thật**. Chạy lại bằng `kiem-thu/chay_thu_e2e.py`.
Kết quả lần chạy cuối: `kiem-thu/ket_qua_chay_thu.txt`.

Mỗi lệnh mở một bản sao **mới** của `excel-mau/Template_PR_Tram.xlsm` rồi đóng mà không lưu, giống trạm.
Các lệnh dùng chung một file DB (B15).

| # | Lệnh | Kết quả |
|---|---|---|
| 1 | TramChayPR: 4 mặt hàng, 2 NCC | PR `1000000001`, item 10–40 |
| 2 | TramChayPO (lệnh mới, đọc PR từ file DB) | PO `4500000001` (NCC A) và `4500000002` (NCC B), "KHOP CONGNO" |
| 3 | TramChayPR: PR bổ sung | PR `1000000002` |
| 4a | SuaPO_ME22N, không lưu: SUA dòng 10 thành 12, THEM PR bổ sung | "DA DIEN, CHUA LUU", dòng mới là 30, DB không đổi |
| 4b | Như 4a, lưu thật, PO mở ở chế độ hiển thị | Macro bấm Ctrl+F1, Save: "Expense PO 4500000001 changed". DB: dòng 10 = 12, có thêm dòng 30 |
| 4c | PO đang bị khoá | Macro dừng: "LOI ... being processed by user GIALAP2" |
| 4d | Save gặp cảnh báo W | Macro Enter qua cảnh báo, Save lại: "DA LUU: Expense PO 4500000002 changed", DB đổi thật (trước khi sửa: ghi DA LUU dù chưa lưu) |
| 4e | Save gặp lỗi E | "LOI ... Enter G/L account for item 00010", DB không đổi |
| 5a | MIGO_TP, không lưu | "DA DIEN + CHECK, CHUA POST", SAP: Document is O.K. |
| 5b | MIGO lưu thật: dòng 10 nhận 4/12, dòng 20 nhận đủ, bỏ dòng 30 | `5000000001` posted, SL_da_nhap 4 / 5,5 / 0 |
| 5c | Nhận 9 trong khi chỉ còn 8 | Macro dừng: "PO quantity exceeded by 1,000 KG" |
| 5d | Ngày `2026-10-26` | Macro dừng: "Document date ... is not valid" |
| 5e | 2 phiếu trên cửa sổ SAP thứ 2 | `5001000001`, `5001000002`. PO nhận đủ 12 / 5,5 / 4 |
| 7a, 7b | PR rồi PO khi sheet CongNo bị chèn 1 dòng ở trên (dữ liệu lệch xuống 1 dòng so với sheet PR) | Đọc đúng dòng: PR đủ 4 mặt hàng, PO "KHOP CONGNO" đúng SL/giá |
| 6 | TramDamBaoPhien(3) | 3 |

Khi chạy trong LibreOffice, bộ chạy thử có sửa tạm một số chỗ trên **bản sao** module. Các file `.bas` giao cho bạn không bị sửa.
- Thay `Scripting.Dictionary` bằng một class tự viết.
- Thay `End(xlUp)` bằng một hàm quét dòng.
- Thay `vbObjectError` bằng một số cố định.
- Bỏ `FormatConditions`.
- Chèn `On Error GoTo 0` trước `On Error GoTo nhãn`, vì LibreOffice không bắt lại được lỗi sau `Resume Next`.

## 5. Phát hiện trong macro

1. Đã sửa (31/10/2026, bản gl.2): `SuaPO_ME22N` ghi "DA LUU" cả khi PO chưa được lưu. Xem mục 2.
2. `SuaPO_ME22N` chỉ bấm Ctrl+F1 khi tiêu đề cửa sổ có chữ `DISPLAY` hoặc `HIEN THI`. Giả lập tạo tiêu đề như vậy để thử nhánh này,
   nhưng **chưa kiểm chứng tiêu đề ME22N trên SAP thật**. Nếu SAP thật không có chữ đó, macro sẽ không chuyển sang chế độ sửa.
   Khi đó macro dừng với lỗi "ô không cho sửa", tức là dừng an toàn chứ không âm thầm sai.
3. MIGO: khi lỗi ngày ở header, cột Status ghi bước là `dong MIGO n` (tên bước cũ). Chỉ là nhãn, không ảnh hưởng kết quả.

## 6. Giới hạn

- Đây là mô phỏng **theo những gì macro dùng**, không phải SAP đầy đủ. Không có kiểm tra thật về ngân sách,
  master data hay kỳ kế toán. Các lỗi đó chỉ xảy ra khi bạn bật kịch bản.
- Câu chữ thông báo lấy từ SAP tiếng Anh chuẩn, có thể khác hệ thống công ty. Bạn sửa được ở cột F của bảng kịch bản.
- Đường dẫn subscreen ME22N (0013 khi sửa, 0019 khi hiển thị) và header MIGO là **suy đoán**.
  Macro tìm control theo tên nên vẫn chạy được, nhưng chưa đối chiếu với file record của SAP thật.
- Bản v2 dùng thay được bản v1 trong dự án CCN: hàm móc nối chỉ thêm một tham số tuỳ chọn.
