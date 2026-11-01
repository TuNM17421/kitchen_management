# Chạy thử giả lập SAP trong LibreOffice

Bộ chạy thử mô phỏng `TramSAP.ps1`. Mỗi lệnh mở một bản sao mới của `excel-mau/Template_PR_Tram.xlsm`,
thay các module VBA bằng bản trong `tram-sap/` và `gia_lap/`, gọi các hàm `Tram*`, rồi đóng file mà không lưu.

```
cd kiem-thu
/usr/bin/python3 chay_thu_e2e.py > ket_qua_chay_thu.txt
```

Cần LibreOffice (đã thử bản 24.2) và Python có module `uno`. File tạm nằm ở `/tmp/gltest/run`.

`xform.py` chỉ sửa **bản sao trong bộ nhớ** để LibreOffice chạy được. Các file `.bas` gốc không bị đụng tới.
- Thay `Scripting.Dictionary` bằng class `GLDict`.
- Thay `End(xlUp)` bằng `GLLastRow`.
- Thay `vbObjectError` bằng một số cố định.
- Bỏ `FormatConditions`.
- Chèn `On Error GoTo 0` trước `On Error GoTo nhãn`.

Vì vậy kết quả ở đây chưa thay cho một lần chạy trong Excel thật.
