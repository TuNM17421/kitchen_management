Attribute VB_Name = "DNTT_Automation"
Option Explicit

' ================================================================================
'  BUOC 4: LAP DE NGHI THANH TOAN (DNTT) TU CAC PO DA TAO
'  v6-Tram (30/10/2026): dung chung ham Hoi() cua SAP_Automation v23-Tram; chay qua tram SAP
'  thi khong hien hop thoai va khong tong hop lai sheet DNTT (giao dien da gui du lieu).
'  Module RIENG, chay doc lap voi SAP_Automation (khong can mo SAP).
'  Import vao Template_PR_v3.xlsm: Alt+F11 > File > Import File > chon file .bas nay
'
'  2 macro (Alt+F8):
'    TongHopDNTT : moi NCC co PO = 1 dong o sheet DNTT (de kiem tra / sua tay neu can)
'    TaoDNTT     : dien vao sheet mau cua tung NCC trong file DNTT.xlsx, chen chu ky,
'                  luu thanh FILE MOI (file mau goc khong bi sua) + 1 file PDF de in
'
'  Thong tin tu dien:
'    - Ngay lap DNTT = hom nay           - So tien = so tien tren PO
'    - Han thanh toan = ngay lap + 15    - Vu viec/Ghi chu = So PO
'    - So bien ban giao nhan = so bien ban cua NCC trong ky chot (doc file Cong no.xlsx,
'      sheet "Don hang chi tiet"). Ky cua PO duoc TU NHAN RA bang cach so so tien PO
'      voi tong tien don hang tung ky.
'    - Chu ky nguoi de nghi = anh chu_ky_Dat.png
'  Cau hinh: sheet DNTT_CauHinh (macro tu tao lan dau chay)
' ================================================================================

Public Const DNTT_VERSION As String = "DNTT v6-Tram (30/10/2026)"
Private Const CFG_SHEET As String = "DNTT_CauHinh"
Private Const OUT_SHEET As String = "DNTT"
Private Const MAP_FIRST_ROW As Long = 20
Private Const SIG_SHAPE As String = "ChuKy_NguoiDeNghi"
Private gBoDau As Object
Private gFso As Object


' --------------------------------------------------------------------------
' CAU HINH
' --------------------------------------------------------------------------
Private Function Fso() As Object
    If gFso Is Nothing Then Set gFso = CreateObject("Scripting.FileSystemObject")
    Set Fso = gFso
End Function

Private Function CfgSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CFG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = CFG_SHEET
        TaoCauHinhMacDinh ws
    End If
    Set CfgSheet = ws
End Function

Private Sub TaoCauHinhMacDinh(ws As Worksheet)
    ws.Range("A1").Value = "CAU HINH LAP DE NGHI THANH TOAN (DNTT) - o mau vang sua duoc"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3:C3").Value = Array("Truong", "Gia_tri", "Ghi_chu")
    ws.Range("A3:C3").Font.Bold = True
    ws.Range("A4:C4").Value = Array("File_mau_DNTT_1", "", "Duong dan day du file DNTT.xlsx. De trong = tim DNTT.xlsx o thu muc file nay, roi thu muc cha")
    ws.Range("A5:C5").Value = Array("File_mau_DNTT_2", "", "(tuy chon) File mau thu 2, vd DNTT-CCDC.xlsx. De trong = tim tuong tu, khong co thi bo qua")
    ws.Range("A6:C6").Value = Array("File_cong_no", "", "Duong dan file Cong no.xlsx (de dem so bien ban + nhan ky). De trong = tim o thu muc file nay, roi thu muc cha")
    ws.Range("A7:C7").Value = Array("File_chu_ky", "", "De trong = KHONG chen chu ky (ky tay). Muon chen thi ghi duong dan day du anh chu ky .png")
    ws.Range("A8:C8").Value = Array("Thu_muc_xuat", "", "De trong = thu muc DNTT_xuat canh file nay")
    ws.Range("A9:C9").Value = Array("Nguoi_de_nghi", ("NGUY" & ChrW(7876) & "N TH" & ChrW(192) & "NH " & ChrW(272) & ChrW(7840) & "T"), "Ghi vao o Nguoi de nghi (H11). De trong = giu nhu file mau")
    ws.Range("A10:C10").Value = Array("Dia_danh", ("H" & ChrW(432) & "ng Y" & ChrW(234) & "n"), ("Dia danh truoc chu ', ngay' (o Q9). De trong = H" & ChrW(432) & "ng Y" & ChrW(234) & "n"))
    ws.Range("A11:C11").Value = Array("So_ngay_han_TT", 15, "Han thanh toan = Ngay lap + so ngay nay")
    ws.Range("A12:C12").Value = Array("Mau_noi_dung", ("Thanh to" & ChrW(225) & "n chi ph" & ChrW(237) & " {LOAI} th" & ChrW(225) & "ng {THANG} k" & ChrW(7923) & " {KY}"), ("{LOAI} = th" & ChrW(7921) & "c ph" & ChrW(7849) & "m t" & ChrW(432) & ChrW(417) & "i s" & ChrW(7889) & "ng / " & ChrW(273) & ChrW(7891) & " kh" & ChrW(244) & " (theo mat hang trong PO); {THANG} = thang lam DNTT - 1; {KY} = ky cua du lieu PO"))
    ws.Range("A13:C13").Value = Array("Thang", "", "CHI dung khi khong tu nhan ra duoc ky tu so tien. De trong = tu tinh theo ngay hom nay")
    ws.Range("A14:C14").Value = Array("Ky", "", "Nhu tren: 1, 2 hoac 0 (= ca thang)")
    ws.Range("A15:C15").Value = Array("Xuat_PDF", "CO", "CO = xuat them 1 file PDF gom cac DNTT de in")
    ws.Range("A16:C16").Value = Array("Thang_noi_dung", "", "Thang ghi trong noi dung thanh toan. De trong = thang hien tai - 1 (vd lam DNTT thang 10 -> thang 9)")
    ws.Range("A17:C17").Value = Array("Bo_phan", ("B" & ChrW(7870) & "P"), ("Bo phan de nghi (o T11). De trong = B" & ChrW(7870) & "P"))
    ws.Range("B4:B17").Interior.Color = RGB(255, 242, 204)
    ws.Range("B11").HorizontalAlignment = xlLeft

    ws.Range("A18").Value = "BANG NCC -> SHEET MAU (Ten_NCC giong cot Nha_cung_cap o sheet PO; Sheet_mau giong ten sheet trong file DNTT). Them dong khi co NCC moi."
    ws.Range("A18").Font.Bold = True
    ws.Range("A19:C19").Value = Array("Ten_NCC", "Sheet_mau", "Ghi_chu")
    ws.Range("A19:C19").Font.Bold = True

    Dim pairs As Variant, i As Long, r As Long
    pairs = Array( _
        "NCC MAU 23", "NCC MAU 23", _
        ("TH" & ChrW(192) & "NH " & ChrW(272) & ChrW(7840) & "T"), ("TH" & ChrW(192) & "NH " & ChrW(272) & ChrW(7840) & "T"), _
        ("HO" & ChrW(7840) & "T TH" & ChrW(192) & "NH"), (ChrW(272) & "NTT HO" & ChrW(7840) & "T TH" & ChrW(192) & "NH"), _
        ("CH" & ChrW(431) & ChrW(416) & "NG M" & ChrW(7928)), (ChrW(272) & "NTT CH" & ChrW(431) & ChrW(416) & "NG M" & ChrW(7928)), _
        "VNM", "NCC MAU 30", _
        "NCC MAU 05", "NCC MAU 05", _
        "NCC MAU 02", "NCC MAU 02", _
        (ChrW(194) & "N NAM"), "NCC MAU 35", _
        (ChrW(272) & ChrW(7840) & "I T" & ChrW(194) & "N VI" & ChrW(7878) & "T"), (ChrW(272) & ChrW(7840) & "I T" & ChrW(194) & "N VI" & ChrW(7878) & "T"), _
        ("SEN V" & ChrW(192) & "NG"), ("Sen V" & ChrW(224) & "ng"), _
        "NCC MAU 21", "NCC MAU 21", _
        "T&P", "T&P", _
        "NCC MAU 24", "NCC MAU 24", _
        ("FAS VI" & ChrW(7878) & "T NAM"), "FAS", _
        "NCC MAU 33", "NCC MAU 33", _
        ("H" & ChrW(7841) & " T" & ChrW(7847) & "ng M" & ChrW(7841) & "ng"), "HTM", _
        ("Mi" & ChrW(7873) & "n th" & ChrW(7921) & "c ph" & ChrW(7849) & "m"), "MTP", _
        "NCC MAU 17", "NCC MAU 17")
    r = MAP_FIRST_ROW
    For i = LBound(pairs) To UBound(pairs) Step 2
        ws.Cells(r, 1).Value = pairs(i)
        ws.Cells(r, 2).Value = pairs(i + 1)
        r = r + 1
    Next i
    ws.Range(ws.Cells(MAP_FIRST_ROW, 1), ws.Cells(r + 10, 2)).Interior.Color = RGB(255, 242, 204)
    ws.Columns("A:B").ColumnWidth = 28
    ws.Columns("C").ColumnWidth = 100
End Sub

Private Function CfgVal(ByVal key As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = CfgSheet()
    For r = 4 To 17
        If Trim$(CStr(ws.Cells(r, 1).Value)) = key Then
            CfgVal = Trim$(CStr(ws.Cells(r, 2).Value))
            Exit Function
        End If
    Next r
End Function

Private Function FileExists(ByVal p As String) As Boolean
    If Len(p) = 0 Then Exit Function
    FileExists = Fso().FileExists(p)
End Function

' Tim file: duong dan cau hinh, neu trong thi tim theo ten o thu muc file nay roi thu muc cha
Private Function TimFile(ByVal cfgKey As String, ByVal tenMacDinh As String) As String
    Dim p As String, parent As String
    p = CfgVal(cfgKey)
    If p <> "" Then
        If FileExists(p) Then TimFile = p
        Exit Function
    End If
    p = ThisWorkbook.Path & "\" & tenMacDinh
    If FileExists(p) Then TimFile = p: Exit Function
    parent = Fso().GetParentFolderName(ThisWorkbook.Path)
    If parent <> "" Then
        p = parent & "\" & tenMacDinh
        If FileExists(p) Then TimFile = p: Exit Function
    End If
End Function

Private Function TemplatePaths() As Collection
    Dim c As New Collection, p As String
    p = TimFile("File_mau_DNTT_1", "DNTT.xlsx")
    If p <> "" Then c.Add p
    p = TimFile("File_mau_DNTT_2", "DNTT-CCDC.xlsx")
    If p <> "" Then c.Add p
    Set TemplatePaths = c
End Function

Private Function OutFolder() As String
    Dim p As String
    p = CfgVal("Thu_muc_xuat")
    If p = "" Then p = ThisWorkbook.Path & "\DNTT_xuat"
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    If Not Fso().FolderExists(p) Then Fso().CreateFolder p
    OutFolder = p
End Function

Private Function SoNgayHan() As Long
    Dim s As String
    s = CfgVal("So_ngay_han_TT")
    If IsNumeric(s) Then SoNgayHan = CLng(s) Else SoNgayHan = 15
End Function

' Ky du phong khi khong tu nhan ra: ngay 16-cuoi thang -> ky 1 thang nay; ngay 1-15 -> ky 2 thang truoc
Private Sub KyDuPhong(ByRef thang As Long, ByRef nam As Long, ByRef ky As Long)
    Dim s As String, d As Date
    d = Date
    If Day(d) >= 16 Then
        thang = Month(d): nam = Year(d): ky = 1
    Else
        d = DateSerial(Year(d), Month(d), 0)
        thang = Month(d): nam = Year(d): ky = 2
    End If
    s = CfgVal("Thang")
    If IsNumeric(s) Then thang = CLng(s)
    s = CfgVal("Ky")
    If IsNumeric(s) Then ky = CLng(s)
End Sub

' Thang trong noi dung = thang hien tai - 1 (thang 1 -> thang 12), tru khi ghi o DNTT_CauHinh!Thang_noi_dung
Private Function ThangNoiDung() As Long
    Dim s As String
    s = CfgVal("Thang_noi_dung")
    If IsNumeric(s) Then ThangNoiDung = CLng(s): Exit Function
    ThangNoiDung = Month(DateSerial(Year(Date), Month(Date), 0))
End Function

Private Function NoiDung(ByVal loai As String, ByVal thang As Long, ByVal ky As Long) As String
    Dim t As String
    t = CfgVal("Mau_noi_dung")
    If t = "" Or InStr(t, "{LOAI}") = 0 Then t = ("Thanh to" & ChrW(225) & "n chi ph" & ChrW(237) & " {LOAI} th" & ChrW(225) & "ng {THANG} k" & ChrW(7923) & " {KY}")
    t = Replace(t, "{LOAI}", loai)
    If ky = 0 Then
        t = Replace(t, (" k" & ChrW(7923) & " {KY}"), "")
        t = Replace(t, " ky {KY}", "")
        t = Replace(t, "{KY}", "")
    End If
    t = Replace(t, "{THANG}", CStr(thang))
    t = Replace(t, "{KY}", CStr(ky))
    NoiDung = Trim$(t)
End Function


' --------------------------------------------------------------------------
' CHUAN HOA TEN DE SO KHOP (bo dau, viet hoa, bo chu DNTT, chi giu chu + so)
' --------------------------------------------------------------------------
Private Function BoDau(ByVal s As String) As String
    Dim i As Long, code As Long, out As String
    If gBoDau Is Nothing Then InitBoDau
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        If code < 0 Then code = code + 65536
        If gBoDau.Exists(code) Then
            out = out & gBoDau(code)
        Else
            out = out & Mid$(s, i, 1)
        End If
    Next i
    BoDau = out
End Function

Private Function KhoaSo(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = UCase$(BoDau(Trim$(s)))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then out = out & ch
    Next i
    If Left$(out, 4) = "DNTT" And Len(out) > 4 Then out = Mid$(out, 5)
    KhoaSo = out
End Function

Private Function NoiChuoi(ByVal s As String, ByVal x As String) As String
    If x = "" Then
        NoiChuoi = s
    ElseIf s = "" Then
        NoiChuoi = x
    ElseIf InStr(", " & s & ", ", ", " & x & ", ") > 0 Then
        NoiChuoi = s
    Else
        NoiChuoi = s & ", " & x
    End If
End Function


' --------------------------------------------------------------------------
' MO FILE (chi doc) / TIM SHEET MAU CUA NCC
' --------------------------------------------------------------------------
Private Function WbDangMo(ByVal p As String) As Workbook
    Dim wb As Workbook, nm As String
    nm = Fso().GetFileName(p)
    For Each wb In Application.Workbooks
        If LCase$(wb.FullName) = LCase$(p) Or LCase$(wb.Name) = LCase$(nm) Then
            Set WbDangMo = wb
            Exit Function
        End If
    Next wb
End Function

Private Function MoChiDoc(ByVal p As String) As Workbook
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Set MoChiDoc = Workbooks.Open(Filename:=p, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Function

Private Function SheetTonTai(wb As Workbook, ByVal nm As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Sheets(nm)
    On Error GoTo 0
    SheetTonTai = Not sh Is Nothing
End Function

' Tra ve ten sheet mau; fileIdx = so thu tu file mau (1, 2) chua sheet do (0 = khong thay)
Private Function TimSheetMau(ByVal ncc As String, wbs As Collection, ByRef fileIdx As Long) As String
    Dim wsC As Worksheet, r As Long, k As String, nm As String, i As Long, sh As Worksheet, ks As String
    k = KhoaSo(ncc)
    fileIdx = 0
    If k = "" Then Exit Function
    ' 1) bang NCC -> Sheet_mau o sheet cau hinh
    Set wsC = CfgSheet()
    For r = MAP_FIRST_ROW To wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
        If KhoaSo(CStr(wsC.Cells(r, 1).Value)) = k Then
            nm = Trim$(CStr(wsC.Cells(r, 2).Value))
            For i = 1 To wbs.Count
                If SheetTonTai(wbs(i), nm) Then fileIdx = i: TimSheetMau = nm: Exit Function
            Next i
        End If
    Next r
    ' 2) ten sheet trung ten NCC (bo dau)
    For i = 1 To wbs.Count
        For Each sh In wbs(i).Worksheets
            If sh.Visible = xlSheetVisible And KhoaSo(sh.Name) = k Then fileIdx = i: TimSheetMau = sh.Name: Exit Function
        Next sh
    Next i
    ' 3) ten nay chua ten kia (it nhat 4 ky tu)
    If Len(k) >= 4 Then
        For i = 1 To wbs.Count
            For Each sh In wbs(i).Worksheets
                ks = KhoaSo(sh.Name)
                If sh.Visible = xlSheetVisible And Len(ks) >= 4 Then
                    If InStr(ks, k) > 0 Or InStr(k, ks) > 0 Then fileIdx = i: TimSheetMau = sh.Name: Exit Function
                End If
            Next sh
        Next i
    End If
End Function

Private Function ThongTinThuHuong(sh As Worksheet) As String
    ThongTinThuHuong = Trim$(CStr(sh.Range("H27").Value)) & " | TK " & Trim$(CStr(sh.Range("R34").Value)) & _
                       " - " & Trim$(CStr(sh.Range("R35").Value))
End Function


' Khong tu tim thay file Cong no -> mo hop thoai cho nguoi dung chon, luu duong dan vao DNTT_CauHinh!B6
Private Function ChonFileCongNo() As String
    Dim f As Variant, wsC As Worksheet, r As Long
    Hoi "Khong tu tim thay file Cong no.xlsx canh file nay." & vbCrLf & _
           "Hay chon file Cong no.xlsx o hop thoai tiep theo (chi can chon 1 lan, macro se nho duong dan).", _
           vbInformation, "DNTT " & DNTT_VERSION
    f = Application.GetOpenFilename("Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , "Chon file Cong no.xlsx")
    If VarType(f) = vbBoolean Then Exit Function
    If Not FileExists(CStr(f)) Then Exit Function
    Set wsC = CfgSheet()
    For r = 4 To 17
        If Trim$(CStr(wsC.Cells(r, 1).Value)) = "File_cong_no" Then wsC.Cells(r, 2).Value = CStr(f): Exit For
    Next r
    ChonFileCongNo = CStr(f)
End Function

' --------------------------------------------------------------------------
' DOC FILE CONG NO: tong tien + so bien ban theo (NCC, thang, ky)
'   ketqua(KhoaSo(NCC) & "|" & "mm/yyyy" & "|" & ky) = Array(tong_tien, so_bien_ban)
'   ky = 1, 2 va 0 (= ca thang)
' --------------------------------------------------------------------------
Private Function DocCongNo(ByRef loi As String, ByRef loaiMa As Object, ByRef loaiNCC As Object) As Object
    Dim p As String, wb As Workbook, moMoi As Boolean, ws As Worksheet
    Dim res As Object, r As Long, lastR As Long
    Set res = CreateObject("Scripting.Dictionary")
    Set DocCongNo = res
    Set loaiMa = CreateObject("Scripting.Dictionary")
    Set loaiNCC = CreateObject("Scripting.Dictionary")

    p = TimFile("File_cong_no", ("C" & ChrW(244) & "ng n" & ChrW(7907) & ".xlsx"))
    If p = "" Then p = ChonFileCongNo()
    If p = "" Then loi = "Chua chon file Cong no.xlsx (chay lai va chon file, hoac ghi duong dan o sheet " & CFG_SHEET & ", o B6)": Exit Function
    Set wb = WbDangMo(p)
    If wb Is Nothing Then Set wb = MoChiDoc(p): moMoi = True
    On Error Resume Next
    Set ws = wb.Sheets((ChrW(272) & ChrW(417) & "n h" & ChrW(224) & "ng chi ti" & ChrW(7871) & "t"))
    On Error GoTo 0
    If ws Is Nothing Then
        loi = "File Cong no khong co sheet 'Don hang chi tiet'"
        If moMoi Then wb.Close SaveChanges:=False
        Exit Function
    End If

    ' cot: B Ngay giao, C NCC, L Thanh tien, M Thang (mm/yyyy), N Ky trong thang, O Nguon file (bien ban)
    Dim tien As Object, bbAnh As Object, ngayCoAnh As Object, ngayKhongAnh As Object
    Set tien = CreateObject("Scripting.Dictionary")
    Set bbAnh = CreateObject("Scripting.Dictionary")
    Set ngayCoAnh = CreateObject("Scripting.Dictionary")
    Set ngayKhongAnh = CreateObject("Scripting.Dictionary")
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "(IMG_)?\d{8}_\d{6}\.(jpg|jpeg|png|heic)"

    Dim ncc As String, thang As String, kyTxt As String, ky As Long, ngay As String, anh As String
    Dim v As Variant, gk As Variant, amt As Double, kk As Variant
    lastR = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    For r = 5 To lastR
        ncc = KhoaSo(CStr(ws.Cells(r, "C").Value))
        v = ws.Cells(r, "M").Value
        If VarType(v) = vbDate Then thang = Format(v, "mm/yyyy") Else thang = Trim$(CStr(v))
        kyTxt = CStr(ws.Cells(r, "N").Value)
        If ncc <> "" And Len(thang) = 7 And kyTxt <> "" Then
            If InStr(kyTxt, "01-15") > 0 Then ky = 1 Else ky = 2
            amt = 0
            If IsNumeric(ws.Cells(r, "L").Value) Then amt = CDbl(ws.Cells(r, "L").Value)
            v = ws.Cells(r, "B").Value
            If IsDate(v) Then ngay = Format(CDate(v), "yyyymmdd") Else ngay = Trim$(CStr(v))
            anh = ""
            Set m = Nothing
            If re.Test(CStr(ws.Cells(r, "O").Value)) Then
                Set m = re.Execute(CStr(ws.Cells(r, "O").Value))
                anh = LCase$(m(0).Value)
            End If
            For Each gk In Array(ncc & "|" & thang & "|" & ky, ncc & "|" & thang & "|0")
                tien(gk) = tien(gk) + amt
                If anh <> "" Then
                    bbAnh(gk & "#" & ngay & "#" & anh) = 1
                    ngayCoAnh(gk & "#" & ngay) = 1
                Else
                    ngayKhongAnh(gk & "#" & ngay) = 1
                End If
            Next gk
        End If
    Next r
    ' loai hang: Bang gia tham chieu (B Ma hang, D Nhom hang cap 2) va Danh muc NCC (C Ten NCC, D Mat hang)
    Dim wsG As Worksheet
    Set wsG = Nothing
    On Error Resume Next
    Set wsG = wb.Sheets(("B" & ChrW(7843) & "ng gi" & ChrW(225) & " tham chi" & ChrW(7871) & "u"))
    On Error GoTo 0
    If Not wsG Is Nothing Then
        For r = 5 To wsG.Cells(wsG.Rows.Count, "B").End(xlUp).Row
            If Trim$(CStr(wsG.Cells(r, "B").Value)) <> "" And Trim$(CStr(wsG.Cells(r, "D").Value)) <> "" Then
                loaiMa(UCase$(Trim$(CStr(wsG.Cells(r, "B").Value)))) = LoaiTuNhom(CStr(wsG.Cells(r, "D").Value))
            End If
        Next r
    End If
    Set wsG = Nothing
    On Error Resume Next
    Set wsG = wb.Sheets(("Danh m" & ChrW(7909) & "c NCC"))
    On Error GoTo 0
    If Not wsG Is Nothing Then
        For r = 5 To wsG.Cells(wsG.Rows.Count, "C").End(xlUp).Row
            If Trim$(CStr(wsG.Cells(r, "C").Value)) <> "" Then
                loaiNCC(KhoaSo(CStr(wsG.Cells(r, "C").Value))) = LoaiTuNhom(CStr(wsG.Cells(r, "D").Value))
            End If
        Next r
    End If
    If moMoi Then wb.Close SaveChanges:=False

    ' so bien ban = so cap (ngay, anh bien ban dau tien) khac nhau
    '             + so ngay khong ghi ten anh ma ngay do chua co bien ban nao
    Dim dem As Object, parts() As String
    Set dem = CreateObject("Scripting.Dictionary")
    For Each kk In bbAnh.Keys
        parts = Split(kk, "#")
        dem(parts(0)) = dem(parts(0)) + 1
    Next kk
    For Each kk In ngayKhongAnh.Keys
        If Not ngayCoAnh.Exists(kk) Then
            parts = Split(kk, "#")
            dem(parts(0)) = dem(parts(0)) + 1
        End If
    Next kk
    For Each kk In tien.Keys
        res(kk) = Array(Round(tien(kk), 0), dem(kk))
    Next kk
End Function


' "Do kho" / "Hang kho" -> do kho; con lai -> thuc pham tuoi song
Private Function LoaiTuNhom(ByVal nhom As String) As String
    If InStr(UCase$(BoDau(nhom)), "KHO") > 0 Then
        LoaiTuNhom = (ChrW(273) & ChrW(7891) & " kh" & ChrW(244))
    Else
        LoaiTuNhom = ("th" & ChrW(7921) & "c ph" & ChrW(7849) & "m t" & ChrW(432) & ChrW(417) & "i s" & ChrW(7889) & "ng")
    End If
End Function

' Loai hang cua NCC theo cac mat hang trong PO (CongNo cot D, tien = F x G; dong CongNo thuoc PO qua PR cot Q)
' Chiem phan lon gia tri thi lay loai do. Khong tra duoc ma hang thi theo Danh muc NCC.
Private Function LoaiHangNCC(ByVal kNcc As String, ByVal poList As String, loaiMa As Object, loaiNCC As Object, _
                             ByRef ghiChu As String) As String
    Dim wsPR As Worksheet, wsCN As Worksheet, r As Long, po As String, ma As String, tien As Double
    Dim tuoi As Double, kho As Double
    Set wsPR = ThisWorkbook.Sheets("PR")
    Set wsCN = ThisWorkbook.Sheets("CongNo")
    For r = 3 To wsCN.Cells(wsCN.Rows.Count, "D").End(xlUp).Row
        po = Trim$(CStr(wsPR.Cells(r, "Q").Value))
        If po <> "" Then
            If InStr(", " & poList & ", ", ", " & po & ", ") > 0 Then
                ma = UCase$(Trim$(CStr(wsCN.Cells(r, "D").Value)))
                tien = 0
                If IsNumeric(wsPR.Cells(r, "E").Value) And IsNumeric(wsPR.Cells(r, "F").Value) Then
                    tien = CDbl(wsPR.Cells(r, "E").Value) * CDbl(wsPR.Cells(r, "F").Value)
                End If
                If loaiMa.Exists(ma) Then
                    If loaiMa(ma) = (ChrW(273) & ChrW(7891) & " kh" & ChrW(244)) Then kho = kho + tien Else tuoi = tuoi + tien
                End If
            End If
        End If
    Next r
    If tuoi + kho > 0 Then
        If kho > tuoi Then LoaiHangNCC = (ChrW(273) & ChrW(7891) & " kh" & ChrW(244)) Else LoaiHangNCC = ("th" & ChrW(7921) & "c ph" & ChrW(7849) & "m t" & ChrW(432) & ChrW(417) & "i s" & ChrW(7889) & "ng")
        If tuoi > 0 And kho > 0 Then ghiChu = " | PO co ca hang tuoi song va do kho - kiem tra noi dung"
    ElseIf loaiNCC.Exists(kNcc) Then
        LoaiHangNCC = loaiNCC(kNcc)
    Else
        LoaiHangNCC = ("th" & ChrW(7921) & "c ph" & ChrW(7849) & "m t" & ChrW(432) & ChrW(417) & "i s" & ChrW(7889) & "ng")
        ghiChu = " | khong tra duoc loai hang, tam ghi tuoi song"
    End If
End Function


' --------------------------------------------------------------------------
' SHEET DNTT
' --------------------------------------------------------------------------
Private Function OutSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(OUT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("MIGO"))
        ws.Name = OUT_SHEET
    End If
    Set OutSheet = ws
End Function

Private Sub VeTieuDe(ws As Worksheet)
    ws.Range("A1:P1").Value = Array("STT", "Nha_cung_cap", "So_PO", "So_tien_PO", "Thang", "Ky", _
        "Noi_dung_thanh_toan", "Ngay_lap", "Han_thanh_toan", "So_bien_ban", "Kiem_tra_ky", _
        "Sheet_mau", "File_mau", "Nguoi_thu_huong (KIEM TRA dung NCC)", "File_DNTT_da_tao", "Status")
    ws.Range("A2:P2").Value = Array("Tu dong", "Tu dong = sheet PO", "Tu dong", "Tu dong = PO (tien tren SAP)", _
        "Tu nhan ra", "Tu nhan ra", "SUA DUOC", "SUA DUOC", "SUA DUOC", "SUA DUOC (tu dem tu Cong no)", _
        "Tu dong", "SUA DUOC (tu tim)", "Tu dong", "Tu dong", "<< TaoDNTT >>", "<< TaoDNTT >>")
    With ws.Range("A1:P1")
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Range("A2:P2").Font.Italic = True
    ws.Range("A2:P2").Font.Color = RGB(118, 118, 118)
End Sub

' Danh sach PO (theo thu tu) dang co o sheet PO, de biet sheet DNTT co con khop khong
Private Function DanhSachPO() As String
    Dim wsPO As Worksheet, r As Long, s As String
    Set wsPO = ThisWorkbook.Sheets("PO")
    For r = 3 To wsPO.Cells(wsPO.Rows.Count, "B").End(xlUp).Row
        If Trim$(CStr(wsPO.Cells(r, "J").Value)) <> "" Then s = NoiChuoi(s, Trim$(CStr(wsPO.Cells(r, "J").Value)))
    Next r
    DanhSachPO = s
End Function

Private Function DanhSachPOTrenDNTT(ws As Worksheet) As String
    Dim r As Long, s As String, x As Variant
    If ws.Range("A1").Value <> "STT" Then Exit Function
    For r = 3 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        For Each x In Split(CStr(ws.Cells(r, "C").Value), ",")
            s = NoiChuoi(s, Trim$(CStr(x)))
        Next x
    Next r
    DanhSachPOTrenDNTT = s
End Function

Private Function CungTapPO(ByVal a As String, ByVal b As String) As Boolean
    Dim x As Variant, n1 As Long, n2 As Long
    If a = "" Or b = "" Then Exit Function
    For Each x In Split(a, ", ")
        n1 = n1 + 1
        If InStr(", " & b & ", ", ", " & x & ", ") = 0 Then Exit Function
    Next x
    n2 = UBound(Split(b, ", ")) + 1
    CungTapPO = (n1 = n2)
End Function


' --------------------------------------------------------------------------
' TONG HOP: moi NCC co PO = 1 dong (gop nhieu PO cung NCC)
' --------------------------------------------------------------------------
Public Sub TongHopDNTT(Optional ByVal silent As Boolean = False)
    Dim wsPO As Worksheet, wsMG As Worksheet, ws As Worksheet
    Set wsPO = ThisWorkbook.Sheets("PO")
    Set wsMG = ThisWorkbook.Sheets("MIGO")
    CfgSheet
    Set ws = OutSheet()

    Dim r As Long, lastR As Long, k As String, ncc As String, soPO As String
    Dim grp As Object, order As Collection, a As Variant, amt As Double, migoDone As Object

    ' PO nao da MIGO
    Set migoDone = CreateObject("Scripting.Dictionary")
    For r = 3 To wsMG.Cells(wsMG.Rows.Count, "B").End(xlUp).Row
        If Trim$(CStr(wsMG.Cells(r, "G").Value)) <> "" Then migoDone(Trim$(CStr(wsMG.Cells(r, "B").Value))) = 1
    Next r

    ' 1) gom PO theo NCC: Array(ten NCC, ds PO, tong tien, ghi chu)
    Set grp = CreateObject("Scripting.Dictionary")
    Set order = New Collection
    lastR = wsPO.Cells(wsPO.Rows.Count, "B").End(xlUp).Row
    For r = 3 To lastR
        soPO = Trim$(CStr(wsPO.Cells(r, "J").Value))
        ncc = Trim$(CStr(wsPO.Cells(r, "B").Value))
        If soPO <> "" And ncc <> "" Then
            k = KhoaSo(ncc)
            If Not grp.Exists(k) Then
                grp.Add k, Array(ncc, "", 0#, "")
                order.Add k
            End If
            a = grp(k)
            a(1) = NoiChuoi(CStr(a(1)), soPO)
            ' so tien tren PO: uu tien tien doc lai tu SAP (cot G), khong co thi lay tong CongNo (cot F)
            If IsNumeric(wsPO.Cells(r, "G").Value) And Trim$(CStr(wsPO.Cells(r, "G").Value)) <> "" Then
                amt = CDbl(wsPO.Cells(r, "G").Value)
            ElseIf IsNumeric(wsPO.Cells(r, "F").Value) Then
                amt = CDbl(wsPO.Cells(r, "F").Value)
            Else
                amt = 0
            End If
            a(2) = CDbl(a(2)) + amt
            If Not migoDone.Exists(soPO) Then a(3) = NoiChuoi(CStr(a(3)), "PO " & soPO & " chua MIGO")
            grp(k) = a
        End If
    Next r

    ' 2) doc file Cong no, nhan ky cua tung NCC bang cach so tien
    Dim cn As Object, loiCN As String, kq As Variant, gk As Variant, parts() As String
    Dim loaiMa As Object, loaiNCC As Object, loai As String, ghiLoai As String
    Dim kyNCC As Object, phieu As Object, v As Variant, best As String, bestN As Long
    Set cn = DocCongNo(loiCN, loaiMa, loaiNCC)
    Set kyNCC = CreateObject("Scripting.Dictionary")
    Set phieu = CreateObject("Scripting.Dictionary")
    For Each v In order
        a = grp(v)
        For Each gk In cn.Keys
            parts = Split(gk, "|")
            If parts(0) = CStr(v) And parts(2) <> "0" Then
                kq = cn(gk)
                If Abs(CDbl(kq(0)) - Round(CDbl(a(2)), 0)) < 1 Then
                    kyNCC(v) = parts(1) & "|" & parts(2)
                    phieu(kyNCC(v)) = phieu(kyNCC(v)) + 1
                    Exit For
                End If
            End If
        Next gk
        If Not kyNCC.Exists(v) Then
            ' thu ca thang (NCC chi bao cong no cuoi thang)
            For Each gk In cn.Keys
                parts = Split(gk, "|")
                If parts(0) = CStr(v) And parts(2) = "0" Then
                    kq = cn(gk)
                    If Abs(CDbl(kq(0)) - Round(CDbl(a(2)), 0)) < 1 Then kyNCC(v) = parts(1) & "|0": Exit For
                End If
            Next gk
        End If
    Next v
    ' ky chung cua dot PO = ky duoc nhieu NCC khop nhat
    For Each gk In phieu.Keys
        If phieu(gk) > bestN Then bestN = phieu(gk): best = CStr(gk)
    Next gk
    If best = "" Then
        Dim tf As Long, nf As Long, kf As Long
        KyDuPhong tf, nf, kf
        best = Format(tf, "00") & "/" & nf & "|" & kf
    End If

    ' 3) mo file mau (chi doc) de tim sheet + doc nguoi thu huong
    Dim paths As Collection, wbs As New Collection, opened As New Collection, i As Long, wb As Workbook
    Set paths = TemplatePaths()
    Application.ScreenUpdating = False
    For i = 1 To paths.Count
        Set wb = WbDangMo(paths(i))
        If wb Is Nothing Then
            Set wb = MoChiDoc(paths(i))
            opened.Add wb
        End If
        wbs.Add wb
    Next i
    ThisWorkbook.Activate

    ' 4) ghi sheet DNTT
    Dim outR As Long, idx As Long, sheetNm As String, fileIdx As Long, kyStr As String, thangStr As String
    Dim kiemTra As String, giuSheet As Object, nBB As Variant
    ' giu lai ten sheet mau nguoi dung da sua tay
    Set giuSheet = CreateObject("Scripting.Dictionary")
    If ws.Range("A1").Value = "STT" Then
        For r = 3 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
            If Trim$(CStr(ws.Cells(r, "L").Value)) <> "" Then giuSheet(KhoaSo(CStr(ws.Cells(r, "B").Value))) = Trim$(CStr(ws.Cells(r, "L").Value))
        Next r
    End If

    ws.Cells.Clear
    VeTieuDe ws
    outR = 3
    For Each v In order
        a = grp(v)
        If kyNCC.Exists(v) Then
            parts = Split(kyNCC(v), "|")
            kiemTra = "KHOP: so tien PO = tong don hang " & IIf(parts(1) = "0", "ca thang ", "ky " & parts(1) & " thang ") & parts(0)
        Else
            parts = Split(best, "|")
            If loiCN <> "" Then
                kiemTra = "KHONG DOC DUOC CONG NO: " & loiCN
            ElseIf cn.Exists(CStr(v) & "|" & parts(0) & "|" & parts(1)) Then
                kq = cn(CStr(v) & "|" & parts(0) & "|" & parts(1))
                kiemTra = "LECH: tong don hang ky " & parts(1) & " thang " & parts(0) & " = " & Format(kq(0), "#,##0") & _
                          " (PO lech " & Format(Round(CDbl(a(2)), 0) - CDbl(kq(0)), "#,##0") & ") - tam lay ky chung cua dot PO"
            Else
                kiemTra = "KHONG THAY DON HANG cua NCC nay trong ky " & parts(1) & " thang " & parts(0) & " - kiem tra"
            End If
        End If
        thangStr = parts(0): kyStr = parts(1)

        ws.Cells(outR, "A").Value = outR - 2
        ws.Cells(outR, "B").Value = a(0)
        ws.Cells(outR, "C").NumberFormat = "@"
        ws.Cells(outR, "C").Value = a(1)
        ws.Cells(outR, "D").Value = Round(CDbl(a(2)), 0)
        ws.Cells(outR, "E").NumberFormat = "@"
        ws.Cells(outR, "E").Value = thangStr
        ws.Cells(outR, "F").Value = CLng(kyStr)
        ghiLoai = ""
        loai = LoaiHangNCC(CStr(v), CStr(a(1)), loaiMa, loaiNCC, ghiLoai)
        ws.Cells(outR, "G").Value = NoiDung(loai, ThangNoiDung(), CLng(kyStr))
        ws.Cells(outR, "H").Value = Date
        ws.Cells(outR, "I").Value = Date + SoNgayHan()
        nBB = Empty
        If cn.Exists(CStr(v) & "|" & thangStr & "|" & kyStr) Then
            kq = cn(CStr(v) & "|" & thangStr & "|" & kyStr)
            nBB = kq(1)
        End If
        If Not IsEmpty(nBB) Then ws.Cells(outR, "J").Value = nBB
        ws.Cells(outR, "K").Value = kiemTra & ghiLoai
        If Left$(kiemTra, 4) <> "KHOP" Then ws.Cells(outR, "K").Interior.Color = RGB(255, 235, 156)

        ' sheet mau
        sheetNm = ""
        fileIdx = 0
        If giuSheet.Exists(CStr(v)) Then
            sheetNm = giuSheet(CStr(v))
            For idx = 1 To wbs.Count
                If SheetTonTai(wbs(idx), sheetNm) Then fileIdx = idx: Exit For
            Next idx
        End If
        If fileIdx = 0 Then sheetNm = TimSheetMau(CStr(a(0)), wbs, fileIdx)

        If fileIdx > 0 Then
            ws.Cells(outR, "L").Value = sheetNm
            ws.Cells(outR, "M").Value = wbs(fileIdx).Name
            ws.Cells(outR, "N").Value = ThongTinThuHuong(wbs(fileIdx).Sheets(sheetNm))
            ws.Cells(outR, "P").Value = "SAN SANG" & IIf(CStr(a(3)) <> "", " (luu y: " & a(3) & ")", "")
        Else
            ws.Cells(outR, "P").Value = "CHUA TIM THAY SHEET MAU - go ten sheet o cot L, hoac them dong vao bang o sheet " & CFG_SHEET
            ws.Cells(outR, "P").Interior.Color = RGB(255, 199, 206)
        End If
        outR = outR + 1
    Next v

    For i = 1 To opened.Count
        opened(i).Close SaveChanges:=False
    Next i
    ThisWorkbook.Activate
    Application.ScreenUpdating = True

    If outR > 3 Then
        ws.Range("G3:J" & outR - 1).Interior.Color = RGB(255, 242, 204)
        ws.Range("L3:L" & outR - 1).Interior.Color = RGB(255, 242, 204)
        ws.Range("D3:D" & outR - 1).NumberFormat = "#,##0"
        ws.Range("H3:I" & outR - 1).NumberFormat = "dd/mm/yyyy"
    End If
    ws.Columns("A:P").AutoFit
    If ws.Columns("G").ColumnWidth > 55 Then ws.Columns("G").ColumnWidth = 55
    If ws.Columns("K").ColumnWidth > 60 Then ws.Columns("K").ColumnWidth = 60
    If ws.Columns("N").ColumnWidth > 60 Then ws.Columns("N").ColumnWidth = 60
    If ws.Columns("O").ColumnWidth > 50 Then ws.Columns("O").ColumnWidth = 50
    If ws.Columns("P").ColumnWidth > 60 Then ws.Columns("P").ColumnWidth = 60

    If Not silent Then
        ws.Activate
        If paths.Count = 0 Then
            Hoi "Khong tim thay file mau DNTT.xlsx." & vbCrLf & _
                   "Dat file DNTT.xlsx cung thu muc voi file nay, hoac ghi duong dan o sheet " & CFG_SHEET & " (o B4).", _
                   vbExclamation, "TongHopDNTT " & DNTT_VERSION
        Else
            Hoi "Da liet ke " & (outR - 3) & " nha cung cap co PO." & vbCrLf & _
                   IIf(loiCN <> "", "CANH BAO: " & loiCN & vbCrLf, "") & vbCrLf & _
                   "Kiem tra sheet DNTT: cot K (ky co khop so tien khong), cot J (so bien ban)," & vbCrLf & _
                   "cot N (nguoi thu huong / so TK dung NCC). Sua o vang neu can, roi chay TaoDNTT.", _
                   vbInformation, "TongHopDNTT " & DNTT_VERSION
        End If
    End If
End Sub


' --------------------------------------------------------------------------
' TAO DNTT: dien vao file mau, chen chu ky, luu file moi + PDF
' --------------------------------------------------------------------------
Public Sub TaoDNTT()
    Dim ws As Worksheet, r As Long, lastR As Long, msg As String, n As Long, nCanhBao As Long
    Set ws = OutSheet()
    ' sheet DNTT chua co / khong con khop cac PO hien tai -> tong hop lai
    ' v6-Tram: chay qua tram thi sheet DNTT do giao dien gui sang -> KHONG tong hop lai tu sheet PO
    If Not TRAM_AUTO Then
        If Not CungTapPO(DanhSachPO(), DanhSachPOTrenDNTT(ws)) Then TongHopDNTT True
    End If
    ws.Activate

    lastR = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For r = 3 To lastR
        If Trim$(CStr(ws.Cells(r, "M").Value)) <> "" And Val(ws.Cells(r, "D").Value) > 0 Then
            n = n + 1
            msg = msg & n & ". " & ws.Cells(r, "B").Value & ": " & Format(ws.Cells(r, "D").Value, "#,##0") & _
                  " d | ky " & ws.Cells(r, "F").Value & " T" & ws.Cells(r, "E").Value & " | " & ws.Cells(r, "J").Value & " bien ban"
            If Left$(CStr(ws.Cells(r, "K").Value), 4) <> "KHOP" Then msg = msg & "  (!) KIEM TRA KY": nCanhBao = nCanhBao + 1
            msg = msg & vbCrLf
        End If
    Next r
    If n = 0 Then
        Hoi "Khong co NCC nao san sang lap DNTT (chua co PO, hoac chua tim thay sheet mau / file mau)." & vbCrLf & _
               "Xem cot P sheet DNTT.", vbExclamation, "TaoDNTT " & DNTT_VERSION
        Exit Sub
    End If
    msg = "Se lap " & n & " de nghi thanh toan (ngay lap " & Format(ws.Cells(3, "H").Value, "dd/mm/yyyy") & _
          ", han " & Format(ws.Cells(3, "I").Value, "dd/mm/yyyy") & "):" & vbCrLf & vbCrLf & msg & vbCrLf
    If nCanhBao > 0 Then msg = msg & "Co " & nCanhBao & " NCC so tien PO khong khop dung 1 ky - xem cot K sheet DNTT." & vbCrLf
    msg = msg & "File mau goc KHONG bi sua; ket qua luu thanh file moi." & vbCrLf & vbCrLf & "Tiep tuc?"
    If Hoi(msg, vbYesNo + vbQuestion + vbDefaultButton2, "TaoDNTT " & DNTT_VERSION) <> vbYes Then Exit Sub

    Dim paths As Collection, i As Long, wb As Workbook, fileNm As String, folder As String
    Dim outPath As String, pdfPath As String, sigPath As String, tenKy As String
    Dim filled As Collection, sh As Worksheet, nm As Variant, vArr As Variant, arr() As String, j As Long, ketQua As String, sOk As String
    Set paths = TemplatePaths()
    folder = OutFolder()
    sigPath = DuongDanChuKy()
    tenKy = "T" & Left$(CStr(ws.Cells(3, "E").Value), 2) & IIf(Val(ws.Cells(3, "F").Value) > 0, "_Ky" & ws.Cells(3, "F").Value, "")

    For i = 1 To paths.Count
        fileNm = Fso().GetFileName(paths(i))
        Set filled = New Collection
        For r = 3 To lastR
            If LCase$(Trim$(CStr(ws.Cells(r, "M").Value))) = LCase$(fileNm) And Val(ws.Cells(r, "D").Value) > 0 Then filled.Add r
        Next r
        If filled.Count > 0 Then
            If Not WbDangMo(paths(i)) Is Nothing Then
                Hoi "File " & fileNm & " dang mo. Hay DONG file do (khong can luu) roi chay lai TaoDNTT.", _
                       vbExclamation, "TaoDNTT " & DNTT_VERSION
                Exit Sub
            End If
            Application.ScreenUpdating = False
            Set wb = MoChiDoc(paths(i))
            ReDim arr(1 To filled.Count)
            j = 0
            For Each nm In filled
                r = CLng(nm)
                Set sh = wb.Sheets(CStr(ws.Cells(r, "L").Value))
                sh.Visible = xlSheetVisible
                DienSheet sh, ws, r
                ChenChuKy sh, sigPath
                j = j + 1
                arr(j) = sh.Name
            Next nm
            ' an cac sheet NCC khac (sheet he thong DS/DS2 van an nhu cu)
            For Each sh In wb.Worksheets
                If sh.Visible = xlSheetVisible And Not TrongMang(arr, sh.Name) Then sh.Visible = xlSheetHidden
            Next sh
            Application.CalculateFull
            For j = 1 To UBound(arr)
                ' (so tien bang chu da xu ly trong LamDep)
            Next j
            Application.CalculateFull
            For Each nm In filled
                r = CLng(nm)
                ws.Cells(r, "N").Value = ThongTinThuHuong(wb.Sheets(CStr(ws.Cells(r, "L").Value)))
            Next nm

            outPath = folder & "\DNTT_" & tenKy & "_" & Format(Now, "yyyymmdd_hhnn") & _
                      IIf(i > 1, "_" & Fso().GetBaseName(fileNm), "") & ".xlsx"
            Application.DisplayAlerts = False
            wb.SaveAs Filename:=outPath, FileFormat:=51
            Application.DisplayAlerts = True
            sOk = Fso().GetFileName(outPath)
            If UCase$(CfgVal("Xuat_PDF")) <> "KHONG" Then
                pdfPath = Left$(outPath, Len(outPath) - 5) & ".pdf"
                On Error Resume Next
                vArr = arr
                wb.Activate
                wb.Worksheets(vArr).Select
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
                    IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
                If Err.Number <> 0 Then
                    sOk = sOk & " (LOI XUAT PDF: " & Err.Description & ")"
                    Err.Clear
                Else
                    sOk = sOk & " + PDF"
                End If
                On Error GoTo 0
            End If
            wb.Close SaveChanges:=False
            ThisWorkbook.Activate
            Application.ScreenUpdating = True

            For Each nm In filled
                r = CLng(nm)
                ws.Cells(r, "O").Value = outPath
                ws.Cells(r, "P").Value = "DA TAO " & Format(Now, "dd/mm/yyyy hh:nn")
                ws.Cells(r, "P").Interior.Color = RGB(198, 239, 206)
            Next nm
            ketQua = ketQua & "- " & sOk & vbCrLf
        End If
    Next i

    ws.Activate
    ThisWorkbook.Save
    If Hoi("Da tao xong (thu muc " & folder & "):" & vbCrLf & ketQua & vbCrLf & _
                            "Nho KIEM TRA so tien bang chu, nguoi thu huong, so TK truoc khi trinh ky." & vbCrLf & vbCrLf & _
              "Mo thu muc chua file?", vbYesNo + vbInformation, "TaoDNTT " & DNTT_VERSION) = vbYes And Not TRAM_AUTO Then
        Shell "explorer.exe """ & folder & """", vbNormalFocus
    End If
End Sub

Private Function TrongMang(arr() As String, ByVal s As String) As Boolean
    Dim j As Long
    For j = LBound(arr) To UBound(arr)
        If arr(j) = s Then TrongMang = True: Exit Function
    Next j
End Function

' Dien 1 sheet mau DNTT
Private Sub DienSheet(sh As Worksheet, ws As Worksheet, ByVal r As Long)
    Dim nd As Date, han As Date, s As String, cur As String, poTxt As String

    If IsDate(ws.Cells(r, "H").Value) Then nd = CDate(ws.Cells(r, "H").Value) Else nd = Date
    If IsDate(ws.Cells(r, "I").Value) Then han = CDate(ws.Cells(r, "I").Value) Else han = nd + SoNgayHan()

    ' ngay lap: "<dia danh>, ngay dd/mm/yyyy"
    s = CfgVal("Dia_danh")
    If s = "" Then s = ("H" & ChrW(432) & "ng Y" & ChrW(234) & "n")
    sh.Range("Q9").Value = s
    sh.Range("W9").Value = Day(nd)
    sh.Range("Y9").Value = Month(nd)
    sh.Range("AA9").Value = Year(nd)

    s = CfgVal("Nguoi_de_nghi")
    If s <> "" Then sh.Range("H11").Value = s
    s = CfgVal("Bo_phan")
    If s = "" Then s = ("B" & ChrW(7870) & "P")
    sh.Range("T11").Value = s

    sh.Range("H14").Value = ws.Cells(r, "G").Value
    sh.Range("H16").Value = Round(CDbl(ws.Cells(r, "D").Value), 0)

    ' han thanh toan
    sh.Range("W16").Value = Day(han)
    sh.Range("Y16").Value = Month(han)
    sh.Range("AA16").Value = Year(han)

    ' so bien ban giao nhan, nghiem thu (H24)
    If IsNumeric(ws.Cells(r, "J").Value) And Trim$(CStr(ws.Cells(r, "J").Value)) <> "" Then
        sh.Range("H24").Value = CLng(ws.Cells(r, "J").Value)
    Else
        sh.Range("H24").ClearContents
    End If

    ' Vu viec / Ghi chu = so PO (giu lai noi dung chuyen khoan rieng cua NCC neu mau dang co, vd NCC MAU 30)
    poTxt = "PO: " & Trim$(CStr(ws.Cells(r, "C").Value))
    cur = Trim$(CStr(sh.Range("H37").Value))
    If cur = "" Or UCase$(Left$(cur, 2)) = "PO" Then
        sh.Range("H37").Value = poTxt
    Else
        sh.Range("H37").Value = poTxt & " | " & cur
    End If

    LamDep sh
End Sub

' Trinh bay giong ban da duyet: bo nen xam, bo khung vien, chu den, du dau "/",
' ten ngan hang / nguoi thu huong dai thi xuong dong, so tien bang chu dung o.
Private Sub LamDep(sh As Worksheet)
    Dim r As Long, f As String, oChu As String, a As Variant, t As String

    With sh.Range("A1:AF60")
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
    End With
    sh.Range("A1:AB50").Font.Color = vbBlack

    If Len(Trim$(CStr(sh.Range("H17").Value))) <= 1 Then sh.Range("H17").ClearContents
    For Each a In Array("X9", "Z9", "X16", "Z16")
        sh.Range(a).Value = "/"
    Next a

    ' ten ngan hang (R35) / nguoi thu huong (H27) qua dai
    For Each a In Array(Array("R35", 48), Array("H27", 70))
        t = CStr(sh.Range(a(0)).Value)
        If Len(t) > a(1) Then
            With sh.Range(a(0)).MergeArea
                .WrapText = True
                .Font.Size = 9
                .VerticalAlignment = xlTop
            End With
            If sh.Range(a(0)).RowHeight < 24 Then sh.Range(a(0)).RowHeight = 24
        End If
    Next a

    ' so tien bang chu: mau thuc pham o J77, mau CCDC o J87 -> tim o co cong thuc "... M16"
    oChu = "J77"
    For r = 60 To 100
        f = ""
        On Error Resume Next
        f = sh.Range("J" & r).Formula
        On Error GoTo 0
        If InStr(f, "M16") > 0 And InStr(f, "C" & r) > 0 Then oChu = "J" & r: Exit For
    Next r
    sh.Range("H18").Formula = "=IF(H16<>0," & oChu & ",0)"
End Sub

Private Function DuongDanChuKy() As String
    Dim p As String
    ' v5: mac dinh KHONG chen chu ky (nguoi de nghi ky tay). Muon chen thi ghi duong dan anh o DNTT_CauHinh!B7
    p = CfgVal("File_chu_ky")
    If p = "" Or UCase$(p) = "KHONG" Then Exit Function
    If FileExists(p) Then DuongDanChuKy = p
End Function

' Chen anh chu ky vao o ky cua "Nguoi de nghi thanh toan" (cot B-G, dong 41-45, tren ten o dong 46)
Private Sub ChenChuKy(sh As Worksheet, ByVal p As String)
    Dim shp As Shape, vung As Range, h As Double, w As Double
    If p = "" Then Exit Sub
    On Error Resume Next
    sh.Shapes(SIG_SHAPE).Delete
    On Error GoTo 0
    Set vung = sh.Range("B41:G45")
    Set shp = sh.Shapes.AddPicture(p, msoFalse, msoTrue, vung.Left, vung.Top, -1, -1)
    shp.Name = SIG_SHAPE
    shp.LockAspectRatio = msoTrue
    h = vung.Height * 0.85
    w = vung.Width * 0.8      ' nam gon trong o Nguoi de nghi, khong lan sang o Bo phan xac nhan
    shp.Height = h
    If shp.Width > w Then shp.Width = w
    shp.Left = vung.Left + (vung.Width - shp.Width) / 2
    shp.Top = vung.Top + (vung.Height - shp.Height) / 2
    shp.Placement = xlMove
    shp.PrintObject = True
End Sub

' Sua ten DocSo (so tien bang chu) neu sheet mau bi loi #REF (vd sheet T&P)
Private Sub SuaDocSo(sh As Worksheet)
    Dim bad As Boolean
    bad = IsError(sh.Range("H18").Value)
    If Not bad Then bad = (Trim$(CStr(sh.Range("H18").Value)) = "0")
    If bad Then
        On Error Resume Next
        sh.Names.Add Name:="DocSo", RefersTo:="='" & Replace(sh.Name, "'", "''") & "'!$J$77"
        If Err.Number <> 0 Then
            Err.Clear
            sh.Range("H18").Formula = "=J77"
        End If
        On Error GoTo 0
        Application.CalculateFull
        If IsError(sh.Range("H18").Value) Then sh.Range("H18").Formula = "=J77"
    End If
End Sub

Private Sub InitBoDau()
    Dim codes As Variant, bases As String, i As Long
    Set gBoDau = CreateObject("Scripting.Dictionary")
    bases = "AAAAEEEIIOOOOUUYaaaaeeeiioooouuyAaDdIiUuOoUuAaAaAaAaAaAaAaAaAaAaAaAaEeEeEeEeEeEeEeEeIiIiOoOoOoOoOoOoOoOoOoOoOoOoUuUuUuUuUuUuUuYyYyYyYy"
    codes = Array( _
        192,193,194,195,200,201,202,204,205,210,211,212,213,217,218,221,224,225,226,227, _
        232,233,234,236,237,242,243,244,245,249,250,253,258,259,272,273,296,297,360,361, _
        416,417,431,432,7840,7841,7842,7843,7844,7845,7846,7847,7848,7849,7850,7851,7852,7853,7854,7855, _
        7856,7857,7858,7859,7860,7861,7862,7863,7864,7865,7866,7867,7868,7869,7870,7871,7872,7873,7874,7875, _
        7876,7877,7878,7879,7880,7881,7882,7883,7884,7885,7886,7887,7888,7889,7890,7891,7892,7893,7894,7895, _
        7896,7897,7898,7899,7900,7901,7902,7903,7904,7905,7906,7907,7908,7909,7910,7911,7912,7913,7914,7915, _
        7916,7917,7918,7919,7920,7921,7922,7923,7924,7925,7926,7927,7928,7929)
    For i = 0 To UBound(codes)
        gBoDau(CLng(codes(i))) = Mid$(bases, i + 1, 1)
    Next i
End Sub

Sub PhienBanDNTT()
    Hoi "Module DNTT_Automation " & DNTT_VERSION, vbInformation
End Sub
