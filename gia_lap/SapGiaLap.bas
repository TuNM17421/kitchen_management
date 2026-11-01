Attribute VB_Name = "SapGiaLap"
Option Explicit

' ================================================================================
'  SAP GIA LAP - chay thu cac macro SAP (CreatePR / CreatePO / PostMIGO /
'  CapNhatPOTuSAP / TuDong; ban v2: them SuaPO_ME22N / PostMIGO_TungPhan /
'  tram SAP v23.1-Tram) NGAY TRONG EXCEL, KHONG CAN SAP THAT.
'
'  v2 (31/10/2026, du an Quy trinh PR truoc):
'    - ME22N: mo PO (Shift+F5, MEPO_SELECT-EBELN), Ctrl+F1 hien thi <-> sua,
'      sua SL dong co san, them dong tu PR bo sung, Save cap nhat DB
'    - MIGO: o header GOHEAD-BLDAT / BUDAT / LFSNR, nhap MOT PHAN so luong,
'      bo tich dong, nut Check (tbar[1]/btn[7]), chan nhan vuot SL con mo
'    - iconify / maximize cua so, nhieu cua so SAP (ses[n]) cho tram chay song song
'    - DB dung chung (o B15): tram mo file Excel moi lan chay nen chung tu
'      phai luu ra file ngoai thi lenh sau (PO, ME22N, MIGO) moi thay
'
'  Gia lap dung cac man hinh + ID control ma macro da chay on tren SAP that:
'    ME51N (tao PR, grid ALV), ME21N (tao PO, bang item + Conditions),
'    ME23N (xem PO), MIGO (nhap kho), popup wnd[1], thanh trang thai (Text +
'    MessageType S/W/I/E/A), cuon bang, thu gon/mo Item Detail.
'  Moi thao tac macro lam tren "SAP" duoc ghi vao sheet SAP_GiaLap_Log.
'  Chung tu "tao" ra (PR/PO/MIGO) luu o sheet SAP_GiaLap_DB - sua tay sheet nay
'  de gia lap viec sua PO tren SAP (ME22N) roi chay CapNhatPOTuSAP.
'  Loi / canh bao / popup cua SAP: bat o bang KICH BAN sheet SAP_GiaLap.
'
'  CAI DAT (1 lan):
'   1) Alt+F11 > File > Import: SapGiaLap.bas, SapGL_Ctl.cls, SapGL_List.cls, SapGL_Dong.cls
'      + ban SAP_Automation.bas / PO_CapNhat.bas co ho tro gia lap.
'   2) Alt+F8 > SapGiaLap_CaiDat > Run (tao 3 sheet, BAT gia lap).
'   3) Chay macro nhu binh thuong. Muon dung SAP that: SapGiaLap_Tat
'      (hoac go TAT o o B3 sheet SAP_GiaLap).
' ================================================================================

Public Const GL_VERSION As String = "SapGiaLap v2 (31/10/2026)"

Private Const SH_CFG As String = "SAP_GiaLap"
Private Const SH_LOG As String = "SAP_GiaLap_Log"
Private Const SH_DB As String = "SAP_GiaLap_DB"
Private Const KB_ROW1 As Long = 19          ' dong dau tien cua bang kich ban
Private Const KB_ROWN As Long = 120
Private Const NGUON As String = "SAP GUI (gia lap)"

' ---- doi tuong goc ----
Private gApp As SapGL_Ctl, gConn As SapGL_Ctl, gSes As SapGL_Ctl, gInfo As SapGL_Ctl
Private gReg As Collection                   ' LCase(id) -> control dang co tren man hinh
Private gTx As String                        ' giao dich dang mo
Private gTitle As String
Private gSbT As String, gSbM As String       ' thanh trang thai
Private gOkcd As String
Private gWAck As Boolean                     ' da Enter xac nhan canh bao (W)
Private gFocusTbl As String

' ---- popup wnd[1] ----
Private gPop As String                       ' "", TB (thong bao), HOI (Yes/No), LOI (Save/Edit/Cancel), CHON_PO (ME23N)
Private gPopTitle As String, gPopText As String, gPopViec As String, gPopSel As String

' ---- ME51N ----
Private prBsart As String, prItems As Collection, prCur As Long, prTab As String, prDetail As Boolean

' ---- ME21N / ME23N ----
Private poBsart As String, poVendor As String, poVendorOK As String
Private poItems As Collection, poCur As Long, poTab As String
Private poDetail As Boolean, poDetailInit As Boolean
Private poPos As Long, poCondPos As Long, poSoPO As String, poXem As Boolean
Private poKhoa As String                     ' ME22N: PO dang bi khoa (noi dung loi), "" = khong
Private poDoi As Boolean                     ' ME22N: da sua gi tu luc mo PO

' ---- MIGO ----
Private mgAction As String, mgRefdoc As String, mgPO As String, mgPOLoaded As String
Private mgItems As Collection, mgPos As Long
Private mgBldat As String, mgBudat As String, mgLfsnr As String

' ---- cau hinh ----
Private cDec As String, cGiaLe As Long, cVisPOThu As Long, cVisPOMo As Long, cVisMigo As Long, cVisCond As Long
Private cLogDoc As Boolean, cSys As String, cClient As String, cUser As String
Private cDbFile As String                    ' B15: file DB dung chung ("" = sheet SAP_GiaLap_DB)
Private gPhien As Long                       ' cua so SAP (session) dang dung: ses[gPhien]
Private gDem As Collection                   ' so lan da dung moi dong kich ban
Private gLogRow As Long


' ================================================================================
'  1. DIEM VAO
' ================================================================================

' Goi tu GetSapSession (SAP_Automation) qua Application.Run.
' Tra ve session gia lap neu sheet SAP_GiaLap co B3 = BAT, nguoc lai Nothing (= dung SAP that).
' phien = so thu tu cua so SAP (0 = cua so dau tien). Tram chay song song dung 0, 1...
Public Function SapGiaLap_PhienNeuBat(Optional ByVal phien As Variant) As Object
    Dim k As Long
    If Not SapGiaLap_DangBat() Then Exit Function
    If Not IsMissing(phien) Then
        If IsNumeric(phien) Then k = CLng(phien)
    End If
    If k < 0 Then k = 0
    If gSes Is Nothing Or k <> gPhien Then
        gPhien = k
        KhoiTao
    End If
    DocCauHinh
    DbTaiVe
    GL_LogDong "=== GetSapSession", SesId(), "Macro ket noi SAP GIA LAP (cua so " & (gPhien + 1) & ", giao dich dang mo: " & gTx & ")" & _
               IIf(cDbFile <> "", " | DB: " & DbDuongDan(), "")
    Set SapGiaLap_PhienNeuBat = gSes
End Function

' Goi tu TramDamBaoPhien: gia lap dang BAT -> coi nhu SAP da mo du n cua so, tra ve n. TAT -> 0.
Public Function SapGiaLap_DamBaoPhien(ByVal n As Long, Optional ByVal anSap As Variant) As Long
    If Not SapGiaLap_DangBat() Then Exit Function
    If n < 1 Then n = 1
    GL_LogDong "=== TramDamBaoPhien", "", "Gia lap: san sang " & n & " cua so SAP" & IIf(IsMissing(anSap), "", IIf(CBool(anSap), " (thu nho)", ""))
    SapGiaLap_DamBaoPhien = n
End Function

Private Function SesId() As String
    SesId = "/app/con[0]/ses[" & gPhien & "]"
End Function

Public Function SapGiaLap_DangBat() As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_CFG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    SapGiaLap_DangBat = (UCase(Trim(CStr(ws.Range("B3").Value))) = "BAT")
End Function

Public Sub SapGiaLap_Bat()
    SapGiaLap_CaiDat
End Sub

Public Sub SapGiaLap_Tat()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_CFG)
    On Error GoTo 0
    If Not ws Is Nothing Then ws.Range("B3").Value = "TAT"
    MsgBox "Da TAT SAP gia lap - cac macro se ket noi SAP THAT.", vbInformation, GL_VERSION
End Sub

' Dat lai trang thai man hinh SAP gia lap + bo dem kich ban (khong xoa du lieu PR/PO/MIGO)
Public Sub SapGiaLap_DatLai()
    Set gSes = Nothing
    Set gDem = Nothing
    gLogRow = 0
End Sub

Public Sub SapGiaLap_XoaLog()
    Dim ws As Worksheet
    Set ws = SheetLog()
    ws.Range("A3:H" & ws.Rows.Count).ClearContents
    gLogRow = 3
End Sub

' Xoa toan bo chung tu da "tao" tren SAP gia lap (sheet SAP_GiaLap_DB)
Public Sub SapGiaLap_XoaDuLieu()
    Dim wb As Workbook
    DocCauHinh
    If MsgBox("Xoa het PR / PO / MIGO da tao tren SAP GIA LAP (sheet " & SH_DB & IIf(cDbFile <> "", " va file " & DbDuongDan(), "") & ")?", vbYesNo + vbQuestion, GL_VERSION) <> vbYes Then Exit Sub
    SheetDB().Range("A2:P" & SheetDB().Rows.Count).ClearContents
    If cDbFile <> "" Then
        Set wb = DbMoNgoai()
        If Not wb Is Nothing Then
            wb.Worksheets(1).Range("A2:P" & wb.Worksheets(1).Rows.Count).ClearContents
            DbDongNgoai wb
        End If
    End If
End Sub

' Tao 3 sheet (cau hinh + kich ban, nhat ky, du lieu) va BAT gia lap
Public Sub SapGiaLap_CaiDat()
    Dim ws As Worksheet
    Set ws = SheetCfg()
    ws.Range("B3").Value = "BAT"
    SheetLog
    SheetDB
    SapGiaLap_DatLai
    ws.Activate
    MsgBox "SAP GIA LAP dang BAT (" & GL_VERSION & ")." & vbCrLf & vbCrLf & _
           "- Chay CreatePR / CreatePO / SuaPO_ME22N / PostMIGO_TungPhan... (hoac tram SAP) nhu binh thuong: macro se lam viec voi SAP gia lap." & vbCrLf & _
           "- Moi thao tac ghi o sheet " & SH_LOG & ", chung tu tao ra o sheet " & SH_DB & "." & vbCrLf & _
           "- Bat loi / canh bao / popup SAP: danh x o cot A bang KICH BAN (sheet " & SH_CFG & ")." & vbCrLf & _
           "- Ve lai SAP that: chay SapGiaLap_Tat (hoac go TAT o o B3).", vbInformation, GL_VERSION
End Sub


' ================================================================================
'  2. SHEET CAU HINH / NHAT KY / DU LIEU
' ================================================================================

Private Function LaySheet(ByVal ten As String, ByRef moi As Boolean) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ten)
    On Error GoTo 0
    moi = ws Is Nothing
    If moi Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ten
    End If
    Set LaySheet = ws
End Function

Private Function SheetCfg() As Worksheet
    Dim ws As Worksheet, moi As Boolean, r As Long
    Set ws = LaySheet(SH_CFG, moi)
    If moi Or Trim(CStr(ws.Range("A1").Value)) = "" Then
        ws.Range("A1").Value = "SAP GIA LAP - " & GL_VERSION & " (o mau vang sua duoc)"
        ws.Range("A1").Font.Bold = True
        ws.Range("A2").Value = "B3 = BAT: GetSapSession tra ve SAP gia lap. B3 = TAT (hoac xoa sheet nay): dung SAP that."
        GhiCot ws, 3, 1, Array( _
            "Trang thai gia lap (BAT / TAT)", _
            "Dau thap phan cua SAP (SU3 > Defaults): , hoac .", _
            "So chu so le cua don gia (VND = 0)", _
            "Bang item PO: so dong hien khi Item Detail THU GON", _
            "Bang item PO: so dong hien khi Item Detail MO", _
            "Bang dong MIGO: so dong hien", _
            "Bang Conditions: so dong hien", _
            "So PR dau tien", _
            "So PO dau tien", _
            "So chung tu MIGO dau tien", _
            "Ghi log ca thao tac DOC (CO / KHONG)", _
            "He thong | Client | User", _
            "File DB dung chung (trong = sheet SAP_GiaLap_DB; ten file = cung thu muc)")
        ws.Range("B3").Value = "BAT"
        ws.Range("B4").NumberFormat = "@"
        ws.Range("B4").Value = ","
        ws.Range("B5").Value = 0
        ws.Range("B6").Value = 8
        ws.Range("B7").Value = 3
        ws.Range("B8").Value = 5
        ws.Range("B9").Value = 4
        ws.Range("B10:B12").NumberFormat = "@"
        ws.Range("B10").Value = "1000000001"
        ws.Range("B11").Value = "4500000001"
        ws.Range("B12").Value = "5000000001"
        ws.Range("B13").Value = "KHONG"
        ws.Range("B14").Value = "GIA"
        ws.Range("C14").NumberFormat = "@"
        ws.Range("C14").Value = "100"
        ws.Range("D14").Value = "GIALAP"
        ws.Range("B15").NumberFormat = "@"
        ws.Range("B15").Value = ""
        ws.Range("B3:B15").Interior.Color = RGB(255, 242, 204)
        ws.Range("C14:D14").Interior.Color = RGB(255, 242, 204)

        ws.Range("A17").Value = "KICH BAN LOI / CANH BAO / POPUP CUA SAP - danh x o cot A de BAT dong do (mac dinh: tat ca TAT = SAP chay suon se)"
        ws.Range("A17").Font.Bold = True
        ws.Range("A18:H18").Value = Array("Bat (x)", "Su_kien", "Giao_dich (trong = moi)", "Dieu_kien (chuoi chua, trong = moi)", _
                                          "Loai", "Noi_dung SAP bao", "So_lan (trong = moi lan)", "Gia_tri")
        ws.Range("A18:H18").Font.Bold = True
        ws.Range("A18:H18").Interior.Color = RGB(221, 235, 247)
        r = KB_ROW1
        KBMau ws, r, "LUU", "ME21N", "", "W", "Delivery date is in the past", 1, ""
        KBMau ws, r, "LUU", "ME21N", "", "POPUP_HOI", "Purchase order contains warnings. Do you want to save the document?", 1, ""
        KBMau ws, r, "LUU", "ME21N", "", "POPUP_LOI", "Purchase order still contains faulty items", 1, ""
        KBMau ws, r, "LUU", "ME51N", "", "W", "Release date of the item is in the past", 1, ""
        KBMau ws, r, "LUU", "ME51N", "", "E", "Budget for fund exceeded", "", ""
        KBMau ws, r, "MO_GD", "ME21N", "", "POPUP_TB", "Note: no vendor evaluation maintained for purchasing org.", 1, ""
        KBMau ws, r, "MO_GD", "ME51N", "", "POPUP_TB", "Your user has no default plant", 1, ""
        KBMau ws, r, "NHAP_DONG", "ME51N", "", "E", "Material does not exist in plant", "", ""
        KBMau ws, r, "NCC", "ME21N", "", "E", "Vendor is blocked for posting", "", ""
        KBMau ws, r, "LAY_DONG_PR", "ME21N", "", "E", "PO item value must be less than or equal to PR item value", "", "105%"
        KBMau ws, r, "KHOA_GIA", "ME21N", "", "", "", "", ""
        KBMau ws, r, "PBXX", "ME21N", "", "E", "Condition type PBXX is not allowed (manual entry)", "", ""
        KBMau ws, r, "MIGO_PO", "MIGO", "", "HET_DONG", "PO does not contain any selectable items", "", ""
        KBMau ws, r, "LUU", "MIGO", "", "E", "Posting only possible in periods 2026/10 and 2026/09", "", ""
        KBMau ws, r, "MO_PO", "ME22N", "", "XEM", "PO mo o che do hien thi (macro phai bam Ctrl+F1 de sua)", "", ""
        KBMau ws, r, "MO_PO", "ME22N", "", "KHOA", "Purchase order is currently being processed by user GIALAP2", "", ""
        KBMau ws, r, "LUU", "ME22N", "", "W", "Delivery date is in the past", 1, ""
        KBMau ws, r, "LUU", "ME22N", "", "E", "Enter G/L account for item 00010", "", ""
        ws.Range("A" & KB_ROW1 & ":A" & r - 1).Interior.Color = RGB(255, 242, 204)

        ws.Range("J17").Value = "Y NGHIA"
        ws.Range("J17").Font.Bold = True
        GhiCot ws, 18, 10, Array( _
            "Su_kien:", _
            "  MO_GD       - khi mo giao dich (/nME51N, /nME21N, /nMIGO...)", _
            "  NHAP_DONG   - Enter sau khi nhap dong ME51N (Dieu_kien = ma hang)", _
            "  NCC         - Enter ma NCC o ME21N (Dieu_kien = ma NCC)", _
            "  LAY_DONG_PR - ME21N lay dong PR vao PO (Dieu_kien = ma hang; Gia_tri = gia mac dinh, vd 125000 hoac 105% gia PR - nhu gia info record)", _
            "  KHOA_GIA    - ME21N: o Net Price KHONG cho sua (Dieu_kien = ma hang) -> macro phai nhap gia qua PBXX", _
            "  PBXX        - ME21N: SAP tu choi dieu kien PBXX (Dieu_kien = ma hang)", _
            "  MIGO_PO     - Enter so PO o MIGO (Dieu_kien = so PO). Loai HET_DONG = PO da nhap kho het", _
            "  MO_PO       - ME22N mo PO (Dieu_kien = so PO). XEM = mo o che do hien thi; KHOA = PO dang bi user khac sua", _
            "  LUU         - bam Save / Post (Giao_dich ME51N / ME21N / MIGO)", _
            "Loai:", _
            "  S / I / W / E / A - thong bao o thanh trang thai (MessageType). W: Enter roi Save lai se qua", _
            "  POPUP_TB    - hop thoai thong bao (chi co nut Tiep tuc)", _
            "  POPUP_HOI   - hop thoai hoi Yes / No / Cancel (khi Save: Yes = luu)", _
            "  POPUP_LOI   - hop thoai 'con dong LOI' Save / Edit / Cancel (khi Save)", _
            "So_lan: 1 = chi xay ra lan dau (vd canh bao Save 1 lan). Trong = moi lan.", _
            "Loi co san (khong can bat): ma PR/dong PR khong co, dong PR da dat PO, SL/gia <= 0,", _
            "  thieu Cost center/Quy khi Acct Cat = K, MIGO khong tich OK dong nao, PO khong ton tai,", _
            "  MIGO nhan vuot SL con mo / SL <= 0 / ngay sai dang dd.mm.yyyy...", _
            "Du lieu SAP gia lap: sheet " & SH_DB & " (hoac file o B15) - sua gia/SL/cot Xoa cua PO de gia lap ai do sua PO tren SAP")
        ws.Columns("A").ColumnWidth = 48
        ws.Columns("B").ColumnWidth = 14
        ws.Columns("C:E").ColumnWidth = 14
        ws.Columns("F").ColumnWidth = 60
        ws.Columns("G:H").ColumnWidth = 12
    End If
    Set SheetCfg = ws
End Function

Private Sub GhiCot(ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal a As Variant)
    Dim i As Long
    For i = 0 To UBound(a)
        ws.Cells(r + i, c).Value = a(i)
    Next i
End Sub

Private Sub KBMau(ws As Worksheet, ByRef r As Long, ByVal sk As String, ByVal gd As String, ByVal dk As String, _
                 ByVal loai As String, ByVal nd As String, ByVal solan As Variant, ByVal gt As Variant)
    ws.Range("A" & r & ":H" & r).Value = Array("", sk, gd, dk, loai, nd, solan, gt)
    r = r + 1
End Sub

Private Function SheetLog() As Worksheet
    Dim ws As Worksheet, moi As Boolean
    Set ws = LaySheet(SH_LOG, moi)
    If moi Or Trim(CStr(ws.Range("A2").Value)) <> "STT" Then
        ws.Range("A1").Value = "NHAT KY SAP GIA LAP - moi thao tac macro lam tren SAP (" & GL_VERSION & ")"
        ws.Range("A1").Font.Bold = True
        ws.Range("A2:G2").Value = Array("STT", "Thoi_gian", "Giao_dich", "Thao_tac", "Doi_tuong (ID SAP)", "Gia_tri", "Ghi_chu")
        ws.Range("A2:G2").Font.Bold = True
        ws.Range("A2:G2").Interior.Color = RGB(221, 235, 247)
        ws.Columns("A").ColumnWidth = 6
        ws.Columns("B").ColumnWidth = 10
        ws.Columns("C").ColumnWidth = 10
        ws.Columns("D").ColumnWidth = 22
        ws.Columns("E").ColumnWidth = 48
        ws.Columns("F").ColumnWidth = 40
        ws.Columns("G").ColumnWidth = 40
    End If
    Set SheetLog = ws
End Function

' Cot sheet DB
Private Function SheetDB() As Worksheet
    Dim ws As Worksheet, moi As Boolean
    Set ws = LaySheet(SH_DB, moi)
    If moi Or Trim(CStr(ws.Range("A1").Value)) <> "Loai" Then
        ws.Range("A1:P1").Value = Array("Loai", "So_CT", "Item", "Ma_hang", "SL", "Don_gia", "NCC", "PR", "Item_PR", _
                                        "SL_da_nhap", "Xoa (L)", "Cost_center", "Quy", "Loai_CT", "Thoi_diem", "Ghi_chu")
        ws.Range("A1:P1").Font.Bold = True
        ws.Range("A1:P1").Interior.Color = RGB(221, 235, 247)
        ws.Columns("B:D").NumberFormat = "@"
        ws.Columns("H:I").NumberFormat = "@"
        ws.Columns("A:P").ColumnWidth = 12
    End If
    Set SheetDB = ws
End Function

Private Sub DocCauHinh()
    Dim ws As Worksheet
    Set ws = SheetCfg()
    cDec = Trim(CStr(ws.Range("B4").Value))
    If cDec <> "." Then cDec = ","
    cGiaLe = SoNguyen(ws.Range("B5").Value, 0)
    cVisPOThu = SoNguyen(ws.Range("B6").Value, 8)
    cVisPOMo = SoNguyen(ws.Range("B7").Value, 3)
    cVisMigo = SoNguyen(ws.Range("B8").Value, 5)
    cVisCond = SoNguyen(ws.Range("B9").Value, 4)
    If cVisPOThu < 1 Then cVisPOThu = 1
    If cVisPOMo < 1 Then cVisPOMo = 1
    If cVisMigo < 1 Then cVisMigo = 1
    If cVisCond < 2 Then cVisCond = 2
    cLogDoc = (UCase(Trim(CStr(ws.Range("B13").Value))) = "CO")
    cSys = Trim(CStr(ws.Range("B14").Value)): If cSys = "" Then cSys = "GIA"
    cClient = Trim(CStr(ws.Range("C14").Value)): If cClient = "" Then cClient = "100"
    cUser = Trim(CStr(ws.Range("D14").Value)): If cUser = "" Then cUser = "GIALAP"
    cDbFile = Trim(CStr(ws.Range("B15").Value))
End Sub

Private Function SoNguyen(ByVal v As Variant, ByVal macDinh As Long) As Long
    If IsNumeric(v) And Trim(CStr(v)) <> "" Then
        SoNguyen = CLng(v)
    Else
        SoNguyen = macDinh
    End If
End Function

' Tim dong kich ban dang BAT khop su kien. Tra ve True + loai/noi dung/gia tri.
Private Function KB(ByVal suKien As String, ByVal doiTuong As String, ByRef loai As String, _
                    ByRef nd As String, ByRef gt As Variant) As Boolean
    Dim ws As Worksheet, r As Long, gd As String, dk As String, solan As Variant, daDung As Long
    If gDem Is Nothing Then Set gDem = New Collection
    Set ws = SheetCfg()
    For r = KB_ROW1 To KB_ROWN
        If UCase(Trim(CStr(ws.Cells(r, 1).Value))) = "X" And UCase(Trim(CStr(ws.Cells(r, 2).Value))) = suKien Then
            gd = UCase(Trim(CStr(ws.Cells(r, 3).Value)))
            dk = UCase(Trim(CStr(ws.Cells(r, 4).Value)))
            If (gd = "" Or gd = gTx) And (dk = "" Or InStr(1, UCase(doiTuong), dk) > 0) Then
                solan = ws.Cells(r, 7).Value
                daDung = 0
                On Error Resume Next
                daDung = gDem.Item("r" & r)
                On Error GoTo 0
                If Not (IsNumeric(solan) And Trim(CStr(solan)) <> "") Or daDung < Val(CStr(solan)) Then
                    On Error Resume Next
                    gDem.Remove "r" & r
                    On Error GoTo 0
                    gDem.Add daDung + 1, "r" & r
                    loai = UCase(Trim(CStr(ws.Cells(r, 5).Value)))
                    nd = Trim(CStr(ws.Cells(r, 6).Value))
                    gt = ws.Cells(r, 8).Value
                    If InStr(CStr(ws.Cells(r, 8).NumberFormat), "%") > 0 And IsNumeric(gt) Then gt = CStr(CDbl(gt) * 100) & "%"
                    GL_LogDong "Kich ban", "", "Dong " & r & ": " & suKien & " / " & loai & " / " & nd
                    KB = True
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

' ---------------------------------------------------------------- nhat ky
Public Sub GL_Log(ByVal thaoTac As String, ByVal c As Object, ByVal giaTri As String)
    Dim s As String
    On Error Resume Next
    s = RutGonId(c)
    On Error GoTo 0
    GL_LogDong thaoTac, s, giaTri
End Sub

Public Sub GL_LogDong(ByVal thaoTac As String, ByVal doiTuong As String, ByVal giaTri As String, Optional ByVal ghiChu As Variant)
    Dim ws As Worksheet
    If IsMissing(ghiChu) Then ghiChu = ""
    On Error GoTo Thoat
    Set ws = SheetLog()
    If gLogRow < 3 Then
        gLogRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If gLogRow < 3 Then gLogRow = 3
    End If
    ws.Cells(gLogRow, 1).Value = gLogRow - 2
    ws.Cells(gLogRow, 2).Value = Format(Now, "hh:mm:ss")
    ws.Cells(gLogRow, 3).Value = gTx
    ws.Cells(gLogRow, 4).Value = "'" & thaoTac
    ws.Cells(gLogRow, 5).Value = "'" & doiTuong
    ws.Cells(gLogRow, 6).Value = "'" & giaTri
    ws.Cells(gLogRow, 7).Value = "'" & ghiChu
    If Left(thaoTac, 7) = "SAP bao" Then
        If InStr(thaoTac, "[E]") > 0 Or InStr(thaoTac, "[A]") > 0 Then
            ws.Range(ws.Cells(gLogRow, 1), ws.Cells(gLogRow, 7)).Interior.Color = RGB(255, 199, 206)
        ElseIf InStr(thaoTac, "[W]") > 0 Then
            ws.Range(ws.Cells(gLogRow, 1), ws.Cells(gLogRow, 7)).Interior.Color = RGB(255, 235, 156)
        ElseIf InStr(thaoTac, "[S]") > 0 Then
            ws.Range(ws.Cells(gLogRow, 1), ws.Cells(gLogRow, 7)).Interior.Color = RGB(198, 239, 206)
        End If
    ElseIf Left(thaoTac, 5) = "Popup" Or Left(thaoTac, 3) = "===" Then
        ws.Range(ws.Cells(gLogRow, 1), ws.Cells(gLogRow, 7)).Interior.Color = RGB(221, 235, 247)
    End If
    gLogRow = gLogRow + 1
Thoat:
End Sub

Private Function RutGonId(ByVal c As Object) As String
    Dim s As String, a() As String, n As Long
    s = c.mId
    If Left(s, Len(SesId()) + 1) = SesId() & "/" Then s = Mid(s, Len(SesId()) + 2)
    a = Split(s, "/")
    n = UBound(a)
    If c.Kind = "CELL" Then
        RutGonId = a(0) & "/.../" & a(n - 1) & "/" & a(n) & " (dong " & (c.Hang + 1) & ")"
    ElseIf n >= 3 Then
        RutGonId = a(0) & "/.../" & a(n)
    Else
        RutGonId = s
    End If
End Function


' ================================================================================
'  3. DUNG MAN HINH (cay control voi DUNG ID nhu SAP that)
' ================================================================================

Private Sub KhoiTao()
    Set gApp = TaoCtl("APP", "/app", "app", "GuiApplication")
    Set gConn = TaoCtl("CONN", "/app/con[0]", "con[0]", "GuiConnection")
    Set gSes = TaoCtl("SESSION", SesId(), "ses[" & gPhien & "]", "GuiSession")
    Set gInfo = TaoCtl("INFO", SesId() & "/info", "info", "GuiSessionInfo")
    gApp.Kids.Them gConn
    Set gConn.Cha = gApp
    gConn.Kids.Them gSes
    Set gSes.Cha = gConn
    gPop = ""
    gFocusTbl = ""
    poDetailInit = False
    SheetDB
    DocCauHinh
    MoGiaoDich "SESSION_MANAGER"
    XayManHinh
End Sub

Private Function TaoCtl(ByVal sKind As String, ByVal sId As String, ByVal nm As String, ByVal tp As String) As SapGL_Ctl
    Dim c As New SapGL_Ctl
    c.Kind = sKind
    c.mId = sId
    c.mName = nm
    c.mType = tp
    Set TaoCtl = c
End Function

Private Sub DangKy(ByVal c As SapGL_Ctl)
    On Error Resume Next
    gReg.Remove LCase(c.mId)
    On Error GoTo 0
    gReg.Add c, LCase(c.mId)
End Sub

Private Function TimDK(ByVal sId As String) As SapGL_Ctl
    On Error Resume Next
    Set TimDK = gReg.Item(LCase(sId))
    On Error GoTo 0
End Function

' Ten + loai control suy tu doan ID (vd "ctxtCOBL-KOSTL" -> COBL-KOSTL / GuiCTextField)
Private Sub KieuTuDoan(ByVal seg As String, ByRef nm As String, ByRef tp As String)
    Dim p As Variant, i As Long, pre As Variant, typ As Variant
    pre = Array("ssub", "sub", "ctxt", "txt", "cmb", "btn", "chk", "rad", "lbl", "tabs", "tabp", "tbl", "cntl", "shellcont", "shell")
    typ = Array("GuiScrollContainer", "GuiSimpleContainer", "GuiCTextField", "GuiTextField", "GuiComboBox", "GuiButton", _
                "GuiCheckBox", "GuiRadioButton", "GuiLabel", "GuiTabStrip", "GuiTab", "GuiTableControl", "GuiCustomControl", _
                "GuiContainerShell", "GuiShell")
    nm = seg: tp = "GuiComponent"
    If Left(seg, 4) = "wnd[" Then
        nm = seg
        tp = IIf(seg = "wnd[0]", "GuiMainWindow", "GuiModalWindow")
        Exit Sub
    End If
    Select Case seg
        Case "usr": tp = "GuiUserArea": Exit Sub
        Case "sbar": tp = "GuiStatusbar": Exit Sub
        Case "titl": tp = "GuiTitlebar": Exit Sub
        Case "okcd": tp = "GuiOkCodeField": Exit Sub
        Case "mbar": tp = "GuiMenubar": Exit Sub
    End Select
    If Left(seg, 5) = "tbar[" Then tp = "GuiToolbar": Exit Sub
    If Left(seg, 4) = "btn[" Then tp = "GuiButton": Exit Sub
    For i = 0 To UBound(pre)
        If Left(seg, Len(pre(i))) = pre(i) Then
            tp = typ(i)
            If pre(i) = "shellcont" Or pre(i) = "shell" Then
                nm = seg
            Else
                nm = Mid(seg, Len(pre(i)) + 1)
                p = InStr(nm, ":")
                If p > 0 Then nm = Left(nm, p - 1)
            End If
            Exit Sub
        End If
    Next i
End Sub

' Them 1 control con ngay duoi cha
Private Function Nut(ByVal cha As SapGL_Ctl, ByVal seg As String, Optional ByVal sKind As Variant, _
                     Optional ByVal sRef As Variant, Optional ByVal txt As Variant) As SapGL_Ctl
    Dim c As SapGL_Ctl, nm As String, tp As String
    If IsMissing(sKind) Then sKind = "CONT"
    If IsMissing(sRef) Then sRef = ""
    If IsMissing(txt) Then txt = ""
    Set c = TimDK(cha.mId & "/" & seg)
    If c Is Nothing Then
        KieuTuDoan seg, nm, tp
        Set c = TaoCtl(sKind, cha.mId & "/" & seg, nm, tp)
        Set c.Cha = cha
        cha.Kids.Them c
        DangKy c
    End If
    If sKind <> "CONT" Then c.Kind = sKind
    If sRef <> "" Then c.mRef = sRef
    If txt <> "" Then c.mText = txt
    Set Nut = c
End Function

' Them ca duong dan (tu tao cac subscreen trung gian)
Private Function Duong(ByVal cha As SapGL_Ctl, ByVal path As String, Optional ByVal sKind As Variant, _
                       Optional ByVal sRef As Variant, Optional ByVal txt As Variant) As SapGL_Ctl
    Dim a() As String, i As Long, c As SapGL_Ctl
    If IsMissing(sKind) Then sKind = "CONT"
    If IsMissing(sRef) Then sRef = ""
    If IsMissing(txt) Then txt = ""
    a = Split(path, "/")
    Set c = cha
    For i = 0 To UBound(a) - 1
        Set c = Nut(c, a(i))
    Next i
    Set Duong = Nut(c, a(UBound(a)), sKind, sRef, txt)
End Function

Private Sub XayManHinh()
    Dim w0 As SapGL_Ctl, tb As SapGL_Ctl, usr As SapGL_Ctl
    Set gReg = New Collection
    Set gSes.Kids = New SapGL_List
    DangKy gSes

    Set w0 = Nut(gSes, "wnd[0]", "WND")
    Nut w0, "mbar"
    Set tb = Nut(w0, "tbar[0]")
    Nut tb, "okcd", "OKCD"
    Nut(tb, "btn[0]", "BTN", "TB:ENTER").Tooltip = "Enter (Enter)"
    Nut(tb, "btn[11]", "BTN", "TB:SAVE").Tooltip = "Save (Ctrl+S)"
    Nut(tb, "btn[3]", "BTN", "TB:BACK").Tooltip = "Back (F3)"
    Nut(tb, "btn[12]", "BTN", "TB:BACK").Tooltip = "Cancel (F12)"
    Nut w0, "titl", "LBL"
    Set tb = Nut(w0, "tbar[1]")
    If gTx = "ME21N" Or gTx = "ME22N" Or gTx = "ME23N" Then
        Nut(tb, "btn[17]", "BTN", "TB:OTHERPO", "Other Purchase Order").Tooltip = "Other Purchase Order (Shift+F5)"
    End If
    If gTx = "ME22N" Then Nut(tb, "btn[25]", "BTN", "TB:DOICHEDO", "Display <-> Change").Tooltip = "Display <-> Change (Ctrl+F1)"
    If gTx = "MIGO" Then Nut(tb, "btn[7]", "BTN", "TB:CHECK", "Check").Tooltip = "Check (Ctrl+F7)"
    Set usr = Nut(w0, "usr")
    Select Case gTx
        Case "ME51N": XayME51N usr
        Case "ME21N", "ME22N", "ME23N": XayME2N usr
        Case "MIGO": XayMIGO usr
        Case Else: Nut usr, "lblEASY_ACCESS", "LBL", "", "SAP Easy Access (gia lap)"
    End Select
    Nut w0, "sbar", "SBAR"

    If gPop <> "" Then XayPopup
End Sub

Private Sub XayME51N(ByVal usr As SapGL_Ctl)
    Dim root As String, d As String, ts As String, t7 As String, acc As String, t As Variant, a() As String
    root = "subSUB0:SAPLMEGUI:0016"
    Duong usr, root & "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/cmbMEREQ_TOPLINE-BSART", "CMB", "PR:BSART"
    Duong usr, root & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell", "GRID", "PR"
    Duong(usr, root & "/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON", "BTN", "PR:DETAIL").Tooltip = "Item Detail (Expand/Collapse)"
    If Not prDetail Then Exit Sub
    d = root & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301"
    Duong usr, d & "/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST", "CMB", "PR:ITEMLIST"
    ts = d & "/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL"
    For Each t In Array("TABREQDT1|Material Data", "TABREQDT2|Quantities/Dates", "TABREQDT3|Valuation", _
                        "TABREQDT7|Account Assignment", "TABREQDT8|Source of Supply")
        a = Split(t, "|")
        Duong usr, ts & "/tabp" & a(0), "TAB", "PR:TAB:" & a(0), a(1)
    Next t
    If prTab = "TABREQDT7" And prItems.Count > 0 Then
        t7 = ts & "/tabpTABREQDT7/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101"
        Duong usr, t7 & "/subSUB1:SAPLMEACCTVI:1200/cmbMEACCT1200-KNTTP", "CMB", "PR:KNTTP"
        If CurPR().Knttp <> "" Then
            acc = t7 & "/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101"
            Duong usr, acc & "/ctxtCOBL-KOSTL", "FLD", "PR:KOSTL"
            Duong usr, acc & "/ctxtCOBL-GEBER", "FLD", "PR:GEBER"
        End If
    End If
End Sub

Private Sub XayME2N(ByVal usr As SapGL_Ctl)
    Dim root As String, hdr As String, d As String, ts As String, t As Variant, a() As String
    ' ME21N / ME22N (che do sua) dung 0013, ME23N / ME22N (che do hien thi) dung 0019 - macro tim theo ten nen so nao cung duoc
    root = IIf(gTx = "ME21N" Or (gTx = "ME22N" And Not poXem), "subSUB0:SAPLMEGUI:0013", "subSUB0:SAPLMEGUI:0019")
    hdr = root & "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105"
    Duong usr, hdr & "/cmbMEPO_TOPLINE-BSART", "CMB", "PO:BSART"
    Duong usr, hdr & "/ctxtMEPO_TOPLINE-SUPERFIELD", "FLD", "PO:VENDOR"
    Duong usr, root & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211", "TBL", "PO"
    Duong(usr, root & "/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON", "BTN", "PO:DETAIL").Tooltip = "Item Detail (Expand/Collapse)"
    If Not poDetail Then Exit Sub
    d = root & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301"
    Duong usr, d & "/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST", "CMB", "PO:ITEMLIST"
    ts = d & "/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL"
    For Each t In Array("TABIDT1|Material Data", "TABIDT2|Quantities/Weights", "TABIDT5|Delivery Schedule", "TABIDT9|Conditions")
        a = Split(t, "|")
        Duong usr, ts & "/tabp" & a(0), "TAB", "PO:TAB:" & a(0), a(1)
    Next t
    If poTab = "TABIDT9" And poItems.Count > 0 Then
        Duong usr, ts & "/tabpTABIDT9/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN", "TBL", "COND"
    End If
End Sub

Private Sub XayMIGO(ByVal usr As SapGL_Ctl)
    Dim c As String, fl As String
    c = "ssubSUB_MAIN_CARRIER:SAPLMIGO:0003"
    fl = c & "/subSUB_FIRSTLINE:SAPLMIGO:0011"
    Duong usr, fl & "/cmbGODYNPRO-ACTION", "CMB", "MG:ACTION"
    Duong usr, fl & "/cmbGODYNPRO-REFDOC", "CMB", "MG:REFDOC"
    Duong usr, fl & "/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/ctxtGODYNPRO-PO_NUMBER", "FLD", "MG:PO"
    Duong usr, c & "/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM", "TBL", "MIGO"
    ' header (tab General): ngay chung tu, ngay hach toan, so phieu giao hang
    fl = c & "/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110"
    Duong usr, fl & "/ctxtGOHEAD-BLDAT", "FLD", "MG:BLDAT"
    Duong usr, fl & "/ctxtGOHEAD-BUDAT", "FLD", "MG:BUDAT"
    Duong usr, fl & "/txtGOHEAD-LFSNR", "FLD", "MG:LFSNR"
End Sub

Private Sub XayPopup()
    Dim w As SapGL_Ctl, tb As SapGL_Ctl, usr As SapGL_Ctl
    Set w = Nut(gSes, "wnd[1]", "WND")
    Set tb = Nut(w, "tbar[0]")
    Set usr = Nut(w, "usr")
    Select Case gPop
        Case "TB"
            Nut(usr, "txtMESSTXT1", "FLD", "", gPopText).mChangeable = False
            Nut(tb, "btn[0]", "BTN", "POP:OK").Tooltip = "Continue (Enter)"
        Case "HOI"
            Nut(usr, "txtSPOP-TEXTLINE1", "FLD", "", gPopText).mChangeable = False
            Nut usr, "btnSPOP-OPTION1", "BTN", "POP:YES", "Yes"
            Nut usr, "btnSPOP-OPTION2", "BTN", "POP:NO", "No"
            Nut usr, "btnSPOP-OPTION_CAN", "BTN", "POP:CAN", "Cancel"
        Case "LOI"
            Nut(usr, "txtSPOP-TEXTLINE1", "FLD", "", gPopText).mChangeable = False
            Nut usr, "btnSPOP-OPTION1", "BTN", "POP:YES", "Save"
            Nut usr, "btnSPOP-OPTION2", "BTN", "POP:NO", "Edit"
            Nut usr, "btnSPOP-OPTION_CAN", "BTN", "POP:CAN", "Cancel"
        Case "CHON_PO"
            Duong usr, "subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN", "FLD", "POP:EBELN"
            Nut(tb, "btn[0]", "BTN", "POP:CHONPO", "Other Document").Tooltip = "Other Document (Enter)"
            Nut(tb, "btn[12]", "BTN", "POP:CAN").Tooltip = "Cancel (F12)"
    End Select
End Sub

Private Sub MoPopup(ByVal kieu As String, ByVal tieuDe As String, ByVal noiDung As String, ByVal viec As String)
    gPop = kieu
    gPopTitle = tieuDe
    gPopText = noiDung
    gPopViec = viec
    gPopSel = ""
    GL_LogDong "Popup mo", "wnd[1]", tieuDe & ": " & noiDung
End Sub

Private Sub DongPopup()
    gPop = ""
    gPopViec = ""
End Sub


' ================================================================================
'  4. CAC THAO TAC MACRO GOI (findById, Text, key, press, sendVKey...)
' ================================================================================

Public Function GL_Children(ByVal c As SapGL_Ctl) As Object
    Set GL_Children = c.Kids
End Function

Public Function GL_FindById(ByVal goc As SapGL_Ctl, ByVal sId As String, ByVal bRaise As Boolean) As Object
    Dim full As String, c As SapGL_Ctl
    sId = Trim(sId)
    If Left(sId, 1) = "/" Then
        full = sId
    Else
        full = goc.mId & "/" & sId
    End If
    Set c = TimDK(full)
    If c Is Nothing Then
        If cLogDoc Then GL_LogDong "findById (khong co)", full, ""
        If bRaise Then Err.Raise 619, NGUON, "The control could not be found by id."
    End If
    Set GL_FindById = c
End Function

Public Function GL_FindByName(ByVal goc As SapGL_Ctl, ByVal nm As String, ByVal tp As String) As Object
    Dim i As Long, ch As SapGL_Ctl, r As Object
    For i = 0 To goc.Kids.Count - 1
        Set ch = goc.Kids.ElementAt(i)
        If LCase(ch.mName) = LCase(nm) And (tp = "" Or LCase(ch.mType) = LCase(tp)) Then
            Set GL_FindByName = ch
            Exit Function
        End If
        Set r = GL_FindByName(ch, nm, tp)
        If Not r Is Nothing Then
            Set GL_FindByName = r
            Exit Function
        End If
    Next i
End Function

Public Function GL_FindAllByName(ByVal goc As SapGL_Ctl, ByVal nm As String, ByVal tp As String) As Object
    Dim res As New SapGL_List
    GomTheoTen goc, nm, tp, res
    Set GL_FindAllByName = res
End Function

Private Sub GomTheoTen(ByVal goc As SapGL_Ctl, ByVal nm As String, ByVal tp As String, ByVal res As SapGL_List)
    Dim i As Long, ch As SapGL_Ctl
    For i = 0 To goc.Kids.Count - 1
        Set ch = goc.Kids.ElementAt(i)
        If LCase(ch.mName) = LCase(nm) And (tp = "" Or LCase(ch.mType) = LCase(tp)) Then res.Them ch
        GomTheoTen ch, nm, tp, res
    Next i
End Sub

Public Function GL_Info() As Object
    Set GL_Info = gInfo
End Function

Public Function GL_InfoVal(ByVal ten As String) As String
    Select Case ten
        Case "TRANSACTION": GL_InfoVal = gTx
        Case "SYSTEMNAME": GL_InfoVal = cSys
        Case "CLIENT": GL_InfoVal = cClient
        Case "USER": GL_InfoVal = cUser
        Case "PROGRAM"
            Select Case gTx
                Case "ME51N", "ME21N", "ME22N", "ME23N": GL_InfoVal = "SAPLMEGUI"
                Case "MIGO": GL_InfoVal = "SAPLMIGO"
                Case Else: GL_InfoVal = "SAPLSMTR_NAVIGATION"
            End Select
    End Select
End Function

' Chan thao tac tren cua so chinh khi dang co popup (SAP that cung khong cho)
Private Sub ChanKhiCoPopup(ByVal c As SapGL_Ctl)
    If gPop <> "" And InStr(c.mId, "/wnd[0]") > 0 Then
        GL_LogDong "LOI thao tac", RutGonId(c), "SAP dang mo popup wnd[1] '" & gPopTitle & "' - khong thao tac duoc cua so chinh"
        Err.Raise 621, NGUON, "The control is not available: modal window wnd[1] '" & gPopTitle & "' is open (gia lap)"
    End If
End Sub

Private Sub ChanKhongSuaDuoc(ByVal c As SapGL_Ctl)
    If Not GL_Changeable(c) Then
        GL_LogDong "LOI thao tac", RutGonId(c), "o khong cho sua (Changeable = False)"
        Err.Raise 622, NGUON, "The field is not changeable (Changeable = False) (gia lap)"
    End If
End Sub

' ---------------------------------------------------------------- Text
Public Function GL_GetText(ByVal c As SapGL_Ctl) As String
    Dim s As String
    Select Case c.Kind
        Case "OKCD": s = gOkcd
        Case "SBAR": s = gSbT
        Case "WND"
            If Right(c.mId, 6) = "wnd[0]" Then
                s = gTitle
            Else
                s = gPopTitle
            End If
        Case "LBL"
            If c.mName = "titl" Then
                s = gTitle
            Else
                s = c.mText
            End If
        Case "FLD"
            Select Case c.mRef
                Case "PR:KOSTL": s = CurPR().Kostl
                Case "PR:GEBER": s = CurPR().Geber
                Case "PO:VENDOR": s = poVendor
                Case "MG:PO": s = mgPO
                Case "MG:BLDAT": s = mgBldat
                Case "MG:BUDAT": s = mgBudat
                Case "MG:LFSNR": s = mgLfsnr
                Case "POP:EBELN": s = gPopSel
                Case Else: s = c.mText
            End Select
        Case "CMB": s = GiaTriEntry(c.mRef, GL_GetKey(c))
        Case "CELL": s = TextO(c)
        Case "ENTRY": s = c.mText
        Case "INFO", "SESSION": s = ""
        Case Else: s = c.mText
    End Select
    If cLogDoc And (c.Kind = "FLD" Or c.Kind = "CELL" Or c.Kind = "SBAR") Then GL_Log "doc Text", c, s
    GL_GetText = s
End Function

Public Sub GL_SetText(ByVal c As SapGL_Ctl, ByVal v As String)
    Dim ok As Boolean, x As Double
    ChanKhiCoPopup c
    GL_Log "Text =", c, v
    Select Case c.Kind
        Case "OKCD": gOkcd = v
        Case "FLD"
            ChanKhongSuaDuoc c
            Select Case c.mRef
                Case "PR:KOSTL": CurPR().Kostl = UCase(Trim(v))
                Case "PR:GEBER": CurPR().Geber = UCase(Trim(v))
                Case "PO:VENDOR": poVendor = UCase(Trim(v))
                Case "MG:PO": mgPO = Trim(v)
                Case "MG:BLDAT": mgBldat = Trim(v)
                Case "MG:BUDAT": mgBudat = Trim(v)
                Case "MG:LFSNR": mgLfsnr = Left(Trim(v), 16)
                Case "POP:EBELN": gPopSel = Trim(v)
                Case Else: c.mText = v
            End Select
        Case "CELL"
            SetO c, v
        Case "SBAR", "WND", "LBL", "TBL", "GRID"
            Err.Raise 622, NGUON, "The property Text is read-only for " & c.mType & " (gia lap)"
        Case Else
            c.mText = v
    End Select
End Sub

' ---------------------------------------------------------------- combobox
Private Function DsEntry(ByVal sRef As String) As Variant
    Dim a() As String, i As Long, n As Long
    Select Case sRef
        Case "PR:BSART": DsEntry = Array("NB|Purchase Requisition", "ZEX|Expense PR", "ZSV|Service PR")
        Case "PO:BSART": DsEntry = Array("NB|Standard PO", "ZEXP|Expense PO", "FO|Framework Order")
        Case "PR:KNTTP": DsEntry = Array(" | ", "A|Asset", "F|Order", "K|Cost center", "P|Project", "U|Unknown")
        Case "MG:ACTION": DsEntry = Array("A01|Goods Receipt", "A04|Display", "A07|Goods Issue", "A08|Transfer Posting")
        Case "MG:REFDOC": DsEntry = Array("R01|Purchase Order", "R02|Material Document", "R10|Other")
        Case "PR:ITEMLIST", "PO:ITEMLIST"
            Dim col As Collection, d As SapGL_Dong
            If sRef = "PR:ITEMLIST" Then
                Set col = prItems
            Else
                Set col = poItems
            End If
            n = col.Count
            If n = 0 Then
                DsEntry = Array()
                Exit Function
            End If
            ReDim a(0 To n - 1)
            For i = 1 To n
                Set d = col.Item(i)
                a(i - 1) = Right(Space(4) & CStr(i), 4) & "|[ " & IIf(sRef = "PR:ITEMLIST", d.Bnfpo, d.Ebelp) & " ] " & d.Mat
            Next i
            DsEntry = a
        Case Else: DsEntry = Array()
    End Select
End Function

Private Function GiaTriEntry(ByVal sRef As String, ByVal k As String) As String
    Dim e As Variant, a() As String
    For Each e In DsEntry(sRef)
        a = Split(e, "|")
        If a(0) = k Or (Trim(a(0)) = "" And Trim(k) = "") Then GiaTriEntry = a(1): Exit Function
    Next e
End Function

Public Function GL_Entries(ByVal c As SapGL_Ctl) As Object
    Dim res As New SapGL_List, e As Variant, a() As String, x As SapGL_Ctl, i As Long
    For Each e In DsEntry(c.mRef)
        a = Split(e, "|")
        Set x = TaoCtl("ENTRY", c.mId & "/entry[" & i & "]", "", "GuiComboBoxEntry")
        x.mKey = a(0)
        x.mText = a(1)
        res.Them x
        i = i + 1
    Next e
    Set GL_Entries = res
End Function

Public Function GL_GetKey(ByVal c As SapGL_Ctl) As String
    Select Case c.Kind
        Case "CMB"
            Select Case c.mRef
                Case "PR:BSART": GL_GetKey = prBsart
                Case "PO:BSART": GL_GetKey = poBsart
                Case "PR:KNTTP": GL_GetKey = CurPR().Knttp
                Case "MG:ACTION": GL_GetKey = mgAction
                Case "MG:REFDOC": GL_GetKey = mgRefdoc
                Case "PR:ITEMLIST": If prItems.Count > 0 Then GL_GetKey = Right(Space(4) & CStr(prCur + 1), 4)
                Case "PO:ITEMLIST": If poItems.Count > 0 Then GL_GetKey = Right(Space(4) & CStr(poCur + 1), 4)
                Case Else: GL_GetKey = c.mKey
            End Select
        Case Else
            GL_GetKey = c.mKey
    End Select
End Function

Public Sub GL_SetKey(ByVal c As SapGL_Ctl, ByVal v As String)
    Dim e As Variant, a() As String, found As Boolean
    ChanKhiCoPopup c
    GL_Log "key =", c, "'" & v & "'"
    If c.Kind <> "CMB" Then
        c.mKey = v
        Exit Sub
    End If
    ChanKhongSuaDuoc c
    For Each e In DsEntry(c.mRef)
        a = Split(e, "|")
        If a(0) = v Or (Trim(a(0)) = "" And Trim(v) = "") Then found = True: Exit For
    Next e
    If Not found Then
        GL_LogDong "LOI thao tac", RutGonId(c), "khong co gia tri '" & v & "' trong danh sach"
        Err.Raise 623, NGUON, "The method got an invalid argument: entry '" & v & "' not found in combobox " & c.mName & " (gia lap)"
    End If
    Select Case c.mRef
        Case "PR:BSART": prBsart = v
        Case "PO:BSART": poBsart = v
        Case "PR:KNTTP": CurPR().Knttp = Trim(v)
        Case "MG:ACTION": mgAction = v
        Case "MG:REFDOC": mgRefdoc = v
        Case "PR:ITEMLIST": prCur = CLng(Trim(v)) - 1
        Case "PO:ITEMLIST": poCur = CLng(Trim(v)) - 1: poCondPos = 0
    End Select
    VongXuLy                                  ' doi combobox = SAP xu ly man hinh
End Sub

Public Function GL_Changeable(ByVal c As SapGL_Ctl) As Boolean
    Select Case c.Kind
        Case "CELL": GL_Changeable = SuaDuocO(c)
        Case "FLD"
            Select Case c.mRef
                Case "PO:VENDOR": GL_Changeable = Not poXem
                Case Else: GL_Changeable = c.mChangeable
            End Select
        Case "CMB"
            GL_Changeable = Not ((c.mRef = "PO:BSART") And poXem)
        Case "SBAR", "WND", "LBL"
            GL_Changeable = False
        Case Else
            GL_Changeable = c.mChangeable
    End Select
End Function

Public Function GL_GetSelected(ByVal c As SapGL_Ctl) As Boolean
    If c.Kind = "CELL" And c.mRef = "MIGO" Then
        If c.Hang < mgItems.Count Then GL_GetSelected = mgItems.Item(c.Hang + 1).Take
    Else
        GL_GetSelected = (c.mText = "X")
    End If
End Function

Public Sub GL_SetSelected(ByVal c As SapGL_Ctl, ByVal v As Boolean)
    ChanKhiCoPopup c
    GL_Log "Selected =", c, CStr(v)
    If c.Kind = "CELL" And c.mRef = "MIGO" Then
        ChanKhongSuaDuoc c
        mgItems.Item(c.Hang + 1).Take = v
    Else
        c.mText = IIf(v, "X", "")
    End If
End Sub

Public Function GL_IconName(ByVal c As SapGL_Ctl) As String
    If c.Kind = "CELL" And c.mRef = "PO" And c.mName = "MEPO1211-LOEKZ" Then
        If c.Hang < poItems.Count Then
            If poItems.Item(c.Hang + 1).Loekz <> "" Then GL_IconName = "@11@"
        End If
    End If
End Function

Public Function GL_MessageType(ByVal c As SapGL_Ctl) As String
    If c.Kind = "SBAR" Then GL_MessageType = gSbM
End Function

Public Sub GL_SetFocus(ByVal c As SapGL_Ctl)
    If c.Kind = "CELL" Then gFocusTbl = c.mRef
    If cLogDoc Then GL_Log "SetFocus", c, ""
End Sub

Public Sub GL_SelectTab(ByVal c As SapGL_Ctl)
    Dim code As String
    ChanKhiCoPopup c
    GL_Log "Select (tab)", c, c.mText
    If c.Kind <> "TAB" Then Err.Raise 438, NGUON, "Object doesn't support this property or method (Select) (gia lap)"
    code = Mid(c.mRef, InStrRev(c.mRef, ":") + 1)
    If Left(c.mRef, 3) = "PR:" Then
        prTab = code
    Else
        poTab = code
        poCondPos = 0
    End If
    VongXuLy
End Sub

Public Sub GL_CloseWnd(ByVal c As SapGL_Ctl)
    GL_Log "Close", c, ""
    If c.Kind = "WND" And Right(c.mId, 6) <> "wnd[0]" Then
        BamPopup "POP:CAN"
    End If
End Sub

Public Sub GL_SendVKey(ByVal c As SapGL_Ctl, ByVal k As Long)
    GL_Log "sendVKey " & k, c, TenPhim(k)
    If Right(c.mId, 6) = "wnd[0]" Or c.Kind <> "WND" Then
        ChanKhiCoPopup c
        Select Case k
            Case 0: VongXuLy
            Case 11: BamLuu
            Case 3, 12, 15: MoGiaoDich "SESSION_MANAGER": XayManHinh
            Case 81, 82: CuonTrang (k = 82)
            Case 17
                If gTx = "ME21N" Or gTx = "ME22N" Or gTx = "ME23N" Then
                    MoPopup "CHON_PO", "Select Document", "", "CHON_PO"
                    XayManHinh
                End If
            Case 25
                If gTx = "ME22N" Then
                    DoiCheDoME22N
                Else
                    VongXuLy
                End If
            Case Else: VongXuLy
        End Select
    Else
        Select Case k
            Case 0
                Select Case gPop
                    Case "TB": BamPopup "POP:OK"
                    Case "HOI", "LOI": BamPopup "POP:YES"
                    Case "CHON_PO": BamPopup "POP:CHONPO"
                End Select
            Case 12, 15, 3: BamPopup "POP:CAN"
        End Select
    End If
End Sub

Private Function TenPhim(ByVal k As Long) As String
    Select Case k
        Case 0: TenPhim = "Enter"
        Case 3: TenPhim = "F3 Back"
        Case 11: TenPhim = "Ctrl+S Save"
        Case 12: TenPhim = "F12 Cancel"
        Case 17: TenPhim = "Shift+F5"
        Case 25: TenPhim = "Ctrl+F1 Display <-> Change"
        Case 81: TenPhim = "Page Up"
        Case 82: TenPhim = "Page Down"
    End Select
End Function

Public Sub GL_Press(ByVal c As SapGL_Ctl)
    GL_Log "press", c, c.mText & IIf(c.Tooltip <> "", " (" & c.Tooltip & ")", "")
    If Left(c.mRef, 4) = "POP:" Then
        BamPopup c.mRef
        Exit Sub
    End If
    ChanKhiCoPopup c
    Select Case c.mRef
        Case "TB:ENTER": VongXuLy
        Case "TB:SAVE": BamLuu
        Case "TB:BACK": MoGiaoDich "SESSION_MANAGER": XayManHinh
        Case "TB:OTHERPO"
            MoPopup "CHON_PO", "Select Document", "", "CHON_PO"
            XayManHinh
        Case "TB:DOICHEDO": DoiCheDoME22N
        Case "TB:CHECK": BamCheck
        Case "PR:DETAIL": prDetail = Not prDetail: VongXuLy
        Case "PO:DETAIL"
            poDetail = Not poDetail
            If poPos > MaxCuon("PO") Then poPos = MaxCuon("PO")
            VongXuLy
        Case Else
            Err.Raise 438, NGUON, "Object doesn't support this property or method (press) (gia lap)"
    End Select
End Sub

Private Sub BamPopup(ByVal hanhDong As String)
    Dim viec As String
    If gPop = "" Then Err.Raise 619, NGUON, "The control could not be found by id. (khong co popup)"
    viec = gPopViec
    GL_LogDong "Popup dong", "wnd[1]", hanhDong
    Select Case hanhDong
        Case "POP:CHONPO"
            DongPopup
            XoaSbar
            MoPOXem gPopSel
        Case "POP:YES", "POP:OK"
            DongPopup
            XoaSbar
            If viec = "LUU" Then LuuChungTu True
        Case Else
            DongPopup
            XoaSbar
    End Select
    XayManHinh
End Sub

' ---------------------------------------------------------------- grid ALV ME51N
Private Function CotGrid() As Variant
    CotGrid = Array("MATNR", "TXZ01", "NAME1", "MENGE", "MEINS", "PREIS", "DISPO", "EKORG", "EKGRP", "KNTTP", "BNFPO")
End Function

Public Sub GL_ModifyCell(ByVal c As SapGL_Ctl, ByVal r As Long, ByVal col As String, ByVal v As String)
    Dim d As SapGL_Dong, x As Double, ok As Boolean
    ChanKhiCoPopup c
    GL_Log "modifyCell", c, "dong " & r & " " & col & " = " & v
    If r < 0 Or r > prItems.Count Then Err.Raise 623, NGUON, "The method got an invalid argument (row " & r & ") (gia lap)"
    If r = prItems.Count Then
        Set d = New SapGL_Dong
        prItems.Add d
    End If
    Set d = prItems.Item(r + 1)
    d.DaDoi = True
    Select Case UCase(col)
        Case "MATNR": d.Mat = UCase(Trim(v))
        Case "NAME1": d.Name1 = Trim(v)
        Case "MENGE"
            x = DocSo(v, ok)
            If ok Then
                d.Menge = x
            Else
                d.Loi = "Enter a numeric value (MENGE '" & v & "')"
            End If
        Case "PREIS"
            x = DocSo(v, ok)
            d.DaGoGia = True
            If ok Then
                d.Preis = Round(x, cGiaLe)
            Else
                d.Loi = "Enter a numeric value (PREIS '" & v & "')"
            End If
        Case "DISPO": d.Dispo = UCase(Trim(v))
        Case "EKORG": d.Ekorg = UCase(Trim(v))
        Case "EKGRP": d.Ekgrp = UCase(Trim(v))
        Case "KNTTP": d.Knttp = UCase(Trim(v))
        Case "TXZ01": d.Txz01 = v
        Case Else
            Err.Raise 623, NGUON, "The method got an invalid argument (column " & col & ") (gia lap)"
    End Select
End Sub

Public Function GL_GetCellValue(ByVal c As SapGL_Ctl, ByVal r As Long, ByVal col As String) As String
    Dim d As SapGL_Dong, s As String
    If r >= 0 And r < prItems.Count Then
        Set d = prItems.Item(r + 1)
        Select Case UCase(col)
            Case "MATNR": s = d.Mat
            Case "TXZ01": s = IIf(d.Txz01 <> "", d.Txz01, d.Mat)
            Case "NAME1": s = d.Name1
            Case "MENGE": s = DinhDangSo(d.Menge, 3)
            Case "MEINS": s = d.Meins
            Case "PREIS": s = DinhDangSo(d.Preis, cGiaLe)
            Case "DISPO": s = d.Dispo
            Case "EKORG": s = d.Ekorg
            Case "EKGRP": s = d.Ekgrp
            Case "KNTTP": s = d.Knttp
            Case "BNFPO": s = d.Bnfpo
        End Select
    ElseIf r < 0 Then
        Err.Raise 623, NGUON, "The method got an invalid argument (row " & r & ") (gia lap)"
    End If
    If cLogDoc Then GL_Log "getCellValue", c, "dong " & r & " " & col & " -> " & s
    GL_GetCellValue = s
End Function

Public Sub GL_GridEnter(ByVal c As SapGL_Ctl)
    ChanKhiCoPopup c
    GL_Log "pressEnter", c, ""
    VongXuLy
End Sub

' ---------------------------------------------------------------- table control
Private Function CotBang(ByVal sRef As String) As Variant
    Select Case sRef
        Case "PO": CotBang = Array("txt|MEPO1211-EBELP", "txt|MEPO1211-LOEKZ", "ctxt|MEPO1211-KNTTP", "ctxt|MEPO1211-EPSTP", _
                                   "ctxt|MEPO1211-EMATN", "txt|MEPO1211-TXZ01", "txt|MEPO1211-MENGE", "ctxt|MEPO1211-MEINS", _
                                   "ctxt|MEPO1211-EEIND", "txt|MEPO1211-NETPR", "ctxt|MEPO1211-WAERS", "txt|MEPO1211-PEINH", _
                                   "txt|MEPO1211-NAME1", "ctxt|MEPO1211-BANFN", "txt|MEPO1211-BNFPO")
        Case "COND": CotBang = Array("ctxt|KOMV-KSCHL", "txt|KOMV-VTEXT", "txt|KOMV-KBETR", "ctxt|RV61A-KOEIN", _
                                     "txt|KOMV-KPEIN", "ctxt|KOMV-KMEIN", "txt|KOMV-KWERT")
        Case "MIGO": CotBang = Array("txt|GOITEM-ZEILE", "txt|GOITEM-MAKTX", "chk|GOITEM-TAKE_IT", "txt|GOITEM-ERFMG", _
                                     "ctxt|GOITEM-ERFME", "txt|GOITEM-LGOBE", "txt|GOITEM-SGTXT", "txt|GOITEM-EBELN", "txt|GOITEM-EBELP")
        Case Else: CotBang = Array()
    End Select
End Function

Public Function GL_Columns(ByVal tbl As SapGL_Ctl) As Object
    Dim res As New SapGL_List, e As Variant, a() As String, x As SapGL_Ctl, i As Long
    For Each e In CotBang(tbl.mRef)
        a = Split(e, "|")
        Set x = TaoCtl("COL", tbl.mId & "/col[" & i & "]", a(1), "GuiTableColumn")
        res.Them x
        i = i + 1
    Next e
    Set GL_Columns = res
End Function

Private Function SoDongHien(ByVal sRef As String) As Long
    Select Case sRef
        Case "PO": SoDongHien = IIf(poDetail, cVisPOMo, cVisPOThu)
        Case "COND": SoDongHien = cVisCond
        Case "MIGO": SoDongHien = cVisMigo
        Case Else: SoDongHien = 10
    End Select
End Function

Private Function SoDong(ByVal sRef As String) As Long
    Select Case sRef
        Case "PO": SoDong = poItems.Count + SoDongHien("PO")          ' + cac dong trong de nhap them
        Case "COND": If poItems.Count > 0 Then SoDong = CurPO().CondCount() + SoDongHien("COND")
        Case "MIGO": SoDong = mgItems.Count
    End Select
End Function

Private Function MaxCuon(ByVal sRef As String) As Long
    MaxCuon = SoDong(sRef) - SoDongHien(sRef)
    If MaxCuon < 0 Then MaxCuon = 0
End Function

Private Function ViTriCuon(ByVal sRef As String) As Long
    Select Case sRef
        Case "PO": ViTriCuon = poPos
        Case "COND": ViTriCuon = poCondPos
        Case "MIGO": ViTriCuon = mgPos
    End Select
End Function

Public Function GL_RowCount(ByVal c As SapGL_Ctl) As Long
    If c.Kind = "GRID" Then
        GL_RowCount = prItems.Count
    Else
        GL_RowCount = SoDong(c.mRef)
    End If
End Function

Public Function GL_VisibleRows(ByVal c As SapGL_Ctl) As Long
    GL_VisibleRows = SoDongHien(c.mRef)
End Function

Public Function GL_Scrollbar(ByVal tbl As SapGL_Ctl) As Object
    Dim s As SapGL_Ctl
    Set s = TaoCtl("SCROLL", tbl.mId & "/verticalScrollbar", "verticalScrollbar", "GuiScrollbar")
    s.mRef = tbl.mRef
    Set s.Cha = tbl
    Set GL_Scrollbar = s
End Function

Public Function GL_ScrollPos(ByVal s As SapGL_Ctl) As Long
    GL_ScrollPos = ViTriCuon(s.mRef)
End Function

Public Function GL_ScrollMax(ByVal s As SapGL_Ctl) As Long
    GL_ScrollMax = MaxCuon(s.mRef)
End Function

Public Sub GL_SetScrollPos(ByVal s As SapGL_Ctl, ByVal v As Long)
    ChanKhiCoPopup s
    GL_Log "Position =", s, CStr(v) & " (toi da " & MaxCuon(s.mRef) & ")"
    If v < 0 Or v > MaxCuon(s.mRef) Then
        GL_LogDong "LOI thao tac", RutGonId(s), "cuon qua gioi han"
        Err.Raise 623, NGUON, "The method got an invalid argument. (Position " & v & ", Maximum " & MaxCuon(s.mRef) & ")"
    End If
    Select Case s.mRef
        Case "PO": poPos = v
        Case "COND": poCondPos = v
        Case "MIGO": mgPos = v
    End Select
    VongXuLy                                  ' cuon bang = SAP xu ly man hinh
End Sub

Private Sub CuonTrang(ByVal xuong As Boolean)
    Dim sRef As String, p As Long
    sRef = gFocusTbl
    If sRef = "" Then
        If gTx = "MIGO" Then
            sRef = "MIGO"
        Else
            sRef = "PO"
        End If
    End If
    p = ViTriCuon(sRef) + IIf(xuong, 1, -1) * SoDongHien(sRef)
    If p > MaxCuon(sRef) Then p = MaxCuon(sRef)
    If p < 0 Then p = 0
    Select Case sRef
        Case "PO": poPos = p
        Case "COND": poCondPos = p
        Case "MIGO": mgPos = p
    End Select
    VongXuLy
End Sub

Public Function GL_GetCell(ByVal tbl As SapGL_Ctl, ByVal r As Long, ByVal col As Long) As Object
    Dim cols As Variant, a() As String, c As SapGL_Ctl
    cols = CotBang(tbl.mRef)
    If r < 0 Or r >= SoDongHien(tbl.mRef) Or col < 0 Or col > UBound(cols) Then
        Err.Raise 623, NGUON, "The method got an invalid argument. (GetCell " & r & "," & col & ")"
    End If
    a = Split(cols(col), "|")
    Set c = TaoCtl("CELL", tbl.mId & "/" & a(0) & a(1) & "[" & col & "," & r & "]", a(1), "")
    Select Case a(0)
        Case "ctxt": c.mType = "GuiCTextField"
        Case "chk": c.mType = "GuiCheckBox"
        Case Else: c.mType = "GuiTextField"
    End Select
    c.mRef = tbl.mRef
    c.HangHien = r
    c.Hang = ViTriCuon(tbl.mRef) + r
    c.Cot = col
    Set c.Cha = tbl
    Set GL_GetCell = c
End Function

' Noi dung o bang (theo dong tuyet doi c.Hang)
Private Function TextO(ByVal c As SapGL_Ctl) As String
    Dim d As SapGL_Dong, i As Long
    Select Case c.mRef
        Case "PO"
            If c.Hang >= poItems.Count Then Exit Function
            Set d = poItems.Item(c.Hang + 1)
            Select Case c.mName
                Case "MEPO1211-EBELP": TextO = d.Ebelp
                Case "MEPO1211-KNTTP": If d.DaLay Then TextO = "K"
                Case "MEPO1211-EMATN": TextO = d.Mat
                Case "MEPO1211-TXZ01": TextO = IIf(d.Txz01 <> "", d.Txz01, d.Mat)
                Case "MEPO1211-MENGE": If d.DaLay Or d.Menge <> 0 Then TextO = DinhDangSo(d.Menge, 3)
                Case "MEPO1211-MEINS": If d.DaLay Then TextO = d.Meins
                Case "MEPO1211-NETPR": If d.DaLay Or d.Netpr <> 0 Then TextO = DinhDangSo(d.Netpr, cGiaLe)
                Case "MEPO1211-WAERS": If d.DaLay Then TextO = "VND"
                Case "MEPO1211-PEINH": If d.DaLay Then TextO = DinhDangSo(d.Peinh, 0)
                Case "MEPO1211-NAME1": If d.DaLay Then TextO = poVendor
                Case "MEPO1211-BANFN": TextO = d.Banfn
                Case "MEPO1211-BNFPO": TextO = d.BanfnItem
                Case "MEPO1211-EEIND": If d.DaLay Then TextO = Format(Date, "dd.mm.yyyy")
            End Select
        Case "COND"
            If poItems.Count = 0 Then Exit Function
            Set d = CurPO()
            i = c.Hang + 1
            If i > d.CondCount() Then Exit Function
            Select Case c.mName
                Case "KOMV-KSCHL": TextO = d.CondK(i)
                Case "KOMV-VTEXT"
                    Select Case d.CondK(i)
                        Case "PB00": TextO = "Gross Price"
                        Case "PBXX": TextO = "Gross Price (manual)"
                    End Select
                Case "KOMV-KBETR": TextO = DinhDangSo(d.CondA(i), cGiaLe)
                Case "RV61A-KOEIN": TextO = "VND"
                Case "KOMV-KPEIN": TextO = "1"
                Case "KOMV-KMEIN": TextO = d.Meins
                Case "KOMV-KWERT": TextO = DinhDangSo(d.CondA(i) * d.Menge, cGiaLe)
            End Select
        Case "MIGO"
            If c.Hang >= mgItems.Count Then Exit Function
            Set d = mgItems.Item(c.Hang + 1)
            Select Case c.mName
                Case "GOITEM-ZEILE": TextO = CStr(c.Hang + 1)
                Case "GOITEM-MAKTX": TextO = d.Mat
                Case "GOITEM-ERFMG": TextO = DinhDangSo(d.Menge, 3)
                Case "GOITEM-ERFME": TextO = d.Meins
                Case "GOITEM-SGTXT": TextO = d.Sgtxt
                Case "GOITEM-EBELN": TextO = d.SoPO
                Case "GOITEM-EBELP": TextO = d.PoItem
            End Select
    End Select
End Function

Private Function SuaDuocO(ByVal c As SapGL_Ctl) As Boolean
    Dim d As SapGL_Dong, i As Long
    Select Case c.mRef
        Case "PO"
            If poXem Then Exit Function
            If c.Hang >= poItems.Count Then
                Select Case c.mName
                    Case "MEPO1211-BANFN", "MEPO1211-BNFPO", "MEPO1211-EMATN", "MEPO1211-MENGE", "MEPO1211-NETPR": SuaDuocO = True
                End Select
                Exit Function
            End If
            Set d = poItems.Item(c.Hang + 1)
            Select Case c.mName
                Case "MEPO1211-BANFN", "MEPO1211-BNFPO": SuaDuocO = Not d.DaLay
                Case "MEPO1211-NETPR": SuaDuocO = Not d.KhoaGia
                Case "MEPO1211-MENGE", "MEPO1211-EMATN", "MEPO1211-PEINH", "MEPO1211-EEIND": SuaDuocO = True
            End Select
        Case "COND"
            If poItems.Count = 0 Or poXem Then Exit Function
            Set d = CurPO()
            i = c.Hang + 1
            Select Case c.mName
                Case "KOMV-KSCHL"
                    If i > d.CondCount() Then
                        SuaDuocO = True
                    Else
                        SuaDuocO = (d.CondK(i) <> "PB00")
                    End If
                Case "KOMV-KBETR"
                    If i <= d.CondCount() Then
                        If d.CondK(i) = "PB00" Then
                            SuaDuocO = Not d.KhoaGia
                        Else
                            SuaDuocO = True
                        End If
                    End If
            End Select
        Case "MIGO"
            If c.Hang >= mgItems.Count Then Exit Function
            Select Case c.mName
                Case "GOITEM-TAKE_IT", "GOITEM-SGTXT", "GOITEM-ERFMG": SuaDuocO = True
            End Select
    End Select
End Function

Private Sub SetO(ByVal c As SapGL_Ctl, ByVal v As String)
    Dim d As SapGL_Dong, x As Double, ok As Boolean, i As Long
    ChanKhongSuaDuoc c
    Select Case c.mRef
        Case "PO"
            Do While poItems.Count <= c.Hang
                Set d = New SapGL_Dong
                poItems.Add d
            Loop
            Set d = poItems.Item(c.Hang + 1)
            d.DaDoi = True
            If gTx = "ME22N" Then poDoi = True
            If c.mName <> "MEPO1211-EMATN" Then d.DaDoiGia = True
            Select Case c.mName
                Case "MEPO1211-BANFN": d.Banfn = Trim(v)
                Case "MEPO1211-BNFPO": d.BanfnItem = Trim(v)
                Case "MEPO1211-EMATN": d.Mat = UCase(Trim(v))
                Case "MEPO1211-MENGE"
                    x = DocSo(v, ok)
                    If ok Then
                        d.Menge = x
                    Else
                        d.Loi = "Enter a numeric value (Quantity '" & v & "')"
                    End If
                Case "MEPO1211-NETPR"
                    x = DocSo(v, ok)
                    If ok Then
                        d.Netpr = Round(x, cGiaLe)
                        DatGiaDieuKien d, "PB00", d.Netpr
                    Else
                        d.Loi = "Enter a numeric value (Net Price '" & v & "')"
                    End If
                Case "MEPO1211-PEINH"
                    x = DocSo(v, ok)
                    If ok And x > 0 Then d.Peinh = x
            End Select
        Case "COND"
            Set d = CurPO()
            d.DaDoi = True
            If gTx = "ME22N" Then poDoi = True
            If c.mName = "KOMV-KBETR" Then d.DaDoiGia = True
            i = c.Hang + 1
            Select Case c.mName
                Case "KOMV-KSCHL"
                    If i <= d.CondCount() Then
                        d.CondDatK i, UCase(Trim(v))
                    ElseIf Trim(v) <> "" Then
                        d.CondThem UCase(Trim(v)), 0
                    End If
                Case "KOMV-KBETR"
                    x = DocSo(v, ok)
                    If ok Then
                        d.CondDatA i, Round(x, cGiaLe)
                    Else
                        d.Loi = "Enter a numeric value (Amount '" & v & "')"
                    End If
            End Select
        Case "MIGO"
            Set d = mgItems.Item(c.Hang + 1)
            Select Case c.mName
                Case "GOITEM-SGTXT": d.Sgtxt = v
                Case "GOITEM-ERFMG"
                    x = DocSo(v, ok)
                    If ok Then d.Menge = x
            End Select
    End Select
End Sub

Private Sub DatGiaDieuKien(ByVal d As SapGL_Dong, ByVal k As String, ByVal a As Double)
    Dim i As Long
    For i = 1 To d.CondCount()
        If d.CondK(i) = k Then d.CondDatA i, a: Exit Sub
    Next i
    d.CondThem k, a
End Sub


' ================================================================================
'  5. XU LY CUA "SAP" (Enter, Save, mo giao dich)
' ================================================================================

Private Function CurPR() As SapGL_Dong
    If prItems Is Nothing Then Set prItems = New Collection
    If prItems.Count = 0 Then
        Set CurPR = New SapGL_Dong
    Else
        If prCur < 0 Or prCur >= prItems.Count Then prCur = 0
        Set CurPR = prItems.Item(prCur + 1)
    End If
End Function

Private Function CurPO() As SapGL_Dong
    If poItems.Count = 0 Then
        Set CurPO = New SapGL_Dong
    Else
        If poCur < 0 Or poCur >= poItems.Count Then poCur = 0
        Set CurPO = poItems.Item(poCur + 1)
    End If
End Function

Private Sub XoaSbar()
    gSbT = ""
    gSbM = ""
End Sub

Private Sub DatSbar(ByVal loai As String, ByVal nd As String)
    If loai = "" Then loai = "S"
    gSbM = loai
    gSbT = nd
    GL_LogDong "SAP bao [" & loai & "]", "wnd[0]/sbar", nd
End Sub

' Enter / doi combobox / doi tab / cuon bang: SAP xu ly lai man hinh
Private Sub VongXuLy()
    If gSbM = "W" Then gWAck = True
    XoaSbar
    If Trim(gOkcd) <> "" Then
        ChayLenh gOkcd
        gOkcd = ""
    Else
        Select Case gTx
            Case "ME51N": XuLyPR
            Case "ME21N", "ME22N": XuLyPO
            Case "MIGO": XuLyMIGO
        End Select
    End If
    XayManHinh
End Sub

Private Sub ChayLenh(ByVal lenh As String)
    Dim tc As String
    tc = UCase(Trim(lenh))
    If Left(tc, 2) = "/N" Or Left(tc, 2) = "/O" Then
        tc = Mid(tc, 3)
    ElseIf gTx <> "SESSION_MANAGER" Then
        DatSbar "E", "Function code cannot be selected"
        Exit Sub
    End If
    Select Case tc
        Case "": MoGiaoDich "SESSION_MANAGER"
        Case "ME51N", "ME21N", "ME22N", "ME23N", "MIGO": MoGiaoDich tc
        Case Else
            DatSbar "E", "Transaction " & tc & " does not exist (SAP gia lap chi co ME51N, ME21N, ME22N, ME23N, MIGO)"
    End Select
End Sub

Private Sub MoGiaoDich(ByVal tc As String)
    Dim loai As String, nd As String, gt As Variant
    DocCauHinh
    gTx = tc
    gWAck = False
    gFocusTbl = ""
    gPop = ""
    XoaSbar
    Select Case tc
        Case "ME51N"
            gTitle = "Create Purchase Requisition"
            prBsart = "NB": Set prItems = New Collection: prCur = 0: prTab = "TABREQDT1": prDetail = True
        Case "ME21N", "ME22N", "ME23N"
            gTitle = Choose(IIf(tc = "ME21N", 1, IIf(tc = "ME22N", 2, 3)), "Create Purchase Order", "Change Purchase Order", "Display Purchase Order")
            poBsart = "NB": poVendor = "": poVendorOK = "": Set poItems = New Collection
            poCur = 0: poTab = "TABIDT1": poPos = 0: poCondPos = 0: poSoPO = "": poXem = (tc = "ME23N")
            poKhoa = "": poDoi = False
            If Not poDetailInit Then poDetail = True: poDetailInit = True   ' SAP nho trang thai Item Detail theo user
        Case "MIGO"
            gTitle = "Goods Receipt Purchase Order - " & cUser
            mgAction = "A01": mgRefdoc = "R01": mgPO = "": mgPOLoaded = "": Set mgItems = New Collection: mgPos = 0
            mgBldat = "": mgBudat = "": mgLfsnr = ""
        Case Else
            gTitle = "SAP Easy Access"
    End Select
    If prItems Is Nothing Then Set prItems = New Collection
    If poItems Is Nothing Then Set poItems = New Collection
    If mgItems Is Nothing Then Set mgItems = New Collection
    GL_LogDong "Mo giao dich", tc, gTitle
    If tc <> "SESSION_MANAGER" Then
        If KB("MO_GD", tc, loai, nd, gt) Then
            Select Case loai
                Case "POPUP_TB": MoPopup "TB", "Information", nd, ""
                Case "POPUP_HOI", "POPUP_LOI": MoPopup "HOI", "Question", nd, ""
                Case Else: DatSbar loai, nd
            End Select
        End If
    End If
End Sub

' ---- ME51N: Enter ----
Private Sub XuLyPR()
    ' SAP chi bao thong bao cho dong VUA SUA (dong loi cu van bi danh dau, chan o buoc Save)
    Dim i As Long, d As SapGL_Dong, loai As String, nd As String, gt As Variant, eMsg As String, wMsg As String, wType As String
    For i = 1 To prItems.Count
        Set d = prItems.Item(i)
        If d.Mat <> "" And d.Bnfpo = "" Then d.Bnfpo = CStr(SoItemPRMoi())
        If d.DaDoi Then
            If d.Mat <> "" Then
                If KB("NHAP_DONG", d.Mat, loai, nd, gt) Then
                    If loai = "E" Or loai = "A" Then
                        d.Loi = nd & " (" & d.Mat & ")"
                    Else
                        wType = loai
                        wMsg = nd & " (" & d.Mat & ")"
                    End If
                End If
            End If
            If d.DaGoGia And d.Preis <= 0 And d.Loi = "" Then d.Loi = "Net price must be greater than 0"
            If d.DaGoGia And d.Preis > 0 And d.Loi = "Net price must be greater than 0" Then d.Loi = ""
            If Left(d.Loi, 22) = "Enter a numeric value " Then
                If eMsg = "" Then eMsg = d.Loi
                d.Loi = ""
            ElseIf d.Loi <> "" And eMsg = "" Then
                eMsg = d.Loi
            End If
            d.DaDoi = False
        End If
    Next i
    If eMsg <> "" Then
        DatSbar "E", eMsg
    ElseIf wMsg <> "" Then
        DatSbar wType, wMsg
    End If
End Sub

Private Function SoItemPRMoi() As Long
    Dim i As Long, mx As Long
    For i = 1 To prItems.Count
        If Val(prItems.Item(i).Bnfpo) > mx Then mx = Val(prItems.Item(i).Bnfpo)
    Next i
    SoItemPRMoi = mx + 10
End Function

' ---- ME21N: Enter ----
Private Sub XuLyPO()
    Dim i As Long, d As SapGL_Dong, r As Long, loai As String, nd As String, gt As Variant
    Dim eMsg As String, wMsg As String, wType As String, ws As Worksheet, loiTruoc As String
    If poXem Then Exit Sub
    If gTx = "ME22N" And poSoPO = "" Then Exit Sub
    Set ws = SheetDB()

    ' nha cung cap
    If poVendor <> "" And poVendor <> poVendorOK Then
        If KB("NCC", poVendor, loai, nd, gt) Then
            If loai = "E" Or loai = "A" Then
                eMsg = nd & " (" & poVendor & ")"
            Else
                wType = loai
                wMsg = nd
            End If
        End If
        If eMsg = "" Then poVendorOK = poVendor
    End If

    For i = 1 To poItems.Count
        Set d = poItems.Item(i)
        loiTruoc = d.Loi
        ' lay dong PR
        If Not d.DaLay And d.Banfn <> "" Then
            r = DbTimPR(d.Banfn, d.BanfnItem)
            If d.BanfnItem = "" Then
                d.Loi = "Enter the requisition item for requisition " & d.Banfn
            ElseIf r = 0 Then
                d.Loi = "Purchase requisition " & d.Banfn & " " & Format(Val(d.BanfnItem), "00000") & " does not exist"
            ElseIf DbPRDaDatPO(d.Banfn, d.BanfnItem) <> "" Then
                d.Loi = "Requisition " & d.Banfn & " " & Format(Val(d.BanfnItem), "00000") & " already ordered (PO " & DbPRDaDatPO(d.Banfn, d.BanfnItem) & ")"
            Else
                d.Loi = ""
                d.DaLay = True
                d.Mat = CStr(ws.Cells(r, 4).Value)
                d.Txz01 = d.Mat
                d.SLPR = CDbl(ws.Cells(r, 5).Value)
                d.GiaPR = CDbl(ws.Cells(r, 6).Value)
                d.Menge = d.SLPR
                d.Netpr = d.GiaPR
                d.Ebelp = CStr(SoItemPOMoi())
                d.BanfnItem = CStr(CLng(Val(d.BanfnItem)))
                d.CondThem "PB00", d.Netpr
                d.DaDoiGia = True
                If KB("KHOA_GIA", d.Mat, loai, nd, gt) Then d.KhoaGia = True
                If KB("LAY_DONG_PR", d.Mat, loai, nd, gt) Then
                    If Right(Trim(CStr(gt)), 1) = "%" Then
                        d.Netpr = Round(d.GiaPR * Val(Trim(CStr(gt))) / 100, cGiaLe)
                        DatGiaDieuKien d, "PB00", d.Netpr
                    ElseIf IsNumeric(gt) And Trim(CStr(gt)) <> "" Then
                        d.Netpr = CDbl(gt)
                        DatGiaDieuKien d, "PB00", d.Netpr
                    End If
                    If loai = "E" Or loai = "A" Then
                        d.Loi = nd
                    ElseIf nd <> "" Then
                        wType = loai: wMsg = nd
                    End If
                End If
                GL_LogDong "Lay dong PR", d.Banfn & "/" & d.BanfnItem, d.Mat & " SL " & d.Menge & " x gia " & d.Netpr & IIf(d.KhoaGia, " (o Net Price KHOA)", "")
            End If
        End If

        ' ME22N: giam SL duoi so da nhap kho -> SAP canh bao (W)
        If gTx = "ME22N" And d.DaLay And d.DaDoi And d.SLNhap > 0 And d.Menge < d.SLNhap - 0.0005 Then
            If wMsg = "" Then wType = "W": wMsg = "Quantity ordered is less than quantity delivered (item " & d.Ebelp & ": " & DinhDangSo(d.SLNhap, 3) & ")"
        End If
        If d.DaLay And d.DaDoi Then
            XuLyDieuKien d                    ' tab Conditions (PBXX thay PB00)
            If d.DaDoiGia Then
                If d.Netpr <= d.GiaPR + 0.5 And d.Loi = "PO item value must be less than or equal to PR item value" Then d.Loi = ""
                If Left(d.Loi, 22) = "Enter a numeric value " Then
                    If eMsg = "" Then eMsg = d.Loi
                    d.Loi = ""
                ElseIf d.Menge <= 0 Then
                    d.Loi = "Enter a PO quantity greater than 0 (item " & d.Ebelp & ")"
                ElseIf d.Netpr <= 0 Then
                    d.Loi = "Net price must be greater than 0 (item " & d.Ebelp & ")"
                ElseIf InStr(d.Loi, "greater than 0") > 0 Then
                    d.Loi = ""
                End If
            End If
        End If
        ' SAP chi bao thong bao cua dong VUA SUA; dong loi cu van bi danh dau va chan o buoc Save
        ' (vd chi go them dong dieu kien thi khong bao lai loi gia cu)
        If d.DaDoi And d.Loi <> "" And eMsg = "" And (d.Loi <> loiTruoc Or d.DaDoiGia) Then eMsg = d.Loi
        d.DaDoi = False
        d.DaDoiGia = False
    Next i
    If eMsg <> "" Then
        DatSbar "E", eMsg
    ElseIf wMsg <> "" Then
        DatSbar wType, wMsg
    End If
End Sub

' Dieu kien gia (tab Conditions): PBXX (> 0) thay cho PB00
Private Sub XuLyDieuKien(ByVal d As SapGL_Dong)
    Dim i As Long, loai As String, nd As String, gt As Variant, gia As Double, coPBXX As Boolean
    i = 1
    Do While i <= d.CondCount()
        If Trim(d.CondK(i)) = "" Then
            d.CondXoa i
        Else
            i = i + 1
        End If
    Loop
    For i = 1 To d.CondCount()
        Select Case d.CondK(i)
            Case "PB00", "PR00", "P000"
                If Not coPBXX Then gia = d.CondA(i)
            Case "PBXX"
                If KB("PBXX", d.Mat, loai, nd, gt) Then
                    d.Loi = nd
                    Exit Sub
                End If
                If d.CondA(i) > 0 Then gia = d.CondA(i): coPBXX = True
            Case Else
                d.Loi = "Condition type " & d.CondK(i) & " is not defined"
                Exit Sub
        End Select
    Next i
    If Left(d.Loi, 15) = "Condition type " Then d.Loi = ""
    If gia > 0 Then d.Netpr = gia
End Sub

Private Function SoItemPOMoi() As Long
    Dim i As Long, mx As Long
    For i = 1 To poItems.Count
        If Val(poItems.Item(i).Ebelp) > mx Then mx = Val(poItems.Item(i).Ebelp)
    Next i
    SoItemPOMoi = mx + 10
End Function

' ---- ME23N: mo PO de xem ----
Private Sub MoPOXem(ByVal soPO As String)
    Dim ws As Worksheet, r As Long, lastR As Long, d As SapGL_Dong
    Dim loai As String, nd As String, gt As Variant
    Set ws = SheetDB()
    Set poItems = New Collection
    poCur = 0: poPos = 0: poCondPos = 0: poSoPO = ""
    poKhoa = "": poDoi = False
    If gTx = "ME22N" Then poXem = False
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastR
        If CStr(ws.Cells(r, 1).Value) = "PO" And Trim(CStr(ws.Cells(r, 2).Value)) = Trim(soPO) Then
            Set d = New SapGL_Dong
            d.DaLay = True
            d.Ebelp = Trim(CStr(ws.Cells(r, 3).Value))
            d.Mat = Trim(CStr(ws.Cells(r, 4).Value))
            d.Txz01 = d.Mat
            d.Menge = SoO(ws.Cells(r, 5).Value)
            d.Netpr = SoO(ws.Cells(r, 6).Value)
            d.Banfn = Trim(CStr(ws.Cells(r, 8).Value))
            d.BanfnItem = Trim(CStr(ws.Cells(r, 9).Value))
            d.Loekz = UCase(Trim(CStr(ws.Cells(r, 11).Value)))
            d.SLNhap = SoO(ws.Cells(r, 10).Value)
            d.GiaPR = d.Netpr
            d.CondThem "PB00", d.Netpr
            poVendor = Trim(CStr(ws.Cells(r, 7).Value))
            poVendorOK = poVendor
            poBsart = Trim(CStr(ws.Cells(r, 14).Value))
            poItems.Add d
        End If
    Next r
    If poItems.Count = 0 Then
        DatSbar "E", "Document " & soPO & " does not exist"
        gTitle = "Display Purchase Order"
    Else
        poSoPO = soPO
        If gTx = "ME22N" Then
            If KB("MO_PO", soPO, loai, nd, gt) Then
                Select Case loai
                    Case "XEM": poXem = True
                    Case "KHOA"
                        poXem = True
                        poKhoa = IIf(nd <> "", nd, "Purchase order " & soPO & " is currently being processed by another user")
                        DatSbar "E", poKhoa
                    Case Else: DatSbar loai, nd
                End Select
            End If
        End If
        TieuDePO
        GL_LogDong IIf(gTx = "ME22N", "Mo PO (" & IIf(poXem, "hien thi", "sua") & ")", "Mo PO (xem)"), soPO, poItems.Count & " dong"
    End If
End Sub

Private Sub TieuDePO()
    gTitle = IIf(poBsart = "ZEXP", "Expense PO", "Standard PO") & " " & poSoPO & " Created by " & cUser
    ' Gia lap: ME22N o che do hien thi them chu "Display" vao tieu de (SuaPO_ME22N dua vao day de bam Ctrl+F1).
    ' CHUA kiem chung tieu de nay tren SAP that.
    If gTx = "ME22N" And poXem Then gTitle = "Display " & gTitle
End Sub

' ---- ME22N: Ctrl+F1 doi che do hien thi <-> sua ----
Private Sub DoiCheDoME22N()
    XoaSbar
    If poSoPO = "" Then
        DatSbar "E", "Function code cannot be selected"
    ElseIf poKhoa <> "" Then
        DatSbar "E", poKhoa
    Else
        poXem = Not poXem
        TieuDePO
        GL_LogDong "Doi che do", poSoPO, IIf(poXem, "hien thi", "sua")
    End If
    XayManHinh
End Sub

' ---- MIGO: Enter ----
Private Sub XuLyMIGO()
    Dim ws As Worksheet, r As Long, lastR As Long, d As SapGL_Dong, conLai As Double, coPO As Boolean
    Dim loai As String, nd As String, gt As Variant
    If mgPO = "" Then Exit Sub
    If mgPO = mgPOLoaded Then
        nd = LoiMIGO(False)
        If nd <> "" Then DatSbar "E", nd
        Exit Sub
    End If
    Set ws = SheetDB()
    Set mgItems = New Collection
    mgPos = 0
    mgPOLoaded = mgPO
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastR
        If CStr(ws.Cells(r, 1).Value) = "PO" And Trim(CStr(ws.Cells(r, 2).Value)) = mgPO Then
            coPO = True
            conLai = SoO(ws.Cells(r, 5).Value) - SoO(ws.Cells(r, 10).Value)
            If Trim(CStr(ws.Cells(r, 11).Value)) = "" And conLai > 0.0005 Then
                Set d = New SapGL_Dong
                d.SoPO = mgPO
                d.PoItem = Trim(CStr(ws.Cells(r, 3).Value))
                d.Mat = Trim(CStr(ws.Cells(r, 4).Value))
                d.Menge = conLai
                d.SLMo = conLai
                mgItems.Add d
            End If
        End If
    Next r
    If Not coPO Then
        DatSbar "E", "Document " & mgPO & " does not exist"
        Exit Sub
    End If
    If KB("MIGO_PO", mgPO, loai, nd, gt) Then
        If loai = "HET_DONG" Then
            Set mgItems = New Collection
            DatSbar "E", IIf(nd <> "", nd, "PO " & mgPO & " does not contain any selectable items")
            Exit Sub
        End If
        DatSbar loai, nd
        Exit Sub
    End If
    If mgItems.Count = 0 Then
        DatSbar "E", "PO " & mgPO & " does not contain any selectable items"
    End If
    If mgBldat = "" Then mgBldat = Format(Date, "dd.mm.yyyy")
    If mgBudat = "" Then mgBudat = Format(Date, "dd.mm.yyyy")
    GL_LogDong "Mo PO trong MIGO", mgPO, mgItems.Count & " dong con nhap duoc"
End Sub

' Loi MIGO hien tai ("" = khong loi). canTich = True: phai co it nhat 1 dong tich OK (Check / Post)
Private Function LoiMIGO(ByVal canTich As Boolean) As String
    Dim i As Long, d As SapGL_Dong, n As Long
    If Not NgayHopLe(mgBldat) Then LoiMIGO = "Document date " & mgBldat & " is not valid (format dd.mm.yyyy)": Exit Function
    If Not NgayHopLe(mgBudat) Then LoiMIGO = "Posting date " & mgBudat & " is not valid (format dd.mm.yyyy)": Exit Function
    For i = 1 To mgItems.Count
        Set d = mgItems.Item(i)
        If d.Take Then
            n = n + 1
            If d.Menge <= 0 Then LoiMIGO = "Enter a quantity (line " & i & ", PO item " & d.PoItem & ")": Exit Function
            If d.Menge > d.SLMo + 0.0005 Then
                LoiMIGO = "PO quantity exceeded by " & DinhDangSo(d.Menge - d.SLMo, 3) & " " & d.Meins & " : " & d.Mat & " " & d.SoPO & " " & d.PoItem
                Exit Function
            End If
        End If
    Next i
    If canTich And n = 0 Then LoiMIGO = "Select at least one item (column OK / Item OK)"
End Function

Private Function NgayHopLe(ByVal s As String) As Boolean
    Dim a() As String
    s = Trim(s)
    If Not (s Like "##.##.####") Then Exit Function
    a = Split(s, ".")
    On Error Resume Next
    NgayHopLe = (Day(DateSerial(CLng(a(2)), CLng(a(1)), CLng(a(0)))) = CLng(a(0))) And (CLng(a(1)) >= 1 And CLng(a(1)) <= 12)
    On Error GoTo 0
End Function

' ---- MIGO: nut Check ----
Private Sub BamCheck()
    Dim loi As String
    XoaSbar
    If gTx <> "MIGO" Then
        DatSbar "E", "Function code cannot be selected"
    Else
        XuLyMIGO
        If gSbM <> "E" Then
            loi = KiemTraMIGO()
            If loi <> "" Then
                DatSbar "E", loi
            Else
                DatSbar "S", "Document is O.K."
            End If
        End If
    End If
    XayManHinh
End Sub

' ---- Save / Post ----
Private Sub BamLuu()
    If gWAck And gSbM = "" Then
        ' giu trang thai da xac nhan canh bao
    End If
    Select Case gTx
        Case "ME51N", "ME21N", "MIGO": LuuChungTu False
        Case "ME22N"
            If poSoPO = "" Then
                DatSbar "E", "Enter a purchase order number (Other Purchase Order)"
            ElseIf poXem Or Not poDoi Then
                DatSbar "S", "No data was changed"
            Else
                LuuChungTu False
            End If
        Case "ME23N": DatSbar "S", "No data was changed"
        Case Else: DatSbar "E", "Function code cannot be selected"
    End Select
    XayManHinh
End Sub

Private Sub LuuChungTu(ByVal boQuaKichBan As Boolean)
    Dim loi As String, loai As String, nd As String, gt As Variant
    XoaSbar
    Select Case gTx
        Case "ME51N": XuLyPR: loi = KiemTraPR()
        Case "ME21N", "ME22N": XuLyPO: loi = KiemTraPO()
        Case "MIGO": loi = KiemTraMIGO()
    End Select
    If loi <> "" Then
        DatSbar "E", loi
        Exit Sub
    End If
    If Not boQuaKichBan Then
        If KB("LUU", gTx, loai, nd, gt) Then
            Select Case loai
                Case "W", "I"
                    If Not gWAck Then
                        DatSbar loai, nd
                        Exit Sub
                    End If
                Case "E", "A"
                    DatSbar loai, nd
                    Exit Sub
                Case "POPUP_HOI"
                    MoPopup "HOI", "Save Document", nd, "LUU"
                    Exit Sub
                Case "POPUP_LOI"
                    MoPopup "LOI", "Save Document", nd, "LUU"
                    Exit Sub
                Case "POPUP_TB"
                    MoPopup "TB", "Information", nd, "LUU"
                    Exit Sub
            End Select
        End If
    End If
    gWAck = False
    Select Case gTx
        Case "ME51N": LuuPR
        Case "ME21N": LuuPO
        Case "ME22N": LuuPOSua
        Case "MIGO": LuuMIGO
    End Select
End Sub

Private Function KiemTraPR() As String
    Dim i As Long, d As SapGL_Dong
    If prItems.Count = 0 Then KiemTraPR = "Document contains no items": Exit Function
    For i = 1 To prItems.Count
        Set d = prItems.Item(i)
        If d.Loi <> "" Then KiemTraPR = d.Loi & " (item " & d.Bnfpo & ")": Exit Function
        If d.Mat = "" Then KiemTraPR = "Enter material number (line " & i & ")": Exit Function
        If d.Menge <= 0 Then KiemTraPR = "Enter a quantity (item " & d.Bnfpo & ")": Exit Function
        If d.Preis <= 0 Then KiemTraPR = "Net price must be greater than 0 (item " & d.Bnfpo & ")": Exit Function
        If d.Knttp = "K" And d.Kostl = "" Then KiemTraPR = "Enter a cost center (item " & d.Bnfpo & ")": Exit Function
        If d.Knttp = "K" And d.Geber = "" Then KiemTraPR = "Enter a fund (item " & d.Bnfpo & ")": Exit Function
        If prBsart = "ZEX" And d.Knttp = "" Then KiemTraPR = "Account assignment category is required for document type ZEX (item " & d.Bnfpo & ")": Exit Function
    Next i
End Function

Private Function KiemTraPO() As String
    Dim i As Long, d As SapGL_Dong, n As Long
    If poVendor = "" Then KiemTraPO = "Enter vendor": Exit Function
    For i = 1 To poItems.Count
        Set d = poItems.Item(i)
        If d.Loi <> "" Then KiemTraPO = d.Loi: Exit Function
        If d.DaLay Then n = n + 1
    Next i
    If n = 0 Then KiemTraPO = "Document contains no items"
End Function

Private Function KiemTraMIGO() As String
    If mgPOLoaded = "" Or mgItems.Count = 0 Then KiemTraMIGO = "No items transferred - enter a purchase order and press Enter": Exit Function
    KiemTraMIGO = LoiMIGO(True)
End Function

Private Sub LuuPR()
    Dim so As String, i As Long, d As SapGL_Dong
    so = DbSoMoi("PR", 10)
    For i = 1 To prItems.Count
        Set d = prItems.Item(i)
        DbThem Array("PR", so, d.Bnfpo, d.Mat, d.Menge, d.Preis, d.Name1, "", "", 0, "", d.Kostl, d.Geber, prBsart, Now, _
                     "DISPO " & d.Dispo & " / EKORG " & d.Ekorg & " / EKGRP " & d.Ekgrp)
    Next i
    DbDongBo Array("PR|" & so)
    MoGiaoDich "ME51N"                         ' SAP mo lai man hinh tao PR moi
    DatSbar "S", "Purchase requisition number " & so & " created"
End Sub

Private Sub LuuPO()
    Dim so As String, i As Long, d As SapGL_Dong, ten As String
    so = DbSoMoi("PO", 11)
    For i = 1 To poItems.Count
        Set d = poItems.Item(i)
        If d.DaLay Then
            DbThem Array("PO", so, d.Ebelp, d.Mat, d.Menge, d.Netpr, poVendor, d.Banfn, d.BanfnItem, 0, d.Loekz, "", "", poBsart, Now, "")
        End If
    Next i
    DbDongBo Array("PO|" & so)
    ten = IIf(poBsart = "ZEXP", "Expense PO", "Standard PO")
    MoGiaoDich "ME21N"
    DatSbar "S", ten & " created under the number " & so
End Sub

' ---- ME22N: luu thay doi PO (sua dong co san, them dong moi) ----
Private Sub LuuPOSua()
    Dim so As String, i As Long, d As SapGL_Dong, ws As Worksheet, r As Long, lastR As Long, ten As String, found As Boolean
    so = poSoPO
    Set ws = SheetDB()
    For i = 1 To poItems.Count
        Set d = poItems.Item(i)
        If d.DaLay Then
            found = False
            lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            For r = 2 To lastR
                If CStr(ws.Cells(r, 1).Value) = "PO" And Trim(CStr(ws.Cells(r, 2).Value)) = so And Trim(CStr(ws.Cells(r, 3).Value)) = d.Ebelp Then
                    If SoO(ws.Cells(r, 5).Value) <> d.Menge Or SoO(ws.Cells(r, 6).Value) <> d.Netpr Then
                        ws.Cells(r, 5).Value = d.Menge
                        ws.Cells(r, 6).Value = d.Netpr
                        ws.Cells(r, 16).Value = "ME22N " & Format(Now, "dd/mm/yyyy hh:mm")
                    End If
                    found = True
                    Exit For
                End If
            Next r
            If Not found Then
                DbThem Array("PO", so, d.Ebelp, d.Mat, d.Menge, d.Netpr, poVendor, d.Banfn, d.BanfnItem, 0, d.Loekz, "", "", poBsart, Now, "ME22N them dong")
            End If
        End If
    Next i
    DbDongBo Array("PO|" & so)
    ten = IIf(poBsart = "ZEXP", "Expense PO", "Standard PO")
    MoGiaoDich "ME22N"
    MoPOXem so
    DatSbar "S", ten & " " & so & " changed"
End Sub

Private Sub LuuMIGO()
    Dim so As String, i As Long, d As SapGL_Dong, ws As Worksheet, r As Long, lastR As Long, khoa As String
    so = DbSoMoi("MIGO", 12)
    Set ws = SheetDB()
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 1 To mgItems.Count
        Set d = mgItems.Item(i)
        If d.Take Then
            For r = 2 To lastR
                If CStr(ws.Cells(r, 1).Value) = "PO" And Trim(CStr(ws.Cells(r, 2).Value)) = d.SoPO _
                   And Trim(CStr(ws.Cells(r, 3).Value)) = d.PoItem Then
                    ws.Cells(r, 10).Value = SoO(ws.Cells(r, 10).Value) + d.Menge
                    Exit For
                End If
            Next r
            DbThem Array("MIGO", so, CStr(i), d.Mat, d.Menge, "", "", d.SoPO, d.PoItem, "", "", "", "", MIGO_BWART_GL(), Now, _
                         "PO " & d.SoPO & "/" & d.PoItem & " | Ngay CT " & mgBldat & " | Ngay HT " & mgBudat & _
                         IIf(mgLfsnr <> "", " | Phieu " & mgLfsnr, "") & " | Text: " & d.Sgtxt)
            If InStr(khoa & ";", ";PO|" & d.SoPO & ";") = 0 Then khoa = khoa & ";PO|" & d.SoPO
        End If
    Next i
    DbDongBo Split("MIGO|" & so & khoa, ";")
    MoGiaoDich "MIGO"
    DatSbar "S", "Material document " & so & " posted"
End Sub

Private Function MIGO_BWART_GL() As String
    MIGO_BWART_GL = "101"
End Function


' ================================================================================
'  6. DU LIEU (sheet SAP_GiaLap_DB) + SO
' ================================================================================

Private Function DbTimPR(ByVal soPR As String, ByVal item As String) As Long
    Dim ws As Worksheet, r As Long, lastR As Long
    If Trim(item) = "" Or Not IsNumeric(Trim(item)) Then Exit Function
    Set ws = SheetDB()
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastR
        If CStr(ws.Cells(r, 1).Value) = "PR" And Trim(CStr(ws.Cells(r, 2).Value)) = Trim(soPR) Then
            If Val(CStr(ws.Cells(r, 3).Value)) = Val(item) Then DbTimPR = r: Exit Function
        End If
    Next r
End Function

Private Function DbPRDaDatPO(ByVal soPR As String, ByVal item As String) As String
    Dim ws As Worksheet, r As Long, lastR As Long
    Set ws = SheetDB()
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastR
        If CStr(ws.Cells(r, 1).Value) = "PO" And Trim(CStr(ws.Cells(r, 8).Value)) = Trim(soPR) _
           And Val(CStr(ws.Cells(r, 9).Value)) = Val(item) And Trim(CStr(ws.Cells(r, 11).Value)) = "" Then
            DbPRDaDatPO = Trim(CStr(ws.Cells(r, 2).Value))
            Exit Function
        End If
    Next r
End Function

' So chung tu moi. Moi cua so SAP (gPhien) co 1 dai so rieng (+1.000.000 moi cua so) de tram chay
' song song nhieu Excel khong bi trung so. Cua so dau tien: dai so cu nhu v1.
Private Function DbSoMoi(ByVal loai As String, ByVal oCfg As Long) As String
    Dim ws As Worksheet, r As Long, lastR As Long, mx As Double, v As Double, lo As Double, hi As Double
    Set ws = SheetDB()
    lo = Val(CStr(SheetCfg().Range("B" & oCfg).Value))
    If lo = 0 Then lo = 1
    lo = lo + CDbl(gPhien) * 1000000#
    hi = lo + 1000000#
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastR
        If CStr(ws.Cells(r, 1).Value) = loai Then
            v = Val(CStr(ws.Cells(r, 2).Value))
            If v >= lo And v < hi And v > mx Then mx = v
        End If
    Next r
    If mx = 0 Then
        mx = lo
    Else
        mx = mx + 1
    End If
    DbSoMoi = Format(mx, "0")
End Function

' ================================================================================
'  6b. DB DUNG CHUNG (o B15) - tram mo Excel moi cho moi lenh va khong luu file,
'      nen chung tu phai ghi ra file ngoai thi lenh sau moi thay.
'      Sheet SAP_GiaLap_DB trong file nay la ban sao: tai ve khi macro ket noi SAP,
'      ghi nguoc ra file sau moi lan Save / Post.
' ================================================================================

Private Function DbDuongDan() As String
    Dim p As String, sep As String
    p = cDbFile
    If p = "" Then Exit Function
    sep = Application.PathSeparator
    If InStr(p, ":") = 0 And Left(p, 1) <> "\" And Left(p, 1) <> "/" Then p = ThisWorkbook.Path & sep & p
    DbDuongDan = p
End Function

Private Function FileCo(ByVal p As String) As Boolean
    On Error Resume Next
    FileCo = (Dir(p) <> "")
End Function

' Mo file DB ngoai de GHI (cho toi 30 giay neu Excel khac dang ghi). Chua co file thi tao moi.
Private Function DbMoNgoai() As Workbook
    Dim p As String, wb As Workbook, t As Single, x As Workbook, wbMe As Workbook
    p = DbDuongDan()
    If p = "" Then Exit Function
    Set wbMe = ThisWorkbook
    Set x = WbDangMo(p)
    If Not x Is Nothing Then
        If Not ChiDoc(x) Then Set DbMoNgoai = x: Exit Function
        x.Close False
    End If
    If Not FileCo(p) Then
        Dim hdr As Variant
        hdr = SheetDB().Range("A1:P1").Value
        Set wb = Application.Workbooks.Add
        KichHoat wbMe
        wb.Worksheets(1).Range("A1:P1").Value = hdr
        wb.Worksheets(1).Name = "DB"
        LuuFileDB wb, p
        KichHoat wbMe
        GL_LogDong "DB dung chung", p, "tao file moi"
        Set DbMoNgoai = wb
        Exit Function
    End If
    t = Timer
    Do
        Set wb = Nothing
        On Error Resume Next
        Set wb = MoFileDB(p)
        On Error GoTo 0
        KichHoat wbMe                        ' file vua mo tu thanh file dang chon - tra lai file macro
        If Not wb Is Nothing Then
            If Not ChiDoc(wb) Then Set DbMoNgoai = wb: Exit Function
            wb.Close False
            KichHoat wbMe
        End If
        If Timer - t > 30 Or Timer < t Then Exit Do
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop
    GL_LogDong "LOI DB dung chung", p, "khong mo duoc de ghi (file dang bi khoa?)"
    Err.Raise 624, NGUON, "Khong ghi duoc file DB gia lap " & p & " (dang bi khoa?)"
End Function

' File DB dang mo trong Excel nay (theo duong dan, hoac theo ten khi FullName khong co thu muc)
Private Function WbDangMo(ByVal p As String) As Workbook
    Dim x As Workbook, ten As String
    ten = LCase(Mid(p, InStrRev(Replace(p, "\", "/"), "/") + 1))
    For Each x In Application.Workbooks
        If LCase(x.FullName) = LCase(p) Or (LCase(x.FullName) = ten And LCase(x.Name) = ten) Then Set WbDangMo = x: Exit Function
    Next x
End Function

' Mo file DB: ghi duoc thi mo de ghi; Excel khac dang ghi thi (Notify:=False) mo chi doc, khong hien hop thoai
Private Function MoFileDB(ByVal p As String) As Workbook
    Set MoFileDB = Application.Workbooks.Open(Filename:=p, UpdateLinks:=0, ReadOnly:=False, Notify:=False)
End Function

Private Function ChiDoc(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    ChiDoc = wb.ReadOnly                     ' LibreOffice khong co ReadOnly -> coi nhu ghi duoc
    Err.Clear
End Function

Private Sub KichHoat(ByVal wb As Workbook)
    On Error Resume Next
    wb.Activate
    Err.Clear
End Sub

Private Sub DbDongNgoai(ByVal wb As Workbook)
    Dim wbMe As Workbook
    Set wbMe = ThisWorkbook
    LuuFileDB wb, DbDuongDan()
    wb.Close False
    KichHoat wbMe
End Sub

' Luu file DB bang SaveAs de len chinh no (LibreOffice khong chiu .Save voi file Excel)
Private Sub LuuFileDB(ByVal wb As Workbook, ByVal p As String)
    Application.DisplayAlerts = False
    Select Case LCase(Mid(p, InStrRev(p, ".") + 1))
        Case "xls": wb.SaveAs p, 56
        Case "xlsm": wb.SaveAs p, 52
        Case Else: wb.SaveAs p, 51
    End Select
    Application.DisplayAlerts = True
End Sub

' Tai file DB ngoai ve sheet SAP_GiaLap_DB (chi doc)
Private Sub DbTaiVe()
    Dim p As String, wb As Workbook, ws As Worksheet, src As Worksheet, n As Long, daMo As Boolean, x As Workbook, wbMe As Workbook
    p = DbDuongDan()
    If p = "" Then Exit Sub
    If Not FileCo(p) Then Exit Sub
    Set wbMe = ThisWorkbook
    Set ws = SheetDB()
    Set wb = WbDangMo(p)
    daMo = Not wb Is Nothing
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = MoFileDB(p)
        On Error GoTo 0
        KichHoat wbMe
    End If
    If wb Is Nothing Then
        GL_LogDong "LOI DB dung chung", p, "khong doc duoc - dung du lieu sheet " & SH_DB
        Exit Sub
    End If
    Set src = wb.Worksheets(1)
    ws.Range("A2:P" & ws.Rows.Count).ClearContents
    n = DongCuoiNgoai(src)
    If n >= 2 Then
        ws.Range("A2:P" & n).NumberFormat = "@"
        ws.Range("A2:P" & n).Value = src.Range("A2:P" & n).Value
        ws.Range("E2:F" & n).NumberFormat = "General"
        ws.Range("J2:J" & n).NumberFormat = "General"
    End If
    If Not daMo Then wb.Close False
    KichHoat wbMe
End Sub

' Dong cuoi co du lieu cua sheet file DB ngoai. Khong dung End(xlUp): file do khong phai file dang chon
' (LibreOffice tinh End theo cua so dang chon). DB khong co dong trong o giua.
Private Function DongCuoiNgoai(ByVal ws As Worksheet) As Long
    Dim r As Long
    r = 1
    Do While Trim(CStr(ws.Cells(r + 1, 1).Value)) <> ""
        r = r + 1
    Loop
    DongCuoiNgoai = r
End Function

' Ghi cac chung tu (mang "LOAI|SO") tu sheet SAP_GiaLap_DB ra file DB ngoai: xoa dong cu cua chung tu do, chep dong moi
Private Sub DbDongBo(ByVal khoa As Variant)
    Dim wb As Workbook, ws As Worksheet, src As Worksheet, k As Variant, a() As String, r As Long, n As Long, m As Long
    Dim giu As New Collection, v As Variant, bo As Boolean, ten As String
    If cDbFile = "" Then Exit Sub
    Set src = SheetDB()
    Set wb = DbMoNgoai()
    On Error GoTo Loi
    Set ws = wb.Worksheets(1)
    ' 1) giu cac dong KHONG thuoc chung tu vua luu (doc het ra truoc, khong xoa dong tren file)
    n = DongCuoiNgoai(ws)
    For r = 2 To n
        bo = False
        For Each k In khoa
            If Trim(CStr(k)) <> "" Then
                a = Split(CStr(k), "|")
                If CStr(ws.Cells(r, 1).Value) = a(0) And Trim(CStr(ws.Cells(r, 2).Value)) = a(1) Then bo = True: Exit For
            End If
        Next k
        If Not bo Then giu.Add ws.Range("A" & r & ":P" & r).Value
    Next r
    ' 2) them dong moi cua cac chung tu do tu sheet SAP_GiaLap_DB
    m = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    For Each k In khoa
        If Trim(CStr(k)) <> "" Then
            a = Split(CStr(k), "|")
            For r = 2 To m
                If CStr(src.Cells(r, 1).Value) = a(0) And Trim(CStr(src.Cells(r, 2).Value)) = a(1) Then giu.Add src.Range("A" & r & ":P" & r).Value
            Next r
        End If
    Next k
    ' 3) ghi lai toan bo
    ws.Range("A2:P" & Application.WorksheetFunction.Max(n, 2) + giu.Count + 1).ClearContents
    r = 2
    For Each v In giu
        ws.Range("B" & r & ":D" & r).NumberFormat = "@"
        ws.Range("H" & r & ":I" & r).NumberFormat = "@"
        ws.Range("A" & r & ":P" & r).Value = v
        r = r + 1
    Next v
    DbDongNgoai wb
    GL_LogDong "DB dung chung", DbDuongDan(), "ghi " & Join(khoa, ", ")
    Exit Sub
Loi:
    ten = Err.Description
    On Error Resume Next
    wb.Close False                           ' khong de file DB mo dang do (lenh sau se bi khoa)
    KichHoat ThisWorkbook
    On Error GoTo 0
    GL_LogDong "LOI DB dung chung", DbDuongDan(), ten
    Err.Raise 625, NGUON, "Loi ghi file DB gia lap: " & ten
End Sub

Private Sub DbThem(ByVal a As Variant)
    Dim ws As Worksheet, r As Long, i As Long
    Set ws = SheetDB()
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    For i = 0 To UBound(a)
        If i = 1 Or i = 2 Or i = 3 Or i = 7 Or i = 8 Then
            ws.Cells(r, i + 1).NumberFormat = "@"
            ws.Cells(r, i + 1).Value = CStr(a(i))
        ElseIf i = 14 Then
            ws.Cells(r, i + 1).Value = Format(a(i), "dd/mm/yyyy hh:mm:ss")
        Else
            ws.Cells(r, i + 1).Value = a(i)
        End If
    Next i
End Sub

Private Function SoO(ByVal v As Variant) As Double
    If IsNumeric(v) And Trim(CStr(v)) <> "" Then SoO = CDbl(v)
End Function

' Doc so go vao SAP theo dau thap phan cua SAP (cDec). Dau kia = phan cach hang nghin
' (giong SAP: "16.5" khi SAP dung dau "," -> 165).
Private Function DocSo(ByVal s As String, ByRef ok As Boolean) As Double
    Dim th As String, t As String
    th = IIf(cDec = ",", ".", ",")
    t = Replace(Replace(Trim(s), " ", ""), th, "")
    ok = False
    If t = "" Then ok = True: Exit Function
    If Len(t) - Len(Replace(t, cDec, "")) > 1 Then Exit Function
    t = Replace(t, cDec, ".")
    If t Like "*[!0-9.-]*" Then Exit Function
    DocSo = Val(t)
    ok = True
End Function

' Hien so theo kieu SAP (dau thap phan cDec, hang nghin dau kia), vd 1.234,500
Private Function DinhDangSo(ByVal v As Double, ByVal soLe As Long) As String
    Dim s As String, p As Long, ip As String, fp As String, th As String, outS As String, i As Long, am As Boolean
    am = (v < 0)
    v = Abs(v)
    s = Trim(Str(Round(v, soLe)))
    If Left(s, 1) = "." Then s = "0" & s
    p = InStr(s, ".")
    If p > 0 Then
        ip = Left(s, p - 1): fp = Mid(s, p + 1)
    Else
        ip = s: fp = ""
    End If
    fp = Left(fp & String(soLe, "0"), soLe)
    th = IIf(cDec = ",", ".", ",")
    For i = Len(ip) To 1 Step -1
        outS = Mid(ip, i, 1) & outS
        If (Len(ip) - i + 1) Mod 3 = 0 And i > 1 Then outS = th & outS
    Next i
    If soLe > 0 Then outS = outS & cDec & fp
    If am Then outS = "-" & outS
    DinhDangSo = outS
End Function
