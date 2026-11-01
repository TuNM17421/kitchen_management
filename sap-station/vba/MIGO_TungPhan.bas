Attribute VB_Name = "MIGO_TungPhan"
Option Explicit

' ================================================================================
'  MIGO_TungPhan v2-Tram (30/10/2026) - NHAP KHO THEO PHIEU (chay tay hoac qua tram SAP)
'  Dung CUNG file voi module SAP_Automation v22 (dung lai cac ham: GetSapSession,
'  FindCtl, FindCtlSoft, MigoRowCell, TblColIndex, CheckPopup, CheckSbarError,
'  ExtractDocNumber, SapNumber, ParseSapNum, WaitSap).
'
'  Sheet "MIGO_TP" (macro tu tao neu chua co), du lieu tu dong 3:
'   A Ma_phieu | B So_PO | C Ngay_hach_toan (dd.mm.yyyy) | D Dong_PO (10,20..)
'   E Ma_vat_tu | F So_luong | G Text | H So_chung_tu_MIGO (macro ghi) | I Status (macro ghi)
'  Dan du lieu tu trang "Quy trinh PR truoc" (buoc 4.3) vao A3.
'
'  Moi Ma_phieu = 1 chung tu MIGO: Goods Receipt / Purchase Order / 101,
'  ngay hach toan = ngay giao, chi tich OK cac dong PO co trong phieu,
'  so luong = so thuc nhan. Dong PO khong co trong phieu se bo tich.
'
'  CHE DO KHONG LUU: MIGO_CHO_PHEP_POST = False -> macro dien man hinh, bam
'  Check, ghi ket qua, KHONG Post. Doi thanh True khi duoc phep ap dung chinh thuc.
' ================================================================================

Public Const MIGO_CHO_PHEP_POST As Boolean = False   ' chay tay (Alt+F8). Qua tram: theo che do luu tren giao dien

Public Const MIGO_DATE_FMT As String = "dd.mm.yyyy"   ' theo SU3 > Defaults > Date format
Private Const TP_TBL As String = "tblSAPLMIGOTV_GOITEM"
Private tpStep As String

Function MigoDuocPost() As Boolean
    If TRAM_AUTO Then MigoDuocPost = TRAM_LUU_THAT Else MigoDuocPost = MIGO_CHO_PHEP_POST
End Function

Sub EnsureMigoTP()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("MIGO_TP")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "MIGO_TP"
        ws.Range("A1:I1").Value = Array("Ma_phieu", "So_PO", "Ngay_hach_toan", "Dong_PO", "Ma_vat_tu", "So_luong", "Text", "So_chung_tu_MIGO", "Status")
        ws.Range("A2:I2").Value = Array("Dan tu trang web", "", "dd.mm.yyyy", "10,20..", "", "so thuc nhan", "", "<< macro ghi >>", "<< macro ghi >>")
        ws.Range("A1:I1").Font.Bold = True
        ws.Range("A1:I1").Interior.Color = RGB(221, 235, 247)
        ws.Range("A2:G2").Interior.Color = RGB(255, 242, 204)
        ws.Columns("A:I").ColumnWidth = 16
        ws.Columns("C").NumberFormat = "@"
    End If
End Sub

Sub PostMIGO_TungPhan()
    EnsureMigoTP
    Dim ws As Worksheet, lastR As Long, r As Long, k As String
    Set ws = ThisWorkbook.Sheets("MIGO_TP")
    lastR = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim phieu As Object: Set phieu = CreateObject("Scripting.Dictionary")
    For r = 3 To lastR
        k = Trim(CStr(ws.Cells(r, "A").Value))
        If k <> "" And Trim(CStr(ws.Cells(r, "H").Value)) = "" Then
            If Not phieu.Exists(k) Then phieu.Add k, New Collection
            phieu(k).Add r
        End If
    Next r
    If phieu.Count = 0 Then
        Hoi "Khong co phieu nao can MIGO (cot H trong).", vbInformation, "MIGO tung phan v2"
        Exit Sub
    End If

    Dim sm As String, v As Variant
    For Each v In phieu.Keys
        sm = sm & "  - " & v & ": PO " & ws.Cells(phieu(v)(1), "B").Value & ", ngay " & ws.Cells(phieu(v)(1), "C").Value & ", " & phieu(v).Count & " dong" & vbCrLf
    Next v
    If Hoi("Se nhap kho tung phan " & phieu.Count & " phieu:" & vbCrLf & sm & vbCrLf & _
              IIf(MigoDuocPost(), "Macro se HOI truoc khi Post tung phieu.", "CHE DO KHONG LUU: chi dien man hinh + Check, KHONG Post.") & vbCrLf & "Bat dau?", _
              vbYesNo + vbQuestion, "MIGO tung phan v2") = vbNo Then Exit Sub

    Dim session As Object: Set session = GetSapSession()
    If session Is Nothing Then Exit Sub

    Dim rows As Collection, msg As String, docNo As String, ans As VbMsgBoxResult, nOK As Long
    For Each v In phieu.Keys
        Set rows = phieu(v)
        On Error GoTo LoiPhieu
        nOK = DienMIGOTungPhan(session, ws, rows)
        If MigoDuocPost() Then
            ans = Hoi("Phieu " & v & ": da dien " & nOK & "/" & rows.Count & " dong." & vbCrLf & "Kiem tra tren SAP roi tra loi: POST nhap kho?", vbYesNo + IIf(nOK = rows.Count, vbQuestion, vbExclamation) + vbDefaultButton2, "MIGO tung phan v2")
            If ans = vbNo Then GhiStatus ws, rows, "", "CHUA POST - dung theo yeu cau": Exit Sub
            tpStep = "bam Post"
            session.findById("wnd[0]/tbar[0]/btn[11]").press
            CheckPopup session
            msg = session.findById("wnd[0]/sbar").Text
            CheckSbarError session
            docNo = ExtractDocNumber(msg)
            GhiStatus ws, rows, docNo, msg
        Else
            tpStep = "bam Check"
            On Error Resume Next
            session.findById("wnd[0]/tbar[1]/btn[7]").press   ' Check (neu co)
            On Error GoTo LoiPhieu
            WaitSap session
            msg = session.findById("wnd[0]/sbar").Text
            GhiStatus ws, rows, "", "DA DIEN + CHECK, CHUA POST (che do khong luu) " & Format(Now, "dd/mm hh:nn") & IIf(msg <> "", " | SAP: " & msg, "")
            If Hoi("Phieu " & v & ": da dien " & nOK & "/" & rows.Count & " dong, chua Post." & vbCrLf & "Kiem tra man hinh SAP. Chuyen sang phieu tiep theo?", vbYesNo + vbQuestion, "MIGO tung phan v2") = vbNo Then Exit Sub
        End If
        GoTo TiepTheo
LoiPhieu:
        GhiStatus ws, rows, "", "LOI buoc [" & tpStep & "]: " & Err.Description
        Hoi "Loi phieu " & v & ":" & vbCrLf & Err.Description & vbCrLf & "Buoc: " & tpStep & vbCrLf & _
               "Man hinh SAP giu nguyen. Neu loi 'khong tim thay control': mo rong phan Header/Item cua MIGO, chay SAP_ChanDoan va gui file.", vbCritical, "MIGO tung phan v2"
        Exit Sub
TiepTheo:
        On Error GoTo 0
    Next v
    Hoi "Xong. Kiem tra cot H/I sheet MIGO_TP, roi copy cot A + H dan vao trang web (buoc 4.3). Nho Ctrl+S.", vbInformation, "MIGO tung phan v2"
End Sub

Private Sub GhiStatus(ws As Worksheet, rows As Collection, ByVal docNo As String, ByVal st As String)
    Dim r As Variant
    For Each r In rows
        ws.Cells(r, "H").NumberFormat = "@"
        If docNo <> "" Then ws.Cells(r, "H").Value = docNo
        ws.Cells(r, "I").Value = st
    Next r
End Sub

' Dien 1 phieu vao MIGO (chua Post). Tra ve so dong da tich OK.
Function DienMIGOTungPhan(session As Object, ws As Worksheet, rows As Collection) As Long
    Dim soPO As String, ngay As String, phieuId As String, c As Object, tbl As Object
    Dim cOK As Long, cTxt As Long, cQty As Long, cItem As Long, nRows As Long, i As Long, r As Variant
    Dim want As Object: Set want = CreateObject("Scripting.Dictionary")   ' Dong_PO -> row Excel

    soPO = Trim(CStr(ws.Cells(rows(1), "B").Value))
    ngay = Trim(CStr(ws.Cells(rows(1), "C").Value))
    phieuId = Trim(CStr(ws.Cells(rows(1), "A").Value))
    For Each r In rows
        want(CLng(Val(ws.Cells(r, "D").Value))) = r
    Next r

    tpStep = "mo MIGO"
    SapCuaSo session
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmigo"
    session.findById("wnd[0]").sendVKey 0
    CheckPopup session
    Set c = FindCtlSoft(session, "cmbGODYNPRO-ACTION")
    If Not c Is Nothing Then
        If c.key <> MIGO_ACTION Then c.key = MIGO_ACTION: CheckPopup session
    End If
    Set c = FindCtlSoft(session, "cmbGODYNPRO-REFDOC")
    If Not c Is Nothing Then
        If c.key <> MIGO_REFDOC Then c.key = MIGO_REFDOC: CheckPopup session
    End If

    tpStep = "go So PO " & soPO
    FindCtl(session, "ctxtGODYNPRO-PO_NUMBER").Text = soPO
    session.findById("wnd[0]").sendVKey 0
    CheckPopup session
    CheckSbarError session

    tpStep = "ngay chung tu / ngay hach toan " & ngay
    FindCtl(session, "ctxtGOHEAD-BLDAT").Text = ngay
    FindCtl(session, "ctxtGOHEAD-BUDAT").Text = ngay
    Set c = FindCtlSoft(session, "txtGOHEAD-LFSNR")
    If Not c Is Nothing Then c.Text = Left(phieuId, 16)
    session.findById("wnd[0]").sendVKey 0
    CheckPopup session

    tpStep = "doc bang dong MIGO"
    Set tbl = FindCtl(session, TP_TBL)
    nRows = tbl.RowCount
    cOK = TblColIndex(tbl, "GOITEM-TAKE_IT")
    cTxt = TblColIndex(tbl, "GOITEM-SGTXT")
    cQty = TblColIndex(tbl, "GOITEM-ERFMG")
    cItem = TblColIndex(tbl, "GOITEM-EBELP")
    If cOK < 0 Or cQty < 0 Or cItem < 0 Then
        Err.Raise vbObjectError + 801, "MIGO_TP", "Bang dong MIGO thieu cot OK (TAKE_IT) / So luong (ERFMG) / Dong PO (EBELP). Hien cot bang Layout roi chay lai."
    End If

    Dim it As Long, cel As Object, nOK As Long, seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    For i = 0 To nRows - 1
        tpStep = "dong MIGO " & (i + 1)
        it = CLng(Val(MigoRowCell(session, i, cItem).Text))
        If it = 0 Then GoTo NextRow
        Set cel = MigoRowCell(session, i, cOK)
        If want.Exists(it) Then
            MigoRowCell(session, i, cQty).Text = SapNumber(CDbl(ws.Cells(want(it), "F").Value))
            If cTxt >= 0 Then MigoRowCell(session, i, cTxt).Text = CStr(ws.Cells(want(it), "G").Value)
            Set cel = MigoRowCell(session, i, cOK)
            If Not cel.Selected Then cel.Selected = True
            seen(it) = True
            nOK = nOK + 1
        Else
            If cel.Selected Then cel.Selected = False
        End If
NextRow:
    Next i
    session.findById("wnd[0]").sendVKey 0
    CheckPopup session
    CheckSbarError session

    Dim miss As String, k As Variant
    For Each k In want.Keys
        If Not seen.Exists(k) Then miss = miss & k & " "
    Next k
    If miss <> "" Then Err.Raise vbObjectError + 802, "MIGO_TP", "PO " & soPO & " khong con (hoac khong co) dong: " & miss & "- kiem tra so dong PO tren ME23N / PO da nhan du?"
    DienMIGOTungPhan = nOK
End Function
