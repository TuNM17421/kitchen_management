Attribute VB_Name = "SAP_Automation"
Option Explicit

' >>> CAU HINH CHUNG <<<
' Ma ky thuat cua Loai PR "Expense PR" - da xac nhan tu ban ghi thuc te = "ZEX"
Public Const PR_DOC_TYPE As String = "ZEX"

' Phien ban macro - hien tren tieu de hop thoai de biet dang chay dung ban
Public Const MACRO_VERSION As String = "v23.1-Tram-gl.2 (31/10/2026)"

' Dau thap phan SAP dang dung cho tai khoan cua ban (SU3 > Defaults > Decimal Notation):
'   "."  neu SAP hien so dang 1,234,567.89   (mac dinh)
'   ","  neu SAP hien so dang 1.234.567,89
' Chi anh huong so luong/don gia co phan le (vd 16,5 kg).
Public Const SAP_DECIMAL_SEP As String = ","
' (Macro se TU KIEM TRA lai: neu SAP khong nhan so le theo dau nay, tu doi sang dau kia)

' Loai Account Assignment Category bat buoc cho PR Expense (K = Cost center)
Public Const ACC_ASSIGN_CAT As String = "K"

' Bien dung chung de tim control SAP (khong can sua)
Private Const GRID_ID As String = "SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell"
Private Const VENDOR_TITLE As String = "3) MA NHA CUNG CAP SAP (LIFNR) - dung de tao PO"
' Dat True khi da co script ME21N (sau khi ban gui file record)
Private Const PO_SCRIPT_READY As Boolean = True
' Loai PO (da xac nhan tu file record ME21N)
Public Const PO_DOC_TYPE As String = "ZEXP"
' Buoc nhay so dong item PR (10, 20, 30...) - chi dung khi cot P (PR_item) trong
Public Const PR_ITEM_STEP As Long = 10
Private Const PO_TBL_ID As String = "tblSAPLMEGUITC_1211"
Private Const COND_TBL_ID As String = "tblSAPLV69ATCTRL_KONDITIONEN"
' MIGO: Goods Receipt (A01) / Purchase Order (R01) / loai nhap 101
Public Const MIGO_ACTION As String = "A01"
Public Const MIGO_REFDOC As String = "R01"
Public Const MIGO_BWART As String = "101"
Private Const MIGO_TBL_ID As String = "tblSAPLMIGOTV_GOITEM"
Private curStep As String
Private gPOWarn As String         ' canh bao khi nhap gia PO (PBXX khong duoc...) - hien o hop thoai truoc khi Save
Private ctlCache As Object
Private gDecSep As String        ' dau thap phan SAP dang dung (tu phat hien)
Private gLastQtyDisp As String   ' gia tri so luong SAP hien sau khi nhap
Private gLastQ As Double, gLastP As Double   ' SL / gia vua doc lai tu SAP
Private gMigoNote As String                 ' ghi chu tinh trang MIGO cua PO dang xu ly
Private Const UNREAD As String = "(khong doc lai duoc tu SAP)"
' v23-Tram: bien dung chung khi chay qua tram SAP (xem cuoi module)
Public TRAM_AUTO As Boolean
Public TRAM_LUU_THAT As Boolean
Public TRAM_LOG As String
Public TRAM_PHIEN As Long      ' cua so SAP (session) tram dung: 0 = cua so dau tien
Public TRAM_AN_SAP As Boolean  ' True = thu nho cua so SAP khi tram chay (chay ngam)

' ================================================================================
'  SAP GUI SCRIPTING AUTOMATION: TAO PR (ME51N) -> PO (ME21N) -> NHAP KHO (MIGO)
'  TU DU LIEU FILE EXCEL
' ================================================================================
'  CACH CHAY:
'   1) Mo SAP Logon, DANG NHAP vao he thong (de san 1 cua so SAP bat ky).
'   2) Mo file .xlsm nay, bam "Enable Content" neu Excel hoi.
'   3) Nhap/dan du lieu vao sheet CongNo (cot D Ma_vat_tu la bat buoc).
'   4) Alt+F8 -> chon CreatePR -> Run.
'  Macro se: kiem tra du lieu -> hien bang tom tat so PR se tao -> hoi xac nhan
'  -> nhap tung PR vao SAP -> dung lai hoi truoc khi Save moi PR.
' ================================================================================


' --------------------------------------------------------------------------
' Ket noi toi phien SAP GUI dang mo san (phai dang nhap SAP truoc khi chay)
' --------------------------------------------------------------------------
Function GetSapSession() As Object
    Dim SapGuiAuto As Object
    Dim objApp As Object
    Dim objConn As Object
    Dim objSession As Object

    ' SAP GIA LAP (test khong can SAP that): neu co module SapGiaLap va sheet
    ' SAP_GiaLap dang BAT thi dung SAP gia lap (cua so so TRAM_PHIEN khi chay qua tram).
    ' Khong co module / gia lap TAT -> chay SAP that nhu cu.
    Dim objGiaLap As Object, idxGL As Long
    If TRAM_AUTO Then idxGL = TRAM_PHIEN
    On Error Resume Next
    Set objGiaLap = Application.Run("SapGiaLap_PhienNeuBat", idxGL)
    Err.Clear
    If Not objGiaLap Is Nothing Then
        On Error GoTo 0
        Set GetSapSession = objGiaLap
        Exit Function
    End If

    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If SapGuiAuto Is Nothing Then
        Hoi "Khong tim thay SAP GUI dang chay." & vbCrLf & _
               "Vui long mo SAP Logon va dang nhap vao he thong truoc khi chay macro.", vbCritical
        Exit Function
    End If

    Set objApp = SapGuiAuto.GetScriptingEngine
    If objApp Is Nothing Then
        Hoi "Khong lay duoc Scripting Engine." & vbCrLf & _
               "Kiem tra SAP Logon > Options > Accessibility & Scripting > Scripting > 'Enable scripting'.", vbCritical
        Exit Function
    End If

    If objApp.Children.Count = 0 Then
        Hoi "SAP Logon dang mo nhung CHUA DANG NHAP vao he thong nao." & vbCrLf & _
               "Hay dang nhap SAP truoc, roi chay lai macro.", vbCritical
        Exit Function
    End If

    Set objConn = objApp.Children(0)      ' Ket noi dau tien dang mo
    Dim idxPhien As Long
    If TRAM_AUTO Then idxPhien = TRAM_PHIEN
    If idxPhien > objConn.Children.Count - 1 Then
        Hoi "SAP chi dang mo " & objConn.Children.Count & " cua so, tram can cua so thu " & (idxPhien + 1) & ".", vbCritical
        Exit Function
    End If
    Set objSession = objConn.Children(idxPhien + 0)  ' Session tram duoc giao (mac dinh: cua so dau tien)

    If objSession Is Nothing Then
        Hoi "Khong lay duoc cua so SAP (session)." & vbCrLf & _
               "Neu SAP bao 'scripting is disabled' thi can IT/Basis bat tham so sapgui/user_scripting.", vbCritical
        Exit Function
    End If

    On Error GoTo 0
    Set GetSapSession = objSession
End Function

' --------------------------------------------------------------------------
' .Select / .Type / .Close cua SAP GUI goi qua 3 ham nay: voi SAP that y het
' goi truc tiep; voi SAP gia lap (VBA khong cho dat ham ten Select/Type/Close)
' thi goi ChonTab / Loai / DongCuaSo.
' --------------------------------------------------------------------------
Public Function LaSapGiaLap(o As Object) As Boolean
    On Error Resume Next
    LaSapGiaLap = CallByName(o, "LaGiaLap", VbGet)
    Err.Clear
End Function

Public Sub SapSelect(o As Object)
    If LaSapGiaLap(o) Then
        CallByName o, "ChonTab", VbMethod
    Else
        o.Select
    End If
End Sub

Public Function SapType(o As Object) As String
    If LaSapGiaLap(o) Then
        SapType = CallByName(o, "Loai", VbGet)
    Else
        SapType = o.Type
    End If
End Function

Public Sub SapClose(o As Object)
    If LaSapGiaLap(o) Then
        CallByName o, "DongCuaSo", VbMethod
    Else
        o.Close
    End If
End Sub


' --------------------------------------------------------------------------
' Tach so chung tu (chuoi so) tu dong thong bao trang thai cua SAP
' Vi du: "Yeu cau mua hang 1000123456 da duoc tao" -> tra ve "1000123456"
' --------------------------------------------------------------------------
Function ExtractDocNumber(ByVal msg As String) As String
    Dim i As Long, ch As String, result As String
    Dim inNumber As Boolean
    result = ""
    For i = 1 To Len(msg)
        ch = Mid(msg, i, 1)
        If ch >= "0" And ch <= "9" Then
            result = result & ch
            inNumber = True
        ElseIf inNumber And Len(result) >= 8 Then
            Exit For
        Else
            result = ""
            inNumber = False
        End If
    Next i
    ExtractDocNumber = result
End Function


' --------------------------------------------------------------------------
' TIM CONTROL SAP THEO TEN (khong phu thuoc so subscreen 0015/0016/0019...)
' idPart = phan cuoi duy nhat cua ID, vd "ctxtCOBL-KOSTL", "tabpTABREQDT7"
' Thu duong dan tim duoc lan truoc (nhanh), neu khong con thi quet lai man hinh.
' --------------------------------------------------------------------------
Function FindCtlSoft(session As Object, ByVal idPart As String) As Object
    Dim c As Object
    If ctlCache Is Nothing Then Set ctlCache = CreateObject("Scripting.Dictionary")
    WaitSap session
    On Error Resume Next
    ' (a) duong dan tim duoc lan truoc
    If ctlCache.Exists(idPart) Then
        Set c = session.findById(ctlCache(idPart))
        If Err.Number <> 0 Then Set c = Nothing: Err.Clear
    End If
    ' (b) duong dan da record, thu voi cac so subscreen khac nhau
    If c Is Nothing Then Set c = TryKnownPaths(session, idPart)
    ' (c) ham tim san cua SAP theo ten + loai control
    If c Is Nothing Then Set c = NativeFind(session, idPart)
    ' (d) quet toan bo man hinh
    If c Is Nothing Then Set c = SearchCtl(session.findById("wnd[0]/usr"), idPart)
    If Not c Is Nothing Then ctlCache(idPart) = c.id
    Err.Clear
    On Error GoTo 0
    Set FindCtlSoft = c
End Function

' Dung ham FindByName / FindAllByName co san cua SAP GUI
Function NativeFind(session As Object, ByVal idPart As String) As Object
    Dim usr As Object, c As Object, allC As Object, k As Long
    Dim prefixes As Variant, types As Variant, p As Long, nm As String
    On Error Resume Next
    Set usr = session.findById("wnd[0]/usr")
    If usr Is Nothing Then Exit Function

    If InStr(idPart, "cntlGRIDCONTROL") > 0 Then
        ' Bang item: lay dung GRIDCONTROL thuoc SAPLMEGUI:3212
        Set allC = usr.FindAllByName("GRIDCONTROL", "GuiCustomControl")
        For k = 0 To allC.Count - 1
            If InStr(allC.ElementAt(k + 0).id, "SAPLMEGUI:3212") > 0 Then
                Set c = allC.ElementAt(k + 0).findById("shellcont/shell")
                If Not c Is Nothing Then Set NativeFind = c: Exit Function
            End If
        Next k
        Exit Function
    End If

    If InStr(idPart, "/") > 0 Then Exit Function
    prefixes = Array("cmb", "ctxt", "txt", "tabp", "chk", "btn", "rad", "tbl")
    types = Array("GuiComboBox", "GuiCTextField", "GuiTextField", "GuiTab", "GuiCheckBox", "GuiButton", "GuiRadioButton", "GuiTableControl")
    For p = 0 To UBound(prefixes)
        If Left(idPart, Len(prefixes(p))) = prefixes(p) Then
            nm = Mid(idPart, Len(prefixes(p)) + 1)
            Set c = usr.FindByName(CStr(nm), CStr(types(p)))
            If Err.Number <> 0 Then Set c = Nothing: Err.Clear
            If Not c Is Nothing Then Set NativeFind = c
            Exit Function
        End If
    Next p
End Function

' Cho SAP xu ly xong (toi da 15 giay)
Sub WaitSap(session As Object)
    Dim t As Single
    On Error Resume Next
    t = Timer
    Do While session.Busy
        DoEvents
        If Timer - t > 15 Or Timer < t Then Exit Do
    Loop
End Sub

Function FindCtl(session As Object, ByVal idPart As String) As Object
    Dim c As Object
    Set c = FindCtlSoft(session, idPart)
    If c Is Nothing Then
        Err.Raise vbObjectError + 513, "FindCtl", _
            "Khong tim thay control '" & idPart & "' tren man hinh SAP" & _
            IIf(InStr(idPart, "GRIDCONTROL") > 0 Or InStr(idPart, "TABREQDT") > 0 Or InStr(idPart, "COBL") > 0, _
                " (kiem tra phan Item Overview / Item Detail trong ME51N co dang MO RONG - Expand - khong)", "")
    End If
    Set FindCtl = c
End Function

' Lay control con thu idx. Dung (idx + 0) de truyen THEO GIA TRI - SAP GUI
' tu choi tham so truyen ByRef tu bien VBA (loi am tham khi duyet cay control).
Function ChildAt(parent As Object, ByVal idx As Long) As Object
    Dim c As Object
    On Error Resume Next
    Set c = parent.Children.ElementAt(idx + 0)
    If c Is Nothing Then
        Err.Clear
        Set c = parent.Children.Item(idx + 0)
    End If
    Set ChildAt = c
End Function

' Duong dan da record (phan SAU subscreen SUB0 hay thay doi so 0010..0030)
Function KnownTail(ByVal idPart As String) As String
    Const TAB7 As String = "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT7"
    Const ACC As String = "/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/"
    Select Case idPart
        Case GRID_ID: KnownTail = "subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell"
        Case "cmbMEREQ_TOPLINE-BSART": KnownTail = "subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/cmbMEREQ_TOPLINE-BSART"
        Case "tabpTABREQDT7": KnownTail = TAB7
        Case "cmbMEACCT1200-KNTTP": KnownTail = TAB7 & "/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB1:SAPLMEACCTVI:1200/cmbMEACCT1200-KNTTP"
        Case "cmbMEPO_TOPLINE-BSART": KnownTail = "subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART"
        Case "ctxtMEPO_TOPLINE-SUPERFIELD": KnownTail = "subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD"
        Case PO_TBL_ID: KnownTail = "subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211"
        Case "tabpTABIDT9": KnownTail = "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT9"
        Case COND_TBL_ID: KnownTail = "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT9/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1333/ssubSUB0:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN"
        Case "cmbDYN_6000-LIST": KnownTail = "subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST"
        Case "ctxtCOBL-KOSTL": KnownTail = TAB7 & ACC & "ctxtCOBL-KOSTL"
        Case "ctxtCOBL-GEBER": KnownTail = TAB7 & ACC & "ctxtCOBL-GEBER"
        Case Else: KnownTail = ""
    End Select
End Function

' Thu duong dan da record voi moi so subscreen SUB0 tu 0010 den 0030
Function TryKnownPaths(session As Object, ByVal idPart As String) As Object
    Dim tail As String, n As Long, c As Object
    On Error Resume Next
    ' MIGO: duong dan theo file record (so SAPLMIGO:000x co the doi)
    Dim mtail As String
    Select Case idPart
        Case MIGO_TBL_ID: mtail = "subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM"
        Case "ctxtGODYNPRO-PO_NUMBER": mtail = "subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/ctxtGODYNPRO-PO_NUMBER"
    End Select
    If mtail <> "" Then
        For n = 1 To 9
            Set c = Nothing
            Set c = session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:" & Format(n, "0000") & "/" & mtail)
            Err.Clear
            If Not c Is Nothing Then Set TryKnownPaths = c: Exit Function
        Next n
        Exit Function
    End If
    On Error GoTo 0

    tail = KnownTail(idPart)
    If tail = "" Then Exit Function
    On Error Resume Next
    For n = 10 To 30
        Set c = Nothing
        Set c = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(n, "0000") & "/" & tail)
        Err.Clear
        If Not c Is Nothing Then Set TryKnownPaths = c: Exit Function
    Next n
End Function

' Quet de quy cay control, tra ve control dau tien co ID chua idPart
Function SearchCtl(parent As Object, ByVal idPart As String) As Object
    Dim i As Long, n As Long, ch As Object, res As Object, sId As String
    On Error Resume Next
    n = -1
    n = parent.Children.Count
    If n <= 0 Then Exit Function
    For i = 0 To n - 1
        Set ch = Nothing
        Set ch = ChildAt(parent, i)
        If Not ch Is Nothing Then
            sId = ""
            sId = ch.id
            If Len(sId) > 0 And InStr(1, sId, idPart, vbTextCompare) > 0 Then
                Set SearchCtl = ch
                Exit Function
            End If
            Set res = Nothing
            Set res = SearchCtl(ch, idPart)
            If Not res Is Nothing Then
                Set SearchCtl = res
                Exit Function
            End If
        End If
    Next i
End Function

' Neu SAP bat len cua so popup (wnd[1]) thi bao loi kem noi dung popup,
' khong tu bam gi ca de tranh xac nhan nham.
Sub CheckPopup(session As Object)
    Dim w As Object, t As String
    WaitSap session
    On Error Resume Next
    Set w = session.findById("wnd[1]")
    If Err.Number <> 0 Then Set w = Nothing
    Err.Clear
    On Error GoTo 0
    If Not w Is Nothing Then
        t = w.Text
        On Error Resume Next
        t = t & " - " & w.findById("usr").Children(0).Text
        On Error GoTo 0
        Err.Raise vbObjectError + 514, "CheckPopup", "SAP hien cua so popup: " & t
    End If
End Sub


' --------------------------------------------------------------------------
' CHAN DOAN: liet ke toan bo control dang co tren man hinh SAP hien tai vao
' sheet "ChanDoan". Dung khi gap loi "khong tim thay control": de nguyen man
' hinh SAP luc loi, chay macro nay, roi gui file Excel de doi chieu.
' --------------------------------------------------------------------------
Sub SAP_ChanDoan()
    Dim session As Object
    Set session = GetSapSession()
    If session Is Nothing Then Exit Sub
    Dim nCtl As Long
    nCtl = WriteDiag(session, "Chay tay SAP_ChanDoan")
    Hoi "Da ghi " & nCtl & " control vao sheet 'ChanDoan'. Luu file va gui de kiem tra.", vbInformation
End Sub

' Ghi danh sach control man hinh SAP hien tai vao sheet ChanDoan, tra ve so control
Function WriteDiag(session As Object, ByVal note As String) As Long
    On Error Resume Next
    Dim wsD As Worksheet
    Set wsD = ThisWorkbook.Sheets("ChanDoan")
    Err.Clear
    If wsD Is Nothing Then
        Set wsD = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsD.Name = "ChanDoan"
    End If
    wsD.Cells.Clear
    wsD.Range("A1:E1").Value = Array("ID (sau wnd[0]/usr/)", "Type", "Name", "Text", "Thoi diem")
    wsD.Range("E2").Value = Now
    wsD.Range("F1").Value = "Macro " & MACRO_VERSION & " | " & note
    wsD.Range("F2").Value = "Transaction: " & session.Info.Transaction & " | Title: " & session.findById("wnd[0]").Text & _
                            " | So cua so: " & session.Children.Count & " | Status bar: " & session.findById("wnd[0]/sbar").Text
    wsD.Range("F3").Value = "SAP GUI: " & session.parent.parent.MajorVersion & "." & session.parent.parent.MinorVersion & _
                            " | He thong: " & session.Info.SystemName & " | Client: " & session.Info.Client

    Dim usrC As Object
    Set usrC = session.findById("wnd[0]/usr")
    Err.Clear
    wsD.Range("F4").Value = "So control con cua usr: " & usrC.Children.Count & IIf(Err.Number <> 0, " (loi: " & Err.Description & ")", "")
    Err.Clear
    Dim testC As Object
    Set testC = ChildAt(usrC, 0)
    wsD.Range("F5").Value = "Doc control dau tien: " & IIf(testC Is Nothing, "KHONG DOC DUOC (" & Err.Description & ")", testC.id)
    Err.Clear

    Dim rowOut As Long
    rowOut = 2
    DumpCtl usrC, wsD, rowOut
    wsD.Columns("A:D").AutoFit
    WriteDiag = rowOut - 2
End Function

Sub DumpCtl(parent As Object, wsD As Worksheet, rowOut As Long)
    Dim i As Long, n As Long, ch As Object, p As Long, sId As String
    On Error Resume Next
    n = -1
    n = parent.Children.Count
    If n <= 0 Then Exit Sub
    For i = 0 To n - 1
        Set ch = Nothing
        Set ch = ChildAt(parent, i)
        If Not ch Is Nothing Then
            sId = ch.id
            p = InStr(sId, "wnd[0]/usr/")
            If p > 0 Then sId = Mid(sId, p + 11)
            wsD.Cells(rowOut, 1).Value = "'" & sId
            wsD.Cells(rowOut, 2).Value = SapType(ch)
            wsD.Cells(rowOut, 3).Value = ch.Name
            wsD.Cells(rowOut, 4).Value = "'" & Left(ch.Text, 60)
            rowOut = rowOut + 1
            If rowOut < 5000 Then DumpCtl ch, wsD, rowOut
        End If
    Next i
End Sub

' Kiem tra nhanh dang dung phien ban macro nao
Sub PhienBan()
    Hoi "Macro SAP_Automation " & MACRO_VERSION, vbInformation
End Sub


' --------------------------------------------------------------------------
' Doi so Excel thanh chuoi dung dinh dang SAP (16.5 -> "16.5" hoac "16,5")
' Str() luon dung dau "." bat ke cai dat Windows, sau do doi theo SAP_DECIMAL_SEP
' --------------------------------------------------------------------------
Function SapNumber(ByVal v As Variant) As String
    Dim s As String
    s = Trim(Str(CDbl(v)))
    If Left(s, 1) = "." Then s = "0" & s
    If gDecSep = "" Then gDecSep = SAP_DECIMAL_SEP
    SapNumber = Replace(s, ".", gDecSep)
End Function

' Doi chuoi so SAP hien (vd "11,500" / "1.234,5") thanh so, theo dau thap phan dec
Function ParseSapNum(ByVal s As String, ByVal dec As String) As Double
    Dim th As String
    th = IIf(dec = ",", ".", ",")
    s = Replace(Replace(Trim(s), " ", ""), th, "")
    s = Replace(s, dec, ".")
    ParseSapNum = Val(s)
End Function

' Doc lai so luong SAP da nhan o dong j, so sanh voi so mong muon
Function QtyMatches(session As Object, ByVal j As Long, ByVal expected As Double) As Boolean
    Dim g As Object
    On Error Resume Next
    Set g = FindCtlSoft(session, GRID_ID)
    gLastQtyDisp = ""
    gLastQtyDisp = g.getCellValue(j + 0, "MENGE")
    On Error GoTo 0
    If gLastQtyDisp = "" Then Exit Function
    QtyMatches = (Abs(ParseSapNum(gLastQtyDisp, gDecSep) - expected) < 0.0005)
End Function


' --------------------------------------------------------------------------
' GIA / SO LUONG GOC = SHEET CongNo (nguon duy nhat). Doc 1 o cua CongNo thanh so.
' Chu y: "16,5" dang chu -> 16.5 (khong de VBA hieu nham thanh 165).
' Tra ve Empty neu o trong / khong phai so.
' --------------------------------------------------------------------------
' --------------------------------------------------------------------------
' v23: Dong sheet CongNo tuong ung voi dong r cua sheet PR.
' Doc tu cong thuc o PR!C (vd =IF(CongNo!D4="","",CongNo!D4) -> 4).
' Neu co ai CHEN/XOA dong o sheet CongNo, Excel tu doi cong thuc sheet PR,
' macro van lay dung dong (truoc v23 macro gia dinh PR dong r = CongNo dong r
' -> lech 1 dong thi doc phai o trong -> go gia 0 -> SAP bao
' "Net price must be greater than 0").
' --------------------------------------------------------------------------
Function CNRow(ByVal r As Long) As Long
    Dim f As String, p As Long, d As String
    CNRow = r
    On Error Resume Next
    f = CStr(ThisWorkbook.Sheets("PR").Range("C" & r).Formula)
    On Error GoTo 0
    p = InStr(1, f, "CongNo!", vbTextCompare)
    If p = 0 Then Exit Function
    p = p + 7
    Do While p <= Len(f)
        If Mid$(f, p, 1) = "$" Or Mid$(f, p, 1) Like "[A-Za-z]" Then p = p + 1 Else Exit Do
    Loop
    Do While p <= Len(f)
        If Mid$(f, p, 1) Like "#" Then d = d & Mid$(f, p, 1): p = p + 1 Else Exit Do
    Loop
    If d <> "" Then CNRow = CLng(d)
End Function

Function CongNoNum(ByVal r As Long, ByVal col As String) As Variant
    Dim v As Variant, s As String
    v = ThisWorkbook.Sheets("CongNo").Range(col & CNRow(r)).Value
    CongNoNum = Empty
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        s = Replace(Replace(Trim(v), " ", ""), ",", ".")
        If s = "" Or s Like "*[!0-9.]*" Then Exit Function
        If Len(s) - Len(Replace(s, ".", "")) > 1 Then Exit Function
        CongNoNum = Val(s)
    ElseIf IsNumeric(v) Then
        CongNoNum = CDbl(v)
    End If
End Function

' Doc so luong (MENGE) va don gia (PREIS) cua dong j tren bang item SAP, so voi
' gia tri CongNo. Tra ve "" neu khop, UNREAD neu khong doc duoc, nguoc lai mo ta lech.
Function VerifyGridLine(session As Object, ByVal j As Long, ByVal expQ As Double, ByVal expP As Double) As String
    Dim g As Object, dq As String, dp As String, msg As String
    On Error Resume Next
    Set g = FindCtlSoft(session, GRID_ID)
    dq = "": dp = ""
    dq = Trim(CStr(g.getCellValue(j + 0, "MENGE")))
    dp = Trim(CStr(g.getCellValue(j + 0, "PREIS")))
    On Error GoTo 0
    If dq = "" Or dp = "" Then
        VerifyGridLine = UNREAD
        Exit Function
    End If
    ' Doc theo ca 2 kieu dau thap phan (SAP co the hien 245.000 hoac 245,000)
    Dim oth As String, q1 As Double, q2 As Double, p1 As Double, p2 As Double
    oth = IIf(gDecSep = ",", ".", ",")
    q1 = ParseSapNum(dq, gDecSep): q2 = ParseSapNum(dq, oth)
    p1 = ParseSapNum(dp, gDecSep): p2 = ParseSapNum(dp, oth)
    If Abs(q1 - expQ) <= 0.0005 Then
        gLastQ = q1
    ElseIf Abs(q2 - expQ) <= 0.0005 Then
        gLastQ = q2
    Else
        gLastQ = q1
        msg = msg & "SL SAP " & dq & " <> CongNo " & expQ & "; "
    End If
    If Abs(p1 - expP) <= 0.5 Then
        gLastP = p1
    ElseIf Abs(p2 - expP) <= 0.5 Then
        gLastP = p2
    Else
        gLastP = p1
        msg = msg & "Gia SAP " & dp & " <> CongNo " & Format(expP, "#,##0") & "; "
    End If
    VerifyGridLine = msg
End Function


' --------------------------------------------------------------------------
' Kiem tra 1 dong sheet PR co du du lieu de tao PR khong. Tra ve "" neu OK,
' nguoc lai tra ve mo ta loi.
' --------------------------------------------------------------------------
Function ValidatePRRow(ws As Worksheet, ByVal r As Long) As String
    Dim errs As String
    If Trim(ws.Range("B" & r).Value) = "" Then errs = errs & "thieu So_hoa_don; "
    If Trim(ws.Range("D" & r).Value) = "" Then errs = errs & "thieu Ma_co_so; "
    If Not IsNumeric(ws.Range("E" & r).Value) Or Trim(ws.Range("E" & r).Value) = "" Then
        errs = errs & "So_luong khong phai so; "
    ElseIf CDbl(ws.Range("E" & r).Value) <= 0 Then
        errs = errs & "So_luong <= 0; "
    End If
    If Not IsNumeric(ws.Range("F" & r).Value) Or Trim(ws.Range("F" & r).Value) = "" Then errs = errs & "Don_gia khong phai so; "
    If Trim(ws.Range("G" & r).Value) = "" Or Left(Trim(ws.Range("G" & r).Value), 2) = "??" Then errs = errs & "DISPO sai (kiem tra Doi_tuong_su_dung); "
    If Trim(ws.Range("H" & r).Value) = "" Then errs = errs & "thieu EKORG; "
    If Trim(ws.Range("I" & r).Value) = "" Then errs = errs & "thieu Cost center; "
    If Trim(ws.Range("J" & r).Value) = "" Then errs = errs & "thieu Quy; "
    If Trim(ws.Range("K" & r).Value) = "" Then errs = errs & "thieu EKGRP; "

    ' Doi chieu TUYET DOI voi sheet CongNo (cung dong)
    Dim cq As Variant, cp As Variant
    cq = CongNoNum(r, "F")
    cp = CongNoNum(r, "G")
    If UCase(Trim(CStr(ThisWorkbook.Sheets("CongNo").Range("D" & CNRow(r)).Value))) <> UCase(Trim(CStr(ws.Range("C" & r).Value))) Then
        errs = errs & "Ma hang khac CongNo dong " & CNRow(r) & "; "
    End If
    If IsEmpty(cq) Then
        errs = errs & "So_luong CongNo khong hop le; "
    ElseIf IsNumeric(ws.Range("E" & r).Value) Then
        If Abs(CDbl(ws.Range("E" & r).Value) - cq) > 0.0005 Then errs = errs & "So_luong PR khac CongNo (" & cq & "); "
    End If
    If IsEmpty(cp) Then
        errs = errs & "Don_gia CongNo khong hop le; "
    ElseIf IsNumeric(ws.Range("F" & r).Value) Then
        If Abs(CDbl(ws.Range("F" & r).Value) - cp) > 0.5 Then errs = errs & "Don_gia PR khac CongNo (" & cp & "); "
    End If
    ValidatePRRow = errs
End Function


' --------------------------------------------------------------------------
' Chon mat hang thu idx (0 = mat hang dau) trong o chon Item cua phan Item
' Detail. Dung danh sach Entries cua o chon (khong phu thuoc dinh dang key).
' --------------------------------------------------------------------------
Sub SelectDetailItem(session As Object, ByVal idx As Long)
    Dim cmb As Object, wantKey As String
    curStep = "chon mat hang so " & (idx + 1) & " o Item Detail"
    Set cmb = FindCtlSoft(session, "cmbDYN_6000-LIST")
    If cmb Is Nothing Then
        Err.Raise vbObjectError + 517, "SelectDetailItem", _
            "Khong tim thay o chon Item o phan Item Detail (kiem tra Item Detail trong ME51N co dang mo rong khong)"
    End If

    wantKey = ""
    On Error Resume Next
    If cmb.Entries.Count > idx Then wantKey = cmb.Entries.ElementAt(idx + 0).key
    Err.Clear
    On Error GoTo 0
    If wantKey = "" Then wantKey = Right(Space(4) & CStr(idx + 1), 4)   ' dang "   1", "  31"

    If cmb.key <> wantKey Then
        cmb.key = wantKey
        CheckPopup session
        Set cmb = FindCtl(session, "cmbDYN_6000-LIST")
    End If
    If cmb.key <> wantKey Then
        Err.Raise vbObjectError + 518, "SelectDetailItem", _
            "Khong chuyen duoc sang mat hang so " & (idx + 1) & " (o chon Item dang la '" & cmb.key & "')"
    End If
End Sub

' --------------------------------------------------------------------------
' Dam bao mat hang idx co: Acct Assignment Cat = K, Cost center, Fund dung.
' Tra ve "" neu da dung san, nguoc lai tra ve mo ta nhung gi vua bo sung.
' --------------------------------------------------------------------------
Function EnsureItemAccAssign(session As Object, ByVal idx As Long, ByVal kostl As String, ByVal geber As String) As String
    Dim tabAcc As Object, cmbK As Object, fK As Object, fG As Object, g As Object
    Dim note As String, changed As Boolean

    SelectDetailItem session, idx

    curStep = "mo tab Account Assignment (mat hang " & (idx + 1) & ")"
    Set tabAcc = FindCtl(session, "tabpTABREQDT7")
    SapSelect tabAcc
    CheckPopup session

    ' Acct Assignment Category
    curStep = "chon Acct Assignment Cat = " & ACC_ASSIGN_CAT & " (mat hang " & (idx + 1) & ")"
    Set cmbK = FindCtlSoft(session, "cmbMEACCT1200-KNTTP")
    If Not cmbK Is Nothing Then
        If Trim(cmbK.key) <> ACC_ASSIGN_CAT Then
            note = note & "Acct Cat '" & Trim(cmbK.key) & "' -> " & ACC_ASSIGN_CAT & "; "
            cmbK.key = ACC_ASSIGN_CAT
            CheckPopup session
            SelectDetailItem session, idx
            Set tabAcc = FindCtl(session, "tabpTABREQDT7")
            SapSelect tabAcc
            CheckPopup session
        End If
    Else
        ' Du phong: dat tren bang item (cot KNTTP)
        On Error Resume Next
        Dim kv As String
        kv = "?"
        Set g = FindCtlSoft(session, GRID_ID)
        kv = g.getCellValue(idx + 0, "KNTTP")
        If kv <> "?" And Trim(kv) <> ACC_ASSIGN_CAT Then
            g.modifyCell idx, "KNTTP", ACC_ASSIGN_CAT
            g.currentCellColumn = "KNTTP"
            g.pressEnter
            note = note & "Acct Cat -> " & ACC_ASSIGN_CAT & " (tren bang); "
        End If
        Err.Clear
        On Error GoTo 0
        CheckPopup session
        SelectDetailItem session, idx
        Set tabAcc = FindCtl(session, "tabpTABREQDT7")
        SapSelect tabAcc
        CheckPopup session
    End If

    ' Cost center + Fund (chi ghi khi khac gia tri mong muon)
    curStep = "nhap Cost center (mat hang " & (idx + 1) & ")"
    Set fK = FindCtl(session, "ctxtCOBL-KOSTL")
    If UCase(Trim(fK.Text)) <> UCase(Trim(kostl)) Then
        note = note & "Cost center '" & Trim(fK.Text) & "' -> " & kostl & "; "
        fK.Text = kostl
        changed = True
    End If

    curStep = "nhap Fund/Quy (mat hang " & (idx + 1) & ")"
    Set fG = FindCtl(session, "ctxtCOBL-GEBER")
    If UCase(Trim(fG.Text)) <> UCase(Trim(geber)) Then
        note = note & "Fund '" & Trim(fG.Text) & "' -> " & geber & "; "
        fG.Text = geber
        changed = True
    End If

    If changed Then
        fG.SetFocus
        session.findById("wnd[0]").sendVKey 0
        CheckPopup session

        ' doc lai de chac chan SAP da nhan
        Set fK = FindCtl(session, "ctxtCOBL-KOSTL")
        Set fG = FindCtl(session, "ctxtCOBL-GEBER")
        If UCase(Trim(fK.Text)) <> UCase(Trim(kostl)) Or UCase(Trim(fG.Text)) <> UCase(Trim(geber)) Then
            Err.Raise vbObjectError + 519, "EnsureItemAccAssign", _
                "Mat hang " & (idx + 1) & ": SAP khong nhan Cost center/Fund (dang la '" & fK.Text & "' / '" & fG.Text & "')" & _
                " | SAP bao: " & session.findById("wnd[0]/sbar").Text
        End If
    End If

    EnsureItemAccAssign = note
End Function

' --------------------------------------------------------------------------
' 1) TAO PURCHASE REQUISITION (ME51N) TU SHEET "PR" - HO TRO NHIEU MAT HANG/PR
'
' GOM NHOM: cac dong CUNG Nhom_PR (cot O, lay tu cot I sheet CongNo do ban
' tu dat) = 1 PR, co the gom nhieu nha cung cap / nhieu hoa don.
' Dong de trong Nhom_PR -> tu gop theo DISPO (be1 chung 1 PR, be2 chung 1 PR).
'
' Cot sheet "PR" (dong 1 = tieu de, dong 2 = ghi chu, du lieu tu dong 3):
'   A = STT (tu dong)       B = So_hoa_don       C = Ma_vat_tu (MATNR)
'   D = Ma_co_so (NAME1)    E = So_luong (MENGE) F = Don_gia (PREIS)
'   G = DISPO (be1/be2)     H = EKORG            I = KOSTL
'   J = GEBER (Quy)         K = EKGRP            L = SoPR (macro ghi)
'   M = Status (macro ghi)  N = Nha_cung_cap     O = Nhom_PR
'
' Dong da co SoPR (cot L) se duoc bo qua -> chay lai khong tao trung.
' --------------------------------------------------------------------------
Sub CreatePR()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PR")
    Application.Calculate
    EnsurePRHeaders ws

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    ' ---- B1. Gom nhom + kiem tra du lieu (CHUA dong vao SAP) ----
    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")

    Dim i As Long, gKey As String, vErr As String
    Dim nBad As Long, badList As String, nSkipDone As Long

    For i = 3 To lastRow
        If Trim(ws.Range("C" & i).Value) = "" Then GoTo SkipRow          ' dong trong
        If Trim(ws.Range("L" & i).Value) <> "" Then                       ' da co SoPR
            nSkipDone = nSkipDone + 1
            GoTo SkipRow
        End If

        vErr = ValidatePRRow(ws, i)
        If vErr <> "" Then
            nBad = nBad + 1
            ws.Range("M" & i).Value = "THIEU/SAI DU LIEU: " & vErr
            If nBad <= 10 Then badList = badList & "  - Dong " & i & " (" & ws.Range("C" & i).Value & "): " & vErr & vbCrLf
            GoTo SkipRow
        End If
        ws.Range("M" & i).ClearContents

        ' Khoa gom PR = cot O (Nhom_PR do ban tu dat ben CongNo, gop duoc nhieu NCC).
        ' De trong -> tu gop theo DISPO (tat ca be1 = 1 PR, tat ca be2 = 1 PR).
        gKey = Trim(ws.Range("O" & i).Value)
        If gKey = "" Then gKey = "(tu dong) DISPO " & Trim(ws.Range("G" & i).Value)
        If Not groups.Exists(gKey) Then groups.Add gKey, New Collection
        groups(gKey).Add i
SkipRow:
    Next i

    ' ---- B2. Khong co gi de tao -> bao ro ly do, KHONG dong vao SAP ----
    If groups.Count = 0 Then
        Hoi "KHONG co dong nao hop le de tao PR." & vbCrLf & vbCrLf & _
               "- Dong da co SoPR (bo qua): " & nSkipDone & vbCrLf & _
               "- Dong thieu/sai du lieu: " & nBad & vbCrLf & badList & vbCrLf & _
               "Kiem tra sheet CongNo (cot D Ma_vat_tu phai co du lieu) va cot Status (M) sheet PR.", vbExclamation, "CreatePR"
        Exit Sub
    End If

    ' ---- B3. Tom tat & xac nhan truoc khi bat dau ----
    Dim summary As String, k As Variant, nRows As Long
    Dim ncc As Object, rw As Variant
    For Each k In groups.Keys
        Set ncc = CreateObject("Scripting.Dictionary")
        For Each rw In groups(k)
            If Not ncc.Exists(Trim(ws.Range("N" & rw).Value)) Then ncc.Add Trim(ws.Range("N" & rw).Value), 1
        Next rw
        summary = summary & "  - " & k & " : " & groups(k).Count & " mat hang, " & ncc.Count & " nha cung cap" & vbCrLf
        nRows = nRows + groups(k).Count
    Next k
    If nBad > 0 Then
        summary = summary & vbCrLf & "CANH BAO: " & nBad & " dong thieu/sai du lieu se BI BO QUA:" & vbCrLf & badList
    End If

    If Hoi("Se tao " & groups.Count & " PR (" & nRows & " mat hang):" & vbCrLf & summary & vbCrLf & _
              "Bat dau nhap vao SAP?", vbYesNo + vbQuestion, "CreatePR " & MACRO_VERSION & " - Xac nhan") = vbNo Then Exit Sub

    ' ---- B4. Ket noi SAP ----
    Dim session As Object
    Set session = GetSapSession()
    If session Is Nothing Then Exit Sub

    ' ---- B5. Xu ly tung nhom = 1 PR ----
    ' Control SAP duoc TIM TU DONG theo ten (khong dung duong dan co dinh),
    ' vi so subscreen (0015/0016/0019...) thay doi tuy trang thai man hinh.
    Set ctlCache = CreateObject("Scripting.Dictionary")
    gDecSep = SAP_DECIMAL_SEP

    Dim gKeyItem As Variant, nDone As Long, nErr As Long
    Dim j As Long, r As Long
    For Each gKeyItem In groups.Keys
        Dim rowsInGroup As Collection
        Set rowsInGroup = groups(gKeyItem)
        j = -1

        On Error GoTo GroupError

        ' 0. Mo ME51N ("/n" = thoat giao dich dang mo) + chon Expense PR
        curStep = "mo ME51N"
        SapCuaSo session
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme51n"
        session.findById("wnd[0]").sendVKey 0
        CheckPopup session
        If UCase(session.Info.Transaction) <> "ME51N" Then
            Err.Raise vbObjectError + 515, "CreatePR", "SAP khong mo duoc ME51N (dang o giao dich '" & _
                session.Info.Transaction & "', tieu de '" & session.findById("wnd[0]").Text & "')"
        End If

        curStep = "chon Loai PR (Expense PR)"
        Dim cmbType As Object, tryN As Integer
        For tryN = 1 To 3
            Set cmbType = FindCtlSoft(session, "cmbMEREQ_TOPLINE-BSART")
            If Not cmbType Is Nothing Then Exit For
            Application.Wait Now + TimeSerial(0, 0, 1)
        Next tryN

        If cmbType Is Nothing Then
            ' Khong tim thay o Loai PR -> ghi chan doan, nho nguoi dung chon tay roi chay tiep
            WriteDiag session, "Khong tim thay cmbMEREQ_TOPLINE-BSART sau khi mo ME51N"
            On Error Resume Next
            TramActivate
            On Error GoTo GroupError
            If Hoi("Macro KHONG tim thay o 'Loai PR' tren man hinh ME51N" & vbCrLf & _
                      "(da ghi chi tiet vao sheet 'ChanDoan' - nho luu file gui lai de kiem tra)." & vbCrLf & vbCrLf & _
                      "Ban hay chuyen sang SAP, TU CHON loai 'Expense PR' o goc tren ben trai," & vbCrLf & _
                      "roi quay lai day bam OK de macro nhap tiep cac mat hang." & vbCrLf & _
                      "Bam Cancel de dung macro.", vbOKCancel + vbExclamation, "CreatePR " & MACRO_VERSION) = vbCancel Then
                Dim rrC As Variant
                For Each rrC In rowsInGroup
                    ws.Range("M" & rrC).Value = "CHUA TAO - dung o buoc chon Loai PR (xem sheet ChanDoan)"
                Next rrC
                Exit Sub
            End If
            ctlCache.RemoveAll
        Else
            cmbType.key = PR_DOC_TYPE
        End If
        CheckPopup session

        Dim grid As Object, cmbItem As Object, gebFld As Object, tabAcc As Object, kostlFld As Object
        Dim cmbKnttp As Object, itemKey As String, qtyVal As Double, qtyOK As Boolean
        Dim priceVal As Double, chk As String, unreadList As String
        For j = 0 To rowsInGroup.Count - 1
            r = rowsInGroup(j + 1)   ' Collection 1-based, dong grid SAP 0-based

            ' 1. Ma vat tu
            curStep = "nhap Ma vat tu " & ws.Range("C" & r).Value
            Set grid = FindCtl(session, GRID_ID)
            grid.modifyCell j, "MATNR", CStr(ws.Range("C" & r).Value)
            grid.currentCellColumn = "MATNR"
            grid.pressEnter
            CheckPopup session

            ' 2. Ma co so (NAME1)
            curStep = "nhap Ma co so (NAME1)"
            Set grid = FindCtl(session, GRID_ID)
            grid.modifyCell j, "NAME1", CStr(ws.Range("D" & r).Value)
            grid.pressEnter
            CheckPopup session

            ' 2b+3. Chon mat hang j o Item Detail -> tab Account Assignment ->
            '       Acct Assignment Cat = K -> Cost center + Fund
            EnsureItemAccAssign session, j, CStr(ws.Range("I" & r).Value), CStr(ws.Range("J" & r).Value)

            ' 4. So luong, Don gia, DISPO, EKORG, EKGRP
            curStep = "nhap So luong/Don gia/DISPO/EKORG/EKGRP"
            qtyVal = CongNoNum(r, "F")       ' SL + gia lay THANG tu sheet CongNo
            priceVal = CongNoNum(r, "G")
            Set grid = FindCtl(session, GRID_ID)
            grid.modifyCell j, "MENGE", SapNumber(qtyVal)
            grid.modifyCell j, "PREIS", SapNumber(priceVal)
            grid.modifyCell j, "DISPO", CStr(ws.Range("G" & r).Value)
            grid.modifyCell j, "EKORG", CStr(ws.Range("H" & r).Value)
            grid.modifyCell j, "EKGRP", CStr(ws.Range("K" & r).Value)
            grid.currentCellColumn = "EKGRP"
            grid.pressEnter
            CheckPopup session

            ' 4b. Ghi lai so dong item cua PR (BNFPO) vao cot P - dung khi tao PO tham chieu PR
            curStep = "doc so dong item PR"
            Dim prItemNo As String
            prItemNo = ""
            On Error Resume Next
            Set grid = FindCtlSoft(session, GRID_ID)
            prItemNo = Trim(CStr(grid.getCellValue(j + 0, "BNFPO")))
            Err.Clear
            On Error GoTo GroupError
            If prItemNo <> "" Then
                ws.Range("P" & r).NumberFormat = "@"
                ws.Range("P" & r).Value = prItemNo
            End If

            ' 4a. So luong le (vd 11,5): doc lai SAP da nhan dung chua,
            '     neu sai thi doi dau thap phan ("," <-> ".") va nhap lai
            If qtyVal <> Int(qtyVal) Then
                curStep = "kiem tra so luong le " & qtyVal
                qtyOK = QtyMatches(session, j, qtyVal)
                If Not qtyOK And gLastQtyDisp <> "" Then
                    gDecSep = IIf(gDecSep = ",", ".", ",")
                    Set grid = FindCtl(session, GRID_ID)
                    grid.modifyCell j, "MENGE", SapNumber(qtyVal)
                    grid.currentCellColumn = "MENGE"
                    grid.pressEnter
                    CheckPopup session
                    qtyOK = QtyMatches(session, j, qtyVal)
                End If
                If Not qtyOK And gLastQtyDisp <> "" Then
                    Err.Raise vbObjectError + 516, "CreatePR", "SAP khong nhan so luong le " & qtyVal & _
                        " (da thu ca dau ',' va '.'; SAP dang hien: '" & gLastQtyDisp & "')"
                End If
            End If

            ' 4c. Doc lai SL + gia tren SAP, phai BANG TUYET DOI sheet CongNo
            curStep = "kiem tra SL/gia tren SAP = CongNo"
            chk = VerifyGridLine(session, j, qtyVal, priceVal)
            If chk <> "" And chk <> UNREAD Then
                Set grid = FindCtl(session, GRID_ID)        ' nhap lai 1 lan
                grid.modifyCell j, "MENGE", SapNumber(qtyVal)
                grid.modifyCell j, "PREIS", SapNumber(priceVal)
                grid.currentCellColumn = "PREIS"
                grid.pressEnter
                CheckPopup session
                chk = VerifyGridLine(session, j, qtyVal, priceVal)
                If chk <> "" And chk <> UNREAD Then
                    Err.Raise vbObjectError + 520, "CreatePR", "Mat hang " & ws.Range("C" & r).Value & ": " & chk & _
                        "| SAP bao: " & session.findById("wnd[0]/sbar").Text
                End If
            End If
            If chk = UNREAD Then
                unreadList = unreadList & ws.Range("C" & r).Value & " "
            Else
                ws.Range("R" & r).Value = gLastP       ' gia that su tren PR (SAP)
                ws.Range("S" & r).Value = gLastQ       ' SL that su tren PR (SAP)
            End If

            ws.Range("M" & r).Value = "Da nhap vao SAP (dong " & (j + 1) & "), cho Save..."
        Next j
        j = -1

        ' 4b. KIEM TRA LAI TOAN BO: di qua TUNG mat hang trong PR, dam bao
        '     Acct Assignment Cat = K va Cost center / Fund dung - sua neu thieu
        Dim k2 As Long, fixNote As String, fixList As String, nFixed As Long
        Dim priceList As String, nPriceDiff As Long
        fixList = "": nFixed = 0: priceList = "": nPriceDiff = 0
        For k2 = 0 To rowsInGroup.Count - 1
            j = k2
            r = rowsInGroup(k2 + 1)
            fixNote = EnsureItemAccAssign(session, k2, CStr(ws.Range("I" & r).Value), CStr(ws.Range("J" & r).Value))
            If fixNote <> "" Then
                nFixed = nFixed + 1
                If nFixed <= 10 Then fixList = fixList & "  - Mat hang " & (k2 + 1) & " (" & ws.Range("C" & r).Value & "): " & fixNote & vbCrLf
            End If
            ' SL + gia lan cuoi truoc khi Save, so voi CongNo
            curStep = "kiem tra lai SL/gia mat hang " & (k2 + 1)
            chk = VerifyGridLine(session, k2, CDbl(CongNoNum(r, "F")), CDbl(CongNoNum(r, "G")))
            If chk <> "" And chk <> UNREAD Then
                nPriceDiff = nPriceDiff + 1
                If nPriceDiff <= 10 Then priceList = priceList & "  - Mat hang " & (k2 + 1) & " (" & ws.Range("C" & r).Value & "): " & chk & vbCrLf
            ElseIf chk = "" Then
                ws.Range("R" & r).Value = gLastP
                ws.Range("S" & r).Value = gLastQ
            End If
        Next k2
        j = -1

        ' 5. DUNG LAI XAC NHAN TRUOC KHI SAVE
        On Error Resume Next
        TramActivate
        On Error GoTo GroupError

        Dim answer As VbMsgBoxResult, priceMsg As String
        If nPriceDiff > 0 Then
            priceMsg = "CANH BAO - SL/GIA TREN SAP KHAC CongNo o " & nPriceDiff & " mat hang:" & vbCrLf & priceList
        ElseIf unreadList <> "" Then
            priceMsg = "SL/gia khop CongNo; RIENG cac ma sau khong doc lai duoc, hay tu kiem tra tren SAP: " & unreadList & vbCrLf
        Else
            priceMsg = "SL va gia tren SAP BANG TUYET DOI sheet CongNo o ca " & rowsInGroup.Count & " mat hang." & vbCrLf
        End If
        answer = Hoi("PR: " & CStr(gKeyItem) & vbCrLf & _
                        rowsInGroup.Count & " mat hang da duoc nhap xong." & vbCrLf & _
                        priceMsg & _
                        "Da kiem tra lai Account Assignment (K + Cost center + Fund) cho ca " & rowsInGroup.Count & " mat hang" & _
                        IIf(nFixed > 0, " - co " & nFixed & " mat hang phai bo sung:" & vbCrLf & fixList, ": deu dung.") & vbCrLf & _
                        "Hay kiem tra lai tren man hinh SAP (Alt+Tab) truoc khi tra loi." & vbCrLf & vbCrLf & _
                        "Ban co muon LUU (Save) PR nay khong?", _
                        vbYesNo + IIf(nPriceDiff > 0, vbExclamation, vbQuestion) + vbDefaultButton2, "Xac nhan truoc khi Save PR")

        If answer = vbNo Then
            Dim rrSkip As Variant
            For Each rrSkip In rowsInGroup
                ws.Range("M" & rrSkip).Value = "DA HUY THEO YEU CAU - chua Save, tu kiem tra tren SAP"
            Next rrSkip
            Hoi "Da dung macro theo yeu cau." & vbCrLf & _
                   "PR dang mo tren SAP CHUA duoc luu - ban tu kiem tra/sua/thoat man hinh do." & vbCrLf & _
                   "Chay lai CreatePR de tiep tuc cac PR con lai.", vbInformation
            Exit Sub
        End If

        ' 6. Save
        curStep = "bam Save"
        ' v24-auto2: xu ly hop thoai "Save Document - System messages issued" giong PO
        Dim msg As String, soPR As String
        soPR = SaveDocWithPopups(session, "PR")
        msg = session.findById("wnd[0]/sbar").Text

        Dim rr As Variant
        For Each rr In rowsInGroup
            ws.Range("L" & rr).Value = soPR
            ws.Range("M" & rr).Value = msg
        Next rr
        If soPR <> "" Then nDone = nDone + 1 Else nErr = nErr + 1

        GoTo NextGroup

GroupError:
        ' Ghi CHI TIET loi (mat hang nao, buoc nao, SAP bao gi) roi DUNG macro
        Dim errText As String, SbarText As String
        errText = Err.Description
        On Error Resume Next
        SbarText = session.findById("wnd[0]/sbar").Text

        Dim errFull As String
        errFull = "LOI" & IIf(j >= 0, " o mat hang thu " & (j + 1), "") & _
                  ", buoc [" & curStep & "]: " & errText & _
                  IIf(SbarText <> "", " | SAP bao: " & SbarText, "")

        Dim rr2 As Variant
        For Each rr2 In rowsInGroup
            ws.Range("M" & rr2).Value = "CHUA SAVE - " & errFull
        Next rr2
        If InStr(errText, "Khong tim thay control") > 0 Or InStr(errText, "could not be found") > 0 Then
            WriteDiag session, errFull
            errFull = errFull & vbCrLf & "(Da tu dong ghi man hinh SAP vao sheet 'ChanDoan' - hay LUU file va gui lai.)"
        End If

        TramActivate
        Hoi errFull & vbCrLf & vbCrLf & _
               "Macro da DUNG. Man hinh SAP van giu nguyen de ban xem." & vbCrLf & _
               "Gui lai file Excel (kem sheet ChanDoan neu co) de kiem tra.", vbCritical, "CreatePR " & MACRO_VERSION & " - Loi"
        Exit Sub

NextGroup:
        On Error GoTo 0
    Next gKeyItem

    Hoi "Xong. PR tao thanh cong: " & nDone & " / " & groups.Count & _
           vbCrLf & "(Dau thap phan SAP dang dung: '" & gDecSep & "')" & _
           IIf(nErr > 0, vbCrLf & "Co " & nErr & " PR loi/khong lay duoc so - xem cot Status (M).", ""), _
           IIf(nErr > 0, vbExclamation, vbInformation), "CreatePR"
End Sub


' ==========================================================================
' 2) PURCHASE ORDER (ME21N) - 1 PO cho moi NHA CUNG CAP
'
' LOGIC:
'  - Lay cac dong sheet PR DA CO SoPR (cot L) va CHUA co SoPO (cot Q).
'  - Gom theo Nha_cung_cap (cot N): moi nha cung cap = 1 PO, co the gom
'    mat hang tu nhieu PR khac nhau.
'  - GIA + SO LUONG o MOI BUOC (PR, PO) phai BANG TUYET DOI sheet CongNo.
'    Sau khi nhap PO tren SAP (truoc khi Save), macro DOC LAI tung dong PO
'    va SO VOI sheet CongNo (ma hang, so luong, don gia). Lech o dau bao o do.
'    Dong thoi so voi gia/SL da ghi nhan luc tao PR (cot R/S sheet PR) de
'    phat hien truong hop CongNo bi sua SAU khi da tao PR.
'  - Ma nha cung cap SAP (LIFNR) lay tu bang "3) MA NHA CUNG CAP SAP" ben
'    sheet Mapping (cot E:F) - nap tu file Cong no bang macro NapDuLieuCongNo.
'
' Sheet "PO" (macro TongHopPO tu tao lai):
'   A STT | B Nha_cung_cap | C Ma_NCC_SAP | D So_dong | E So_PR
'   F Tong_tien_theo_CongNo (= tong SL x Don gia sheet CongNo)
'   G Tong_tien_tren_SAP (macro doc lai tu PO tren SAP) | H Chenh_lech (= G - F)
'   I Kiem_tra | J SoPO | K Status
' ==========================================================================


' Them cot PR_item (P) va SoPO (Q) vao sheet PR neu chua co
Sub EnsurePRHeaders(wsPR As Worksheet)
    If wsPR.Range("P1").Value <> "PR_item" Then
        wsPR.Range("L1:L2").Copy
        wsPR.Range("P1:Q2").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        wsPR.Range("P1").Value = "PR_item"
        wsPR.Range("P2").Value = "<< Macro CreatePR ghi (so dong item trong PR) >>"
        wsPR.Range("Q1").Value = "SoPO"
        wsPR.Range("Q2").Value = "<< Macro CreatePO ghi vao day >>"
        wsPR.Columns("P:Q").ColumnWidth = 14
    End If
    If wsPR.Range("R1").Value <> "Gia_tren_SAP_PR" Then
        wsPR.Range("L1:L2").Copy
        wsPR.Range("R1:S2").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        wsPR.Range("R1").Value = "Gia_tren_SAP_PR"
        wsPR.Range("R2").Value = "<< Macro CreatePR doc lai tu SAP >>"
        wsPR.Range("S1").Value = "SL_tren_SAP_PR"
        wsPR.Range("S2").Value = "<< Macro CreatePR doc lai tu SAP >>"
        wsPR.Columns("R:S").ColumnWidth = 16
    End If
End Sub

' Tao bang ma nha cung cap ben sheet Mapping (cot E:G) neu chua co
Sub EnsureVendorTable(wsM As Worksheet)
    If wsM.Range("E1").Value <> VENDOR_TITLE Then
        wsM.Range("E1").Value = VENDOR_TITLE
        wsM.Range("E1").Font.Bold = True
        wsM.Range("E3:G3").Value = Array("Ten_NCC (giong cot Nha_cung_cap)", "Ma_NCC_SAP (LIFNR)", "Ghi_chu")
        wsM.Range("E3:G3").Font.Bold = True
        wsM.Range("E3:G3").Interior.Color = RGB(221, 235, 247)
        wsM.Columns("E").ColumnWidth = 32
        wsM.Columns("F").ColumnWidth = 20
        wsM.Columns("G").ColumnWidth = 30
    End If
End Sub

' Tra ma NCC SAP theo ten; neu chua co va addIfMissing=True thi them dong moi (ma de trong)
Function VendorCode(wsM As Worksheet, ByVal sup As String, ByVal addIfMissing As Boolean) As String
    Dim r As Long, lastR As Long
    lastR = wsM.Cells(wsM.Rows.Count, "E").End(xlUp).Row
    If lastR < 3 Then lastR = 3
    For r = 4 To lastR
        If UCase(Trim(CStr(wsM.Cells(r, "E").Value))) = UCase(Trim(sup)) Then
            VendorCode = Trim(CStr(wsM.Cells(r, "F").Value))
            Exit Function
        End If
    Next r
    If addIfMissing Then
        wsM.Cells(lastR + 1, "E").Value = sup
        wsM.Cells(lastR + 1, "F").Interior.Color = RGB(255, 242, 204)
        wsM.Cells(lastR + 1, "G").Value = "Can dien ma NCC SAP"
    End If
    VendorCode = ""
End Function


' --------------------------------------------------------------------------
' NAP MA NHA CUNG CAP tu file "Cong no.xlsx" (project Xu ly cong no):
' sheet "Danh muc NCC" (cot B Ma NCC, cot C Ten NCC) -> Mapping cot E:F.
' Chay lai moi khi danh muc NCC thay doi. Khong sua file Cong no.
' (Tim sheet theo tieu de cot, vi ten sheet co dau tieng Viet)
' --------------------------------------------------------------------------
Sub NapDuLieuCongNo()
    Dim f As Variant, wbSrc As Workbook, sh As Worksheet, shNCC As Worksheet
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Sheets("Mapping")

    f = Application.GetOpenFilename("Excel (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , "Chon file Cong no.xlsx (project Xu ly cong no)")
    If VarType(f) = vbBoolean Then Exit Sub

    Application.ScreenUpdating = False
    Set wbSrc = Workbooks.Open(Filename:=CStr(f), ReadOnly:=True, UpdateLinks:=0)
    For Each sh In wbSrc.Worksheets
        If Trim(CStr(sh.Range("A4").Value)) = "STT" And Trim(CStr(sh.Range("E4").Value)) = "Email" Then
            Set shNCC = sh
            Exit For
        End If
    Next sh

    EnsureVendorTable wsM
    Dim r As Long, lastR As Long, nNCC As Long, nm As String, cd As String
    If Not shNCC Is Nothing Then
        lastR = shNCC.Cells(shNCC.Rows.Count, "C").End(xlUp).Row
        For r = 5 To lastR
            nm = Trim(CStr(shNCC.Cells(r, "C").Value))
            cd = Trim(CStr(shNCC.Cells(r, "B").Value))
            If nm <> "" And cd <> "" Then
                UpsertVendor wsM, nm, cd
                nNCC = nNCC + 1
            End If
        Next r
    End If

    wbSrc.Close SaveChanges:=False
    Application.ScreenUpdating = True

    If shNCC Is Nothing Then
        Hoi "Khong tim thay sheet Danh muc NCC trong file da chon.", vbExclamation, "NapDuLieuCongNo " & MACRO_VERSION
    Else
        Hoi "Da nap " & nNCC & " ma nha cung cap vao sheet Mapping (cot E:F).", vbInformation, "NapDuLieuCongNo " & MACRO_VERSION
    End If
End Sub

' Cap nhat (hoac them moi) 1 dong ten NCC -> ma NCC trong bang Mapping E:F
Sub UpsertVendor(wsM As Worksheet, ByVal nm As String, ByVal cd As String)
    Dim r As Long, lastR As Long
    lastR = wsM.Cells(wsM.Rows.Count, "E").End(xlUp).Row
    If lastR < 3 Then lastR = 3
    For r = 4 To lastR
        If UCase(Trim(CStr(wsM.Cells(r, "E").Value))) = UCase(nm) Then Exit For
    Next r
    If r > lastR Then r = lastR + 1
    wsM.Cells(r, "E").Value = nm
    wsM.Cells(r, "F").NumberFormat = "@"
    wsM.Cells(r, "F").Value = cd
    wsM.Cells(r, "F").Interior.Pattern = xlNone
    wsM.Cells(r, "G").Value = ""
End Sub


' --------------------------------------------------------------------------
' LAY LAI SO PR TU SAP: khi PR da tao tren SAP nhung cot L (SoPR) sheet PR trong.
' B1. Tren SAP chay ME5A (danh sach PR), loc: nguoi tao = ban, ngay tao, loai ZEX
'     -> List > Export > Spreadsheet -> luu file .xlsx
' B2. Chay macro nay, chon file vua xuat. Macro ghep tung dong sheet PR (chua co
'     SoPR) voi dong PR tren SAP co CUNG ma hang + so luong + don gia (theo CongNo),
'     hien tom tat, ban dong y moi ghi So PR (cot L) + so dong item (cot P).
' --------------------------------------------------------------------------
Sub NapSoPRTuSAP()
    Dim ws As Worksheet, f As Variant, wbX As Workbook, wx As Worksheet
    Set ws = ThisWorkbook.Sheets("PR")
    EnsurePRHeaders ws

    f = Application.GetOpenFilename("Excel (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm", , "Chon file xuat tu ME5A (danh sach PR)")
    If VarType(f) = vbBoolean Then Exit Sub
    Application.ScreenUpdating = False
    Set wbX = Workbooks.Open(Filename:=CStr(f), ReadOnly:=True, UpdateLinks:=0)
    Set wx = wbX.Worksheets(1)

    ' --- Tim dong tieu de + cac cot ---
    Dim hr As Long, c As Long, lastC As Long, h As String
    Dim cPR As Long, cItem As Long, cMat As Long, cQty As Long, cPrice As Long
    For hr = 1 To 15
        cPR = 0: cItem = 0: cMat = 0: cQty = 0: cPrice = 0
        lastC = wx.Cells(hr, wx.Columns.Count).End(xlToLeft).Column
        For c = 1 To lastC
            h = LCase(Trim(CStr(wx.Cells(hr, c).Value)))
            If h = "banfn" Or ((InStr(h, "req") > 0 Or InStr(h, "pr ") = 1) And InStr(h, "item") = 0 And InStr(h, "date") = 0 _
                And InStr(h, "qty") = 0 And InStr(h, "quantity") = 0 And InStr(h, "track") = 0 And InStr(h, "type") = 0 _
                And InStr(h, "requisitioner") = 0) Then
                If cPR = 0 Then cPR = c
            ElseIf h = "bnfpo" Or InStr(h, "item") > 0 Then
                If cItem = 0 Then cItem = c
            ElseIf h = "matnr" Or h = "material" Or (Left(h, 8) = "material" And InStr(h, "group") = 0 And InStr(h, "desc") = 0 And InStr(h, "text") = 0) Then
                If cMat = 0 Then cMat = c
            ElseIf h = "menge" Or ((InStr(h, "quantity") > 0 Or InStr(h, "qty") > 0) And InStr(h, "order") = 0) Then
                If cQty = 0 Then cQty = c
            ElseIf h = "preis" Or InStr(h, "valuation price") > 0 Or InStr(h, "valn price") > 0 Or InStr(h, "valn. price") > 0 Then
                If cPrice = 0 Then cPrice = c
            End If
        Next c
        If cPR > 0 And cItem > 0 And cMat > 0 And cQty > 0 Then Exit For
    Next hr
    If cPR = 0 Or cItem = 0 Or cMat = 0 Or cQty = 0 Then
        wbX.Close False
        Application.ScreenUpdating = True
        Hoi "Khong nhan ra cac cot trong file xuat (can co: Purchase Requisition, Item, Material, Quantity)." & vbCrLf & _
               "Hay gui file xuat ME5A de kiem tra.", vbExclamation, "NapSoPRTuSAP"
        Exit Sub
    End If

    ' --- Doc danh sach PR tren SAP ---
    Dim lastX As Long, rX As Long, n As Long
    lastX = wx.Cells(wx.Rows.Count, cPR).End(xlUp).Row
    Dim xPR() As String, xItem() As String, xMat() As String, xQ() As Variant, xP() As Variant, xUsed() As Boolean
    ReDim xPR(1 To Application.Max(1, lastX)): ReDim xItem(1 To Application.Max(1, lastX)): ReDim xMat(1 To Application.Max(1, lastX))
    ReDim xQ(1 To Application.Max(1, lastX)): ReDim xP(1 To Application.Max(1, lastX)): ReDim xUsed(1 To Application.Max(1, lastX))
    For rX = hr + 1 To lastX
        If Trim(CStr(wx.Cells(rX, cPR).Value)) Like "*#*" And Trim(CStr(wx.Cells(rX, cMat).Value)) <> "" Then
            n = n + 1
            xPR(n) = Trim(CStr(wx.Cells(rX, cPR).Value))
            xItem(n) = Trim(CStr(wx.Cells(rX, cItem).Value))
            If IsNumeric(xItem(n)) Then xItem(n) = CStr(CLng(xItem(n)))
            xMat(n) = UCase(Trim(CStr(wx.Cells(rX, cMat).Value)))
            xQ(n) = wx.Cells(rX, cQty).Value
            If cPrice > 0 Then xP(n) = wx.Cells(rX, cPrice).Value Else xP(n) = Empty
        End If
    Next rX
    wbX.Close False
    Application.ScreenUpdating = True

    ' --- Ghep voi sheet PR (dong chua co SoPR) ---
    Dim lastR As Long, r As Long, k As Long, mat As String, eq As Double, ep As Double
    Dim cand As Long, nCand As Long, okQ As Boolean, okP As Boolean
    Dim mPR As Object, mItem As Object, nOK As Long, nMulti As Long, nNone As Long, noneList As String
    Set mPR = CreateObject("Scripting.Dictionary")
    Set mItem = CreateObject("Scripting.Dictionary")
    lastR = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    For r = 3 To lastR
        mat = UCase(Trim(CStr(ws.Range("C" & r).Value)))
        If mat <> "" And Trim(CStr(ws.Range("L" & r).Value)) = "" Then
            eq = 0: ep = 0
            If Not IsEmpty(CongNoNum(r, "F")) Then eq = CongNoNum(r, "F")
            If Not IsEmpty(CongNoNum(r, "G")) Then ep = CongNoNum(r, "G")
            cand = 0: nCand = 0
            For k = 1 To n
                If Not xUsed(k) And xMat(k) = mat Then
                    okQ = (Abs(ExportNum(xQ(k), eq) - eq) <= 0.0005)
                    okP = True
                    If cPrice > 0 Then okP = (Abs(ExportNum(xP(k), ep) - ep) <= 0.5)
                    If okQ And okP Then
                        nCand = nCand + 1
                        If cand = 0 Then
                            cand = k
                        ElseIf xPR(k) > xPR(cand) Then
                            cand = k                  ' nhieu PR khop -> lay PR moi nhat
                        End If
                    End If
                End If
            Next k
            If cand > 0 Then
                xUsed(cand) = True
                mPR(r) = xPR(cand)
                mItem(r) = xItem(cand)
                nOK = nOK + 1
                If nCand > 1 Then nMulti = nMulti + 1
            Else
                nNone = nNone + 1
                If nNone <= 10 Then noneList = noneList & "  - Dong " & r & " " & mat & " (" & eq & " x " & Format(ep, "#,##0") & ")" & vbCrLf
            End If
        End If
    Next r

    If nOK = 0 Then
        Hoi "Khong ghep duoc dong nao (" & n & " dong PR trong file xuat)." & vbCrLf & _
               IIf(noneList <> "", "Vi du dong khong tim thay:" & vbCrLf & noneList, "") & vbCrLf & _
               "Kiem tra bo loc ME5A (nguoi tao, ngay tao) va file xuat co cot Material / Quantity / Valuation Price.", _
               vbExclamation, "NapSoPRTuSAP"
        Exit Sub
    End If

    If Hoi("Ghep duoc " & nOK & " dong voi PR tren SAP" & IIf(nMulti > 0, " (" & nMulti & " dong co nhieu PR khop - da lay PR moi nhat, ghi chu o cot M)", "") & "." & vbCrLf & _
              IIf(nNone > 0, "Khong tim thay " & nNone & " dong:" & vbCrLf & noneList, "") & vbCrLf & _
              "Ghi So PR + so dong item vao sheet PR?", vbYesNo + vbQuestion, "NapSoPRTuSAP") = vbNo Then Exit Sub

    Dim key As Variant
    For Each key In mPR.Keys
        ws.Range("L" & key).NumberFormat = "@"
        ws.Range("L" & key).Value = mPR(key)
        ws.Range("P" & key).NumberFormat = "@"
        ws.Range("P" & key).Value = mItem(key)
        ws.Range("M" & key).Value = "Lay lai So PR tu file xuat ME5A (" & Format(Now, "dd/mm/yyyy hh:nn") & ")"
    Next key
    Hoi "Da ghi So PR cho " & nOK & " dong. Nho Ctrl+S luu file, roi chay TongHopPO.", vbInformation, "NapSoPRTuSAP"
End Sub

' Doi gia tri o file xuat SAP thanh so (so that, hoac chu kieu "1.234,5" / "1,234.5")
Function ExportNum(ByVal v As Variant, ByVal expected As Double) As Double
    If IsEmpty(v) Then ExportNum = -1: Exit Function
    If VarType(v) = vbString Then
        ExportNum = ParseBest(Trim(CStr(v)), expected)
    ElseIf IsNumeric(v) Then
        ExportNum = CDbl(v)
    Else
        ExportNum = -1
    End If
End Function


' --------------------------------------------------------------------------
' GAN SO PR BANG TAY: dung khi PR da tao tren SAP nhung cot L (SoPR) sheet PR
' dang trong (vd tao PR o file khac / chua luu file). Chon cac dong cua 1 PR
' tren sheet PR (theo DUNG thu tu mat hang trong PR), chay macro, nhap So PR.
' Macro ghi So PR vao cot L, so dong item 10, 20, 30... vao cot P.
' --------------------------------------------------------------------------
Sub GanSoPR()
    Dim ws As Worksheet, rng As Range, cell As Range, soPR As String
    Dim n As Long, stepAns As String, stp As Long, rowsDone As Object
    Set ws = ThisWorkbook.Sheets("PR")
    If Not ActiveSheet Is ws Then
        Hoi "Hay mo sheet PR, chon (boi den) cac dong thuoc 1 PR roi chay lai GanSoPR.", vbExclamation, "GanSoPR"
        Exit Sub
    End If
    EnsurePRHeaders ws
    Set rng = Intersect(Selection.EntireRow, ws.Range("A3:A" & ws.Rows.Count))
    If rng Is Nothing Then
        Hoi "Chua chon dong nao (tu dong 3 tro xuong).", vbExclamation, "GanSoPR"
        Exit Sub
    End If

    soPR = Trim(InputBox("Nhap So PR (lay tu SAP, vd ME53N) cho " & rng.Rows.Count & " dong dang chon:", "GanSoPR"))
    If soPR = "" Then Exit Sub
    If soPR Like "*[!0-9]*" Then
        Hoi "So PR chi gom chu so.", vbExclamation, "GanSoPR"
        Exit Sub
    End If
    stepAns = Trim(InputBox("Buoc nhay so dong item tren PR (thuong la 10 -> 10, 20, 30...):", "GanSoPR", CStr(PR_ITEM_STEP)))
    If stepAns = "" Or Not IsNumeric(stepAns) Then Exit Sub
    stp = CLng(stepAns)

    Set rowsDone = CreateObject("Scripting.Dictionary")
    For Each cell In rng.Cells
        If Trim(CStr(ws.Range("C" & cell.Row).Value)) <> "" And Not rowsDone.Exists(cell.Row) Then
            rowsDone.Add cell.Row, 1
            n = n + 1
            ws.Range("L" & cell.Row).NumberFormat = "@"
            ws.Range("L" & cell.Row).Value = soPR
            ws.Range("P" & cell.Row).NumberFormat = "@"
            ws.Range("P" & cell.Row).Value = CStr(n * stp)
            ws.Range("M" & cell.Row).Value = "Nhap tay So PR (GanSoPR) - dong item " & (n * stp)
        End If
    Next cell
    Hoi "Da gan So PR " & soPR & " cho " & n & " dong (item " & stp & " den " & (n * stp) & ")." & vbCrLf & _
           "Kiem tra lai thu tu mat hang khop voi PR tren SAP, roi chay TongHopPO / CreatePO.", vbInformation, "GanSoPR"
End Sub


' --------------------------------------------------------------------------
' TONG HOP PO: gom cac dong PR (da co SoPR) theo Nha_cung_cap.
' Chay bao nhieu lan cung duoc (giu lai cot G/K da co).
' --------------------------------------------------------------------------
Sub TongHopPO(Optional ByVal silent As Boolean = False)
    Dim wsPR As Worksheet, wsPO As Worksheet, wsM As Worksheet
    Set wsPR = ThisWorkbook.Sheets("PR")
    Set wsM = ThisWorkbook.Sheets("Mapping")
    Set wsPO = ThisWorkbook.Sheets("PO")
    Application.Calculate

    EnsurePRHeaders wsPR
    EnsureVendorTable wsM

    ' --- 1. Giu lai Tong_tien_tren_SAP + Status cua lan chay truoc ---
    Dim keep As Object, r As Long, lastPO As Long, kKey As String
    Set keep = CreateObject("Scripting.Dictionary")
    If wsPO.Range("G1").Value = "Tong_tien_tren_SAP" Then
        lastPO = wsPO.Cells(wsPO.Rows.Count, "B").End(xlUp).Row
        For r = 3 To lastPO
            If Trim(CStr(wsPO.Cells(r, "B").Value)) <> "" Then
                kKey = UCase(Trim(CStr(wsPO.Cells(r, "B").Value))) & "|" & Trim(CStr(wsPO.Cells(r, "J").Value))
                keep(kKey) = Array(wsPO.Cells(r, "G").Value, wsPO.Cells(r, "K").Value)
            End If
        Next r
    End If

    ' --- 2. Gom dong PR theo NCC (+ SoPO neu da tao) ---
    Dim grpOrder As Collection, grpLines As Object, grpTotal As Object, grpPR As Object
    Set grpOrder = New Collection
    Set grpLines = CreateObject("Scripting.Dictionary")
    Set grpTotal = CreateObject("Scripting.Dictionary")
    Set grpPR = CreateObject("Scripting.Dictionary")

    Dim lastPR As Long, sup As String, soPO As String, gk As String, soPRv As String
    lastPR = wsPR.Cells(wsPR.Rows.Count, "C").End(xlUp).Row
    For r = 3 To lastPR
        soPRv = Trim(CStr(wsPR.Range("L" & r).Value))
        If Trim(CStr(wsPR.Range("C" & r).Value)) <> "" And soPRv <> "" Then
            sup = Trim(CStr(wsPR.Range("N" & r).Value))
            If sup = "" Then sup = "(chua co NCC)"
            soPO = Trim(CStr(wsPR.Range("Q" & r).Value))
            gk = sup & "|" & soPO
            If Not grpLines.Exists(gk) Then
                grpOrder.Add gk
                grpLines.Add gk, 0
                grpTotal.Add gk, 0#
                grpPR.Add gk, ""
            End If
            grpLines(gk) = grpLines(gk) + 1
            If Not IsEmpty(CongNoNum(r, "F")) And Not IsEmpty(CongNoNum(r, "G")) Then
                grpTotal(gk) = grpTotal(gk) + CongNoNum(r, "F") * CongNoNum(r, "G")
            End If
            If InStr(", " & grpPR(gk) & ",", ", " & soPRv & ",") = 0 Then
                grpPR(gk) = IIf(grpPR(gk) = "", soPRv, grpPR(gk) & ", " & soPRv)
            End If
        End If
    Next r

    ' --- 3. Ve lai sheet PO ---
    wsPO.Cells.Clear
    wsPO.Range("A1:K1").Value = Array("STT", "Nha_cung_cap", "Ma_NCC_SAP", "So_dong", "So_PR", _
                                      "Tong_tien_theo_CongNo", "Tong_tien_tren_SAP", "Chenh_lech", "Kiem_tra", "SoPO", "Status")
    wsPO.Range("A2:K2").Value = Array("Tu dong", "Tu dong = PR cot N", "Tu dong = Mapping cot F", "Tu dong", "Tu dong", _
                                      "Tong SL x Don gia (CongNo)", "<< Macro doc lai tu PO tren SAP >>", "Bang G tru F", "Tu dong", _
                                      "<< Macro CreatePO >>", "<< Macro CreatePO >>")
    With wsPO.Range("A1:K1")
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
    End With
    wsPO.Range("A2:K2").Font.Italic = True
    wsPO.Range("A2:K2").Font.Color = RGB(118, 118, 118)

    Dim gItem As Variant, outR As Long, parts() As String, vcode As String, kv As Variant
    outR = 3
    For Each gItem In grpOrder
        parts = Split(gItem, "|")
        sup = parts(0)
        soPO = parts(1)
        vcode = ""
        If sup <> "(chua co NCC)" Then vcode = VendorCode(wsM, sup, True)

        wsPO.Cells(outR, "A").Value = outR - 2
        wsPO.Cells(outR, "B").Value = sup
        wsPO.Cells(outR, "C").NumberFormat = "@"
        wsPO.Cells(outR, "C").Value = vcode
        wsPO.Cells(outR, "D").Value = grpLines(gItem)
        wsPO.Cells(outR, "E").NumberFormat = "@"
        wsPO.Cells(outR, "E").Value = grpPR(gItem)
        wsPO.Cells(outR, "F").Value = grpTotal(gItem)
        wsPO.Cells(outR, "J").NumberFormat = "@"
        wsPO.Cells(outR, "J").Value = soPO

        kKey = UCase(sup) & "|" & soPO
        If keep.Exists(kKey) Then
            kv = keep(kKey)
            wsPO.Cells(outR, "G").Value = kv(0)
            wsPO.Cells(outR, "K").Value = kv(1)
        End If
        wsPO.Cells(outR, "H").Formula = "=IF(G" & outR & "="""","""",G" & outR & "-F" & outR & ")"
        wsPO.Cells(outR, "I").Formula = "=IF(C" & outR & "="""",""THIEU MA NCC"",IF(G" & outR & "="""",IF(J" & outR & "<>"""",""DA TAO PO"",""SAN SANG"")," & _
                                        "IF(ABS(G" & outR & "-F" & outR & ")<1,IF(J" & outR & "<>"""",""DA TAO PO - KHOP CONGNO"",""KHOP CONGNO""),""LECH CONGNO"")))"
        outR = outR + 1
    Next gItem

    wsPO.Range("F3:H" & Application.Max(3, outR)).NumberFormat = "#,##0"
    wsPO.Columns("A:K").AutoFit

    Dim fc As FormatCondition
    With wsPO.Range("I3:I1000")
        .FormatConditions.Delete
        Set fc = .FormatConditions.Add(Type:=xlTextString, String:="KHOP CONGNO", TextOperator:=xlContains)
        fc.Interior.Color = RGB(198, 239, 206)
        Set fc = .FormatConditions.Add(Type:=xlTextString, String:="LECH", TextOperator:=xlContains)
        fc.Interior.Color = RGB(255, 199, 206)
        Set fc = .FormatConditions.Add(Type:=xlTextString, String:="THIEU", TextOperator:=xlContains)
        fc.Interior.Color = RGB(255, 199, 206)
    End With
    Application.Calculate

    If Not silent And outR = 3 Then
        wsPO.Activate
        Hoi "Chua co dong PR nao co So PR (cot L sheet PR) ma chua tao PO." & vbCrLf & vbCrLf & _
               "Macro chi tao PO cho cac dong da co So PR. Neu PR da tao tren SAP nhung cot L dang trong," & vbCrLf & _
               "hay chon cac dong do o sheet PR roi chay macro GanSoPR de nhap So PR.", vbExclamation, "TongHopPO " & MACRO_VERSION
        Exit Sub
    End If
    If Not silent Then
        wsPO.Activate
        Hoi "Da tong hop " & (outR - 3) & " dong PO (moi nha cung cap 1 PO) tu cac dong PR da co SoPR." & vbCrLf & vbCrLf & _
               "Neu cot Kiem_tra bao THIEU MA NCC: chay NapDuLieuCongNo hoac dien ma o Mapping cot F." & vbCrLf & _
               "Tong tien tren SAP (cot G) se duoc macro CreatePO doc lai sau khi nhap PO.", vbInformation, "TongHopPO " & MACRO_VERSION
    End If
End Sub


' --------------------------------------------------------------------------
' SO SANH PO tren SAP voi sheet CongNo (nguon gia/SL duy nhat).
' prLines  : cac so dong sheet PR thuoc PO nay (dong PR r <-> dong CongNo r)
' sapLines : Collection, moi phan tu = Array(ma_hang, so_luong, don_gia) doc tu SAP
' Tra ve bao cao (chuoi), dong thoi tra sapTotal (tong tien tren SAP) va nDiff (so cho lech)
' Ghep dong theo MA HANG (cung ma xuat hien nhieu lan thi ghep lan luot).
' --------------------------------------------------------------------------
Function ComparePOWithCongNo(wsPR As Worksheet, prLines As Collection, sapLines As Collection, _
                             ByRef sapTotal As Double, ByRef nDiff As Long) As String
    Dim used() As Boolean, k As Long, rep As String, wsC As Worksheet
    Dim mat As String, q As Double, p As Double, found As Boolean, s As Variant
    Set wsC = ThisWorkbook.Sheets("CongNo")
    ReDim used(1 To Application.Max(1, sapLines.Count))
    sapTotal = 0: nDiff = 0

    For k = 1 To sapLines.Count
        s = sapLines(k)
        sapTotal = sapTotal + CDbl(s(1)) * CDbl(s(2))
    Next k

    Dim rr As Variant
    For Each rr In prLines
        mat = UCase(Trim(CStr(wsC.Range("D" & CNRow(rr)).Value)))
        If IsEmpty(CongNoNum(rr, "F")) Or IsEmpty(CongNoNum(rr, "G")) Then
            nDiff = nDiff + 1
            rep = rep & "  - " & mat & ": SL/gia o sheet CongNo dong " & CNRow(rr) & " khong hop le" & vbCrLf
            GoTo NextLine
        End If
        q = CongNoNum(rr, "F")
        p = CongNoNum(rr, "G")

        ' CongNo bi sua sau khi da tao PR?
        If Trim(CStr(wsPR.Range("R" & rr).Value)) <> "" Then
            If Abs(CDbl(wsPR.Range("R" & rr).Value) - p) > 0.5 Or Abs(CDbl(wsPR.Range("S" & rr).Value) - q) > 0.0005 Then
                nDiff = nDiff + 1
                rep = rep & "  - " & mat & ": CongNo (" & q & " x " & Format(p, "#,##0") & ") KHAC luc tao PR (" & _
                      wsPR.Range("S" & rr).Value & " x " & Format(wsPR.Range("R" & rr).Value, "#,##0") & ")" & vbCrLf
            End If
        End If

        found = False
        For k = 1 To sapLines.Count
            s = sapLines(k)
            If Not used(k) And UCase(Trim(CStr(s(0)))) = mat Then
                used(k) = True
                found = True
                If Abs(CDbl(s(1)) - q) > 0.0005 Or Abs(CDbl(s(2)) - p) > 0.5 Then
                    nDiff = nDiff + 1
                    rep = rep & "  - " & mat & ": CongNo " & q & " x " & Format(p, "#,##0") & _
                          " / PO tren SAP " & s(1) & " x " & Format(s(2), "#,##0") & vbCrLf
                End If
                Exit For
            End If
        Next k
        If Not found Then
            nDiff = nDiff + 1
            rep = rep & "  - " & mat & ": co trong CongNo/PR nhung KHONG thay tren PO" & vbCrLf
        End If
NextLine:
    Next rr

    For k = 1 To sapLines.Count
        If Not used(k) Then
            nDiff = nDiff + 1
            s = sapLines(k)
            rep = rep & "  - " & s(0) & ": co tren PO nhung KHONG co trong danh sach" & vbCrLf
        End If
    Next k
    ComparePOWithCongNo = rep
End Function


' --------------------------------------------------------------------------
' TAO PO tren SAP (ME21N): 1 PO / nha cung cap
' --------------------------------------------------------------------------
Sub CreatePO()
    TongHopPO True

    Dim wsPO As Worksheet, wsPR As Worksheet
    Set wsPO = ThisWorkbook.Sheets("PO")
    Set wsPR = ThisWorkbook.Sheets("PR")

    Dim lastPO As Long, r As Long
    Dim todo As New Collection, badList As String, summary As String, v As Variant
    lastPO = wsPO.Cells(wsPO.Rows.Count, "B").End(xlUp).Row

    For r = 3 To lastPO
        If Trim(CStr(wsPO.Cells(r, "J").Value)) = "" And Trim(CStr(wsPO.Cells(r, "B").Value)) <> "" Then
            If Trim(CStr(wsPO.Cells(r, "C").Value)) = "" Then
                badList = badList & "  - " & wsPO.Cells(r, "B").Value & ": thieu Ma NCC SAP" & vbCrLf
                wsPO.Cells(r, "K").Value = "CHUA TAO - thieu Ma NCC SAP"
            Else
                todo.Add r
            End If
        End If
    Next r

    If todo.Count = 0 Then
        Hoi "Khong co nha cung cap nao san sang tao PO." & vbCrLf & vbCrLf & _
               IIf(badList <> "", "Can xu ly:" & vbCrLf & badList & vbCrLf, "") & _
               "Chi cac dong sheet PR DA CO So PR (cot L) va chua co So PO moi duoc tao PO." & vbCrLf & _
               "Neu PR da tao tren SAP nhung cot L trong: chon cac dong do roi chay macro GanSoPR.", vbExclamation, "CreatePO " & MACRO_VERSION
        wsPO.Activate
        Exit Sub
    End If

    For Each v In todo
        summary = summary & "  - " & wsPO.Cells(v, "B").Value & " (ma " & wsPO.Cells(v, "C").Value & "): " & _
                  wsPO.Cells(v, "D").Value & " dong, " & Format(wsPO.Cells(v, "F").Value, "#,##0") & vbCrLf
    Next v
    If badList <> "" Then summary = summary & vbCrLf & "Bo qua (can xu ly):" & vbCrLf & badList

    If Not PO_SCRIPT_READY Then
        wsPO.Activate
        Hoi "Da tong hop xong. Se tao " & todo.Count & " PO:" & vbCrLf & summary & vbCrLf & _
               "PHAN NHAP VAO SAP (ME21N) CHUA CO: can file record thao tac tao 1 PO tu PR da duyet." & vbCrLf & _
               "Gui file .vbs do de hoan thien macro.", vbInformation, "CreatePO " & MACRO_VERSION
        Exit Sub
    End If

    If Hoi("Se tao " & todo.Count & " PO:" & vbCrLf & summary & vbCrLf & "Bat dau nhap vao SAP?", _
              vbYesNo + vbQuestion, "CreatePO " & MACRO_VERSION) = vbNo Then Exit Sub

    Dim session As Object
    Set session = GetSapSession()
    If session Is Nothing Then Exit Sub

    Dim lines As Collection, i As Long, lastPR As Long, sup As String, newPO As String, rr As Variant
    Dim sapLines As Collection, cmpRep As String, sapTotal As Double, nDiff As Long, ans As VbMsgBoxResult
    lastPR = wsPR.Cells(wsPR.Rows.Count, "C").End(xlUp).Row

    For Each v In todo
        sup = Trim(CStr(wsPO.Cells(v, "B").Value))
        Set lines = New Collection
        For i = 3 To lastPR
            If UCase(Trim(CStr(wsPR.Range("N" & i).Value))) = UCase(sup) And _
               Trim(CStr(wsPR.Range("L" & i).Value)) <> "" And Trim(CStr(wsPR.Range("Q" & i).Value)) = "" Then
                lines.Add i
            End If
        Next i

        On Error GoTo POError

        ' (a) Nhap PO tren SAP (chua Save), doc lai cac dong PO
        Set sapLines = EnterPOInSAP(session, wsPR, lines, CStr(wsPO.Cells(v, "C").Value))

        ' (b) So voi danh sach PR
        curStep = "so sanh PO voi CongNo"
        cmpRep = ComparePOWithCongNo(wsPR, lines, sapLines, sapTotal, nDiff)
        wsPO.Cells(v, "G").Value = sapTotal

        ' (c) Dung lai xac nhan truoc khi Save
        On Error Resume Next
        TramActivate
        On Error GoTo POError
        If gPOWarn <> "" Then cmpRep = cmpRep & vbCrLf & "Canh bao khi nhap gia:" & vbCrLf & gPOWarn
        If nDiff = 0 Then
            ans = Hoi("PO cho " & sup & ": " & lines.Count & " dong, tong " & Format(sapTotal, "#,##0") & vbCrLf & _
                         IIf(gPOWarn <> "", "Canh bao khi nhap gia:" & vbCrLf & gPOWarn & vbCrLf, "") & _
                         "Gia + SL BANG TUYET DOI sheet CongNo (ma hang, so luong, don gia)." & vbCrLf & vbCrLf & _
                         "Luu (Save) PO nay?", vbYesNo + vbQuestion, "CreatePO " & MACRO_VERSION)
        Else
            ans = Hoi("PO cho " & sup & " LECH voi sheet CongNo o " & nDiff & " cho:" & vbCrLf & cmpRep & vbCrLf & _
                         "Tong theo CongNo: " & Format(wsPO.Cells(v, "F").Value, "#,##0") & " / tren SAP: " & Format(sapTotal, "#,##0") & vbCrLf & vbCrLf & _
                         "Van Luu (Save) PO nay? (No = dung lai de ban tu kiem tra tren SAP)", _
                         vbYesNo + vbExclamation + vbDefaultButton2, "CreatePO " & MACRO_VERSION)
        End If
        If ans = vbNo Then
            wsPO.Cells(v, "K").Value = "CHUA SAVE - dung theo yeu cau" & IIf(nDiff > 0, " (lech " & nDiff & " cho)", "")
            Exit Sub
        End If

        ' (d) Save - v23: tu xu ly hop thoai canh bao SAP hien ra khi Save (bam Save/Yes de luu de)
        curStep = "bam Save PO"
        newPO = SavePOWithPopups(session, sup)
        On Error GoTo 0

        wsPO.Cells(v, "J").NumberFormat = "@"
        wsPO.Cells(v, "J").Value = newPO
        wsPO.Cells(v, "K").Value = session.findById("wnd[0]/sbar").Text
        If newPO <> "" Then
            For Each rr In lines
                wsPR.Range("Q" & rr).NumberFormat = "@"
                wsPR.Range("Q" & rr).Value = newPO
            Next rr
        End If
        GoTo NextPO
POError:
        wsPO.Cells(v, "K").Value = "LOI, buoc [" & curStep & "]: " & Err.Description
        Hoi "Loi khi tao PO cho " & sup & ":" & vbCrLf & Err.Description & vbCrLf & _
               "Buoc: " & curStep & vbCrLf & _
               "Macro da DUNG, man hinh SAP giu nguyen.", vbCritical, "CreatePO " & MACRO_VERSION
        Exit Sub
NextPO:
    Next v

    TongHopPO True
    wsPO.Activate
    Hoi "Da xu ly xong tao PO. Kiem tra cot SoPO / Status sheet PO.", vbInformation, "CreatePO " & MACRO_VERSION
End Sub

' --------------------------------------------------------------------------
' NHAP 1 PO tren ME21N (CHUA Save) - theo file record cua ban:
'   1. /nme21n, Loai PO = ZEXP, go Ma NCC vao o NCC, Enter
'   2. Them cac dong PR vao PO: go SO PR + SO DONG ITEM PR vao bang item
'      (cot "Purch.Req." / "Item of requisition") -> SAP tu lay ma hang, SL, gia.
'      (Khong dung cay Document Overview ben trai vi so thu tu node thay doi moi lan.)
'   3. Doc lai tung dong PO: SL khac CongNo -> go lai o Quantity.
'      GIA khac CongNo -> CACH 1: go thang vao o Net Price; neu loi / khong an
'      -> CACH 2: nhap dieu kien PBXX o tab Conditions (theo file record).
'   4. Tra ve Collection cac Array(ma_hang, so_luong, don_gia) doc tu SAP de so sanh.
' --------------------------------------------------------------------------
Function EnterPOInSAP(session As Object, wsPR As Worksheet, lines As Collection, ByVal vendor As String) As Collection
    Dim TBL As Object, cmb As Object, fld As Object
    Dim cBanfn As Long, cBnfpo As Long, cMat As Long, cQty As Long, cPrice As Long
    Dim i As Long, r As Long, soPR As String, prItem As String, prevPR As String, posInPR As Long
    Dim cel As Object

    gPOWarn = ""

    ' 1. Mo ME21N
    curStep = "mo ME21N"
    SapCuaSo session
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme21n"
    session.findById("wnd[0]").sendVKey 0
    CheckPopupPO session
    If UCase(session.Info.Transaction) <> "ME21N" Then
        Err.Raise vbObjectError + 601, "EnterPOInSAP", "SAP khong mo duoc ME21N (dang o '" & session.Info.Transaction & "')"
    End If

    curStep = "chon Loai PO (" & PO_DOC_TYPE & ")"
    Set cmb = FindCtl(session, "cmbMEPO_TOPLINE-BSART")
    If cmb.key <> PO_DOC_TYPE Then cmb.key = PO_DOC_TYPE
    CheckPopupPO session

    curStep = "nhap Ma NCC " & vendor
    Set fld = FindCtl(session, "ctxtMEPO_TOPLINE-SUPERFIELD")
    fld.Text = vendor
    fld.SetFocus
    session.findById("wnd[0]").sendVKey 0
    CheckPopupPO session

    ' Thu gon Item Detail de bang item hien duoc nhieu dong hon (it phai cuon)
    On Error Resume Next
    SetItemDetail session, False
    Err.Clear
    On Error GoTo 0

    ' 2. Xac dinh vi tri cac cot trong bang item
    curStep = "tim cac cot bang item PO"
    Set TBL = FindCtl(session, PO_TBL_ID)
    cBanfn = TblColIndex(TBL, "MEPO1211-BANFN")
    cBnfpo = TblColIndex(TBL, "MEPO1211-BNFPO")
    cMat = TblColIndex(TBL, "MEPO1211-EMATN")
    cQty = TblColIndex(TBL, "MEPO1211-MENGE")
    cPrice = TblColIndex(TBL, "MEPO1211-NETPR")
    If cBanfn < 0 Or cBnfpo < 0 Then
        Err.Raise vbObjectError + 602, "EnterPOInSAP", _
            "Bang item PO khong co cot 'Purch.Req.' / 'Item of requisition' (BANFN/BNFPO). " & _
            "Hay hien 2 cot nay trong layout bang item ME21N roi chay lai."
    End If
    If cMat < 0 Or cQty < 0 Or cPrice < 0 Then
        Err.Raise vbObjectError + 603, "EnterPOInSAP", "Bang item PO thieu cot Material / PO Quantity / Net Price"
    End If

    ' 3. Go So PR + dong item PR cho tung dong
    prevPR = "": posInPR = 0
    For i = 0 To lines.Count - 1
        r = lines(i + 1)
        soPR = Trim(CStr(wsPR.Range("L" & r).Value))
        If soPR <> prevPR Then posInPR = 0
        posInPR = posInPR + 1
        prevPR = soPR
        prItem = Trim(CStr(wsPR.Range("P" & r).Value))
        If prItem = "" Then prItem = CStr(PrItemGuess(wsPR, r))   ' suy ra theo thu tu dong trong PR

        ' v23: khong bao gio go SL/gia 0 vao SAP
        If IsEmpty(CongNoNum(r, "F")) Or IsEmpty(CongNoNum(r, "G")) Then
            Err.Raise vbObjectError + 610, "EnterPOInSAP", "Sheet CongNo dong " & CNRow(r) & " (" & wsPR.Range("C" & r).Value & _
                ") khong co SL/don gia hop le. Kiem tra sheet CongNo (co bi chen/xoa dong khong?)"
        End If
        If CDbl(CongNoNum(r, "F")) <= 0 Or CDbl(CongNoNum(r, "G")) <= 0 Then
            Err.Raise vbObjectError + 611, "EnterPOInSAP", "Sheet CongNo dong " & CNRow(r) & " (" & wsPR.Range("C" & r).Value & _
                "): SL hoac don gia = 0. Kiem tra sheet CongNo truoc khi tao PO"
        End If

        curStep = "them dong PR " & soPR & "/" & prItem & " (" & wsPR.Range("C" & r).Value & ")"
        Set cel = PORowCell(session, i, cBanfn)
        cel.Text = soPR
        Set cel = PORowCell(session, i, cBnfpo)
        cel.Text = prItem
        session.findById("wnd[0]").sendVKey 0
        CheckPopupPO session

        ' SAP co the bao loi ngay khi lay dong PR sang (vd "Gia tri item PO phai <= gia tri
        ' item PR"), do gia mac dinh tren PO (tu info record...) khac gia PR.
        ' -> go ngay SL + gia theo CongNo (= gia PR) vao dong do roi Enter lai.
        If SbarIsError(session) Then
            Dim adoptErr As String
            adoptErr = session.findById("wnd[0]/sbar").Text
            curStep = "chinh SL/gia theo CongNo ngay khi them dong PR " & soPR & "/" & prItem & _
                      " (SAP bao: " & adoptErr & ")"
            On Error Resume Next
            Set cel = PORowCell(session, i, cQty)
            If cel.Changeable Then cel.Text = SapNumber(CDbl(CongNoNum(r, "F")))
            Err.Clear
            On Error GoTo 0
            Dim adoptNetErr As String
            If Not TrySetNetPrice(session, i, cPrice, CDbl(CongNoNum(r, "G")), "", adoptNetErr) Then
                ' o Net Price khong sua duoc (hang kho...) -> nhap qua dieu kien PBXX, KHONG dung chu trinh
                Dim adoptPb As String
                adoptPb = TryPBXX(session, i, CDbl(CongNoNum(r, "G")))
                If adoptPb <> "" Then
                    gPOWarn = gPOWarn & "  - dong " & (i + 1) & " " & wsPR.Range("C" & r).Value & ": " & adoptPb & vbCrLf
                End If
            End If
        End If
    Next i

    ' 4. Doc lai tung dong, chinh SL/gia cho BANG CongNo, tra ve ket qua
    Dim res As New Collection, mat As String, expQ As Double, expP As Double
    Dim dq As String, dp As String, q As Double, pr As Double, fixedSomething As Boolean
    For i = 0 To lines.Count - 1
        r = lines(i + 1)
        expQ = CongNoNum(r, "F")
        expP = CongNoNum(r, "G")
        curStep = "doc lai dong PO " & (i + 1) & " (" & wsPR.Range("C" & r).Value & ")"

        ReadPORow session, i, cQty, cPrice, expQ, expP, dq, dp, q, pr

        ' 4a. So luong khac CongNo -> go lai o Quantity
        If Abs(q - expQ) > 0.0005 Then
            curStep = "chinh SL dong PO " & (i + 1) & " theo CongNo"
            Set cel = PORowCell(session, i, cQty)
            cel.Text = SapNumber(expQ)
            session.findById("wnd[0]").sendVKey 0
            CheckPopupPO session
            CheckSbarError session
            ReadPORow session, i, cQty, cPrice, expQ, expP, dq, dp, q, pr
        End If

        ' 4b. Gia khac CongNo -> CACH 1: go thang vao o Net Price
        Dim netOK As Boolean, netErr As String
        If Abs(pr - expP) > 0.5 Then
            curStep = "go gia vao Net Price dong PO " & (i + 1)
            netOK = TrySetNetPrice(session, i, cPrice, expP, dp, netErr)
            ReadPORow session, i, cQty, cPrice, expQ, expP, dq, dp, q, pr
            If Not netOK Or Abs(pr - expP) > 0.5 Then
                ' CACH 2: Net Price khong duoc -> nhap dieu kien PBXX o tab Conditions
                curStep = "nhap gia qua dieu kien PBXX dong PO " & (i + 1) & _
                          IIf(netErr <> "", " (Net Price loi: " & netErr & ")", "")
                Dim pbErr As String
                pbErr = TryPBXX(session, i, expP)
                ReadPORow session, i, cQty, cPrice, expQ, expP, dq, dp, q, pr
                If pbErr <> "" Or Abs(pr - expP) > 0.5 Then
                    ' v23: KHONG dung chu trinh. Ghi lai, dong nay se hien LECH o hop thoai truoc khi Save
                    gPOWarn = gPOWarn & "  - dong " & (i + 1) & " " & wsPR.Range("C" & r).Value & ": gia SAP " & _
                              Format(pr, "#,##0") & " / CongNo " & Format(expP, "#,##0") & _
                              IIf(pbErr <> "", " (PBXX: " & pbErr & ")", "") & vbCrLf
                End If
            End If
        End If

        Set cel = PORowCell(session, i, cMat)
        mat = Trim(cel.Text)
        res.Add Array(mat, q, pr)
    Next i

    On Error Resume Next
    Set TBL = FindCtl(session, PO_TBL_ID)
    TBL.VerticalScrollbar.Position = 0
    Err.Clear
    On Error GoTo 0
    Set EnterPOInSAP = res
End Function

' Lay o (dong idx cua PO, cot col) tren bang item. Bang chi hien vai dong mot
' luc nen chi cuon khi dong can lay nam ngoai vung dang hien, va khong cuon
' qua gioi han cua thanh cuon (tranh loi "The method got an invalid argument").
Function PORowCell(session As Object, ByVal idx As Long, ByVal col As Long) As Object
    Dim TBL As Object, vis As Long, pos As Long, mx As Long, newPos As Long, tries As Long, lastPos As Long
    Set TBL = FindCtl(session, PO_TBL_ID)
    vis = TBL.VisibleRowCount
    pos = TBL.VerticalScrollbar.Position

    For tries = 1 To 6
        If idx >= pos And idx <= pos + vis - 1 Then Exit For
        lastPos = pos
        mx = TBL.VerticalScrollbar.Maximum
        If idx < pos Then
            newPos = idx
        Else
            newPos = idx - vis + 1          ' dua dong can lay xuong hang CUOI vung dang hien
        End If
        If newPos > mx Then newPos = mx
        If newPos < 0 Then newPos = 0

        Select Case tries
            Case 1, 2
                ' Cach 1: dat vi tri thanh cuon
                TBL.VerticalScrollbar.Position = newPos
            Case 3, 4
                ' Cach 2: dat con tro vao hang cuoi dang hien roi bam Page Down / Page Up
                On Error Resume Next
                TBL.GetCell(IIf(idx < pos, 0, vis - 1) + 0, col + 0).SetFocus
                On Error GoTo 0
                session.findById("wnd[0]").sendVKey IIf(idx < pos, 81, 82)
            Case 5
                ' Cach 3: thu gon phan Item Detail de bang item hien nhieu dong hon
                SetItemDetail session, False
        End Select
        WaitSap session
        Set TBL = FindCtl(session, PO_TBL_ID)
        vis = TBL.VisibleRowCount
        pos = TBL.VerticalScrollbar.Position
    Next tries

    If idx - pos < 0 Or idx - pos > vis - 1 Then
        Err.Raise vbObjectError + 607, "PORowCell", "Khong cuon bang item toi dong " & (idx + 1) & _
            " (dang hien dong " & (pos + 1) & "-" & (pos + vis) & ", toi da cuon " & TBL.VerticalScrollbar.Maximum & ")"
    End If
    Set PORowCell = TBL.GetCell(idx - pos + 0, col + 0)
End Function

' Tim nut thu gon / mo rong phan Item Detail (nut DYN_4000-BUTTON trong khu vuc SUB3)
Function DetailToggleBtn(session As Object) As Object
    Dim n As Long, d As Long, b As Object
    On Error Resume Next
    For n = 10 To 30
        For d = 4000 To 4003
            Set b = Nothing
            Set b = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Format(n, "0000") & _
                                     "/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:" & d & "/btnDYN_4000-BUTTON")
            Err.Clear
            If Not b Is Nothing Then
                Set DetailToggleBtn = b
                Exit Function
            End If
        Next d
    Next n
End Function

' Mo rong (expand=True) hoac thu gon (False) phan Item Detail cua ME21N
Sub SetItemDetail(session As Object, ByVal expand As Boolean)
    Dim isOpen As Boolean, b As Object
    isOpen = Not (FindCtlSoft(session, "cmbDYN_6000-LIST") Is Nothing)
    If isOpen = expand Then Exit Sub
    Set b = DetailToggleBtn(session)
    If b Is Nothing Then Exit Sub
    b.press
    WaitSap session
    CheckPopupPO session
End Sub

' Doc lai SL + gia cua dong PO thu idx tren bang item
Sub ReadPORow(session As Object, ByVal idx As Long, ByVal cQty As Long, ByVal cPrice As Long, _
              ByVal expQ As Double, ByVal expP As Double, _
              ByRef dq As String, ByRef dp As String, ByRef q As Double, ByRef pr As Double)
    dq = Trim(PORowCell(session, idx, cQty).Text)
    dp = Trim(PORowCell(session, idx, cPrice).Text)
    q = ParseBest(dq, expQ)
    pr = ParseBest(dp, expP)
End Sub

' CACH 1: go gia thang vao o Net Price cua dong idx.
' Tra ve False (kem ly do o errText) neu o khong sua duoc hoac SAP bao loi;
' khi SAP bao loi thi tra lai gia cu de man hinh khong bi ket loi.
Function TrySetNetPrice(session As Object, ByVal idx As Long, ByVal cPrice As Long, ByVal price As Double, _
                        ByVal oldText As String, ByRef errText As String) As Boolean
    Dim cel As Object, mt As String
    errText = ""
    Set cel = PORowCell(session, idx, cPrice)

    On Error Resume Next
    If Not cel.Changeable Then
        If Err.Number = 0 Then
            errText = "o Net Price bi khoa (gia lay tu dieu kien)"
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
    End If
    If oldText = "" Then oldText = cel.Text
    cel.Text = SapNumber(price)
    If Err.Number <> 0 Then
        errText = "o Net Price khong sua duoc (" & Err.Description & ")"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    session.findById("wnd[0]").sendVKey 0
    CheckPopupPO session
    mt = session.findById("wnd[0]/sbar").MessageType
    If mt = "E" Or mt = "A" Then
        errText = session.findById("wnd[0]/sbar").Text
        ' tra lai gia cu de thoat trang thai loi
        On Error Resume Next
        Set cel = PORowCell(session, idx, cPrice)
        cel.Text = oldText
        session.findById("wnd[0]").sendVKey 0
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    TrySetNetPrice = True
End Function

' CACH 2: nhap gia qua dieu kien PBXX o tab Conditions cua dong idx.
' v23 (sua loi dung chu trinh o hang kho):
'   - Quet CA bang Conditions (co cuon), khong chi vai dong dang hien
'   - Bo qua hop thoai thong bao (chi co nut Tiep tuc) thay vi dung lai
'   - Neu SAP khong nhan PBXX hoac o PBXX bi khoa -> sua dong gia co san
'     (PB00 / PR00 / P000) neu dong do cho sua
'   - Khong bao gio go gia 0
Sub SetPriceByPBXX(session As Object, ByVal idx As Long, ByVal price As Double)
    Dim tabC As Object, ct As Object, cK As Long, cA As Long
    Dim rowPB As Long, rowEmpty As Long, rowAlt As Long, cel As Object, st As String, ok As Boolean

    If price <= 0 Then
        Err.Raise vbObjectError + 609, "SetPriceByPBXX", "Don gia CongNo = 0 -> khong nhap PBXX (kiem tra sheet CongNo)"
    End If

    SetItemDetail session, True        ' can mo Item Detail de vao tab Conditions
    SelectDetailItem session, idx
    CheckPopupPO session

    Set tabC = FindCtl(session, "tabpTABIDT9")
    SapSelect tabC
    WaitSap session
    CheckPopupPO session

    Set ct = FindCtl(session, COND_TBL_ID)
    cK = TblColIndex(ct, "KOMV-KSCHL")
    cA = TblColIndex(ct, "KOMV-KBETR")
    If cK < 0 Or cA < 0 Then
        Err.Raise vbObjectError + 605, "SetPriceByPBXX", "Bang Conditions khong co cot Condition type / Amount (KSCHL/KBETR)"
    End If

    FindCondRows session, cK, cA, rowPB, rowEmpty, rowAlt

    ' Chua co dong PBXX -> go PBXX vao dong trong dau tien
    If rowPB < 0 And rowEmpty >= 0 Then
        Set cel = CondCell(session, rowEmpty, cK)
        cel.Text = "PBXX"
        session.findById("wnd[0]").sendVKey 0
        CheckPopupPO session
        If SbarIsError(session) Then
            ' SAP khong cho dung PBXX cho dong nay -> xoa o vua go, dung dong gia co san
            st = SbarText(session)
            On Error Resume Next
            Set cel = CondCell(session, rowEmpty, cK)
            cel.Text = ""
            session.findById("wnd[0]").sendVKey 0
            WaitSap session
            Err.Clear
            On Error GoTo 0
            CheckPopupPO session
            If rowAlt < 0 Then
                Err.Raise vbObjectError + 607, "SetPriceByPBXX", "SAP khong nhan PBXX (" & st & ") va khong co dong gia nao sua duoc"
            End If
            rowPB = rowAlt
        Else
            FindCondRows session, cK, cA, rowPB, rowEmpty, rowAlt
        End If
    End If

    If rowPB < 0 Then
        If rowAlt < 0 Then
            Err.Raise vbObjectError + 606, "SetPriceByPBXX", "Bang Conditions khong co dong PBXX va khong them duoc (het dong trong)"
        End If
        rowPB = rowAlt
    End If

    ' O Amount cua dong PBXX bi khoa -> thu dong gia khac
    Set cel = CondCell(session, rowPB, cA)
    ok = False
    On Error Resume Next
    ok = cel.Changeable
    Err.Clear
    On Error GoTo 0
    If Not ok Then
        If rowAlt >= 0 And rowAlt <> rowPB Then
            rowPB = rowAlt
            Set cel = CondCell(session, rowPB, cA)
        Else
            Err.Raise vbObjectError + 608, "SetPriceByPBXX", "O so tien cua dieu kien gia bi khoa (khong sua duoc)"
        End If
    End If

    cel.Text = SapNumber(price)
    session.findById("wnd[0]").sendVKey 0
    CheckPopupPO session
    If SbarIsError(session) Then
        Err.Raise vbObjectError + 604, "SetPriceByPBXX", "SAP bao loi khi nhap gia o Conditions: " & SbarText(session)
    End If
End Sub

' Goi SetPriceByPBXX nhung KHONG lam dung chu trinh: tra ve "" neu thanh cong,
' nguoc lai tra ve ly do (de ghi canh bao va chay tiep).
Function TryPBXX(session As Object, ByVal idx As Long, ByVal price As Double) As String
    Dim saveStep As String
    saveStep = curStep
    On Error GoTo EH
    SetPriceByPBXX session, idx, price
    curStep = saveStep
    Exit Function
EH:
    TryPBXX = Err.Description
    Resume DonDep
DonDep:
    curStep = saveStep
    On Error Resume Next
    ClosePopupCancel session
    SetItemDetail session, False
    Err.Clear
End Function

' Tim tren bang Conditions (quet het, co cuon): dong PBXX, dong trong dau tien
' (o Condition type cho go), va 1 dong gia co san cho sua (PB00/PR00/P000/PBXX).
' So dong tra ve la so dong TUYET DOI (tinh tu dau bang).
Sub FindCondRows(session As Object, ByVal cK As Long, ByVal cA As Long, _
                 ByRef rowPB As Long, ByRef rowEmpty As Long, ByRef rowAlt As Long)
    Dim ct As Object, total As Long, vis As Long, pos As Long, lastPos As Long, mx As Long
    Dim rr As Long, t As String, chg As Boolean, guard As Long
    rowPB = -1: rowEmpty = -1: rowAlt = -1
    Set ct = FindCtl(session, COND_TBL_ID)
    On Error Resume Next
    vis = ct.VisibleRowCount
    total = ct.RowCount
    mx = ct.VerticalScrollbar.Maximum
    If total <= 0 Then total = vis
    If ct.VerticalScrollbar.Position <> 0 Then
        ct.VerticalScrollbar.Position = 0
        WaitSap session
    End If
    Err.Clear
    lastPos = -1
    For guard = 1 To 30
        Set ct = Nothing
        Set ct = FindCtlSoft(session, COND_TBL_ID)
        If ct Is Nothing Then Exit For
        pos = ct.VerticalScrollbar.Position
        If pos = lastPos Then Exit For
        lastPos = pos
        For rr = 0 To vis - 1
            If pos + rr >= total And total > vis Then Exit For
            t = ""
            t = UCase(Trim(ct.GetCell(rr + 0, cK + 0).Text))
            If t = "PBXX" And rowPB < 0 Then rowPB = pos + rr
            If t = "" And rowEmpty < 0 Then
                chg = False
                chg = ct.GetCell(rr + 0, cK + 0).Changeable
                If chg Then rowEmpty = pos + rr
            End If
            If rowAlt < 0 And (t = "PB00" Or t = "PR00" Or t = "P000" Or t = "PBXX") Then
                chg = False
                chg = ct.GetCell(rr + 0, cA + 0).Changeable
                If chg Then rowAlt = pos + rr
            End If
        Next rr
        If rowPB >= 0 And rowEmpty >= 0 Then Exit For
        If pos + vis >= total Or pos >= mx Then Exit For
        ct.VerticalScrollbar.Position = IIf(pos + vis > mx, mx, pos + vis)
        WaitSap session
    Next guard
    Err.Clear
    On Error GoTo 0
End Sub

' O (dong tuyet doi absRow, cot col) cua bang Conditions - tu cuon neu can
Function CondCell(session As Object, ByVal absRow As Long, ByVal col As Long) As Object
    Dim ct As Object, pos As Long, vis As Long, mx As Long, np As Long
    Set ct = FindCtl(session, COND_TBL_ID)
    vis = ct.VisibleRowCount
    pos = ct.VerticalScrollbar.Position
    If absRow < pos Or absRow > pos + vis - 1 Then
        mx = ct.VerticalScrollbar.Maximum
        np = absRow
        If np > mx Then np = mx
        If np < 0 Then np = 0
        ct.VerticalScrollbar.Position = np
        WaitSap session
        Set ct = FindCtl(session, COND_TBL_ID)
        pos = ct.VerticalScrollbar.Position
    End If
    If absRow - pos < 0 Or absRow - pos > vis - 1 Then
        Err.Raise vbObjectError + 612, "CondCell", "Khong cuon duoc bang Conditions toi dong " & (absRow + 1)
    End If
    Set CondCell = ct.GetCell(absRow - pos + 0, col + 0)
End Function

Function SbarText(session As Object) As String
    On Error Resume Next
    SbarText = session.findById("wnd[0]/sbar").Text
End Function

' ==========================================================================
' v23: XU LY HOP THOAI (POPUP) SAP
' ==========================================================================
' Cua so popup tren cung dang mo (wnd[1], wnd[2]...), Nothing neu khong co
Function PopupWnd(session As Object) As Object
    Dim k As Long, w As Object
    On Error Resume Next
    For k = 3 To 1 Step -1
        Set w = Nothing
        Set w = session.findById("wnd[" & k & "]")
        Err.Clear
        If Not w Is Nothing Then
            Set PopupWnd = w
            Exit Function
        End If
    Next k
End Function

' Gom noi dung chu + cac nut bam tren popup
Sub CollectPopup(o As Object, ByRef txt As String, btns As Collection, ByVal depth As Long)
    Dim i As Long, n As Long, ch As Object, tp As String, t As String
    If depth > 10 Then Exit Sub
    On Error Resume Next
    n = -1
    n = o.Children.Count
    If n <= 0 Then Exit Sub
    For i = 0 To n - 1
        Set ch = Nothing
        Set ch = ChildAt(o, i)
        If Not ch Is Nothing Then
            tp = "": t = ""
            tp = SapType(ch)
            t = Trim(ch.Text)
            If tp = "GuiButton" Then
                btns.Add ch
            ElseIf t <> "" And (tp = "GuiLabel" Or tp = "GuiTextField" Or tp = "GuiCTextField" Or tp = "GuiTitlebar") Then
                txt = txt & t & " "
            End If
            CollectPopup ch, txt, btns, depth + 1
        End If
    Next i
End Sub

Function PopupText(w As Object) As String
    Dim txt As String, btns As New Collection
    On Error Resume Next
    txt = w.Text & ": "
    CollectPopup w, txt, btns, 0
    PopupText = Trim(txt)
End Function

Private Function HasWord(ByVal t As String, ByVal words As String) As Boolean
    Dim a As Variant, i As Long
    a = Split(words, "|")
    For i = LBound(a) To UBound(a)
        If InStr(1, t, " " & a(i) & " ", vbTextCompare) > 0 Then HasWord = True: Exit Function
    Next i
End Function

' Cham diem nut: positive=True tim nut Save/Yes/Tiep tuc; False tim nut Cancel/No
Function BtnScore(b As Object, ByVal positive As Boolean) As Long
    Dim t As String, id As String, ch As Variant
    On Error Resume Next
    t = b.Text & " " & b.Tooltip
    id = UCase(b.id)
    On Error GoTo 0
    For Each ch In Array("&", ".", ",", ":", "(", ")", "?", "!", "-", "/")
        t = Replace(t, ch, " ")
    Next ch
    t = " " & LCase(Trim(t)) & " "
    If positive Then
        If HasWord(t, "no|cancel|khong|huy|edit|check|messages|message|log|hold|park|back|exit|print") Then Exit Function
        If HasWord(t, "save|luu") Then BtnScore = 100: Exit Function
        If HasWord(t, "yes|continue|ok|confirm|enter|tiep") Then BtnScore = 80: Exit Function
        If InStr(id, "SPOP-OPTION1") > 0 Or InStr(id, "BUTTON_1") > 0 Then BtnScore = 60: Exit Function
        If InStr(id, "TBAR[0]/BTN[0]") > 0 Then BtnScore = 40: Exit Function
    Else
        If HasWord(t, "cancel|huy|no|khong|back") Then BtnScore = 100: Exit Function
        If InStr(id, "SPOP-OPTION_CAN") > 0 Then BtnScore = 90: Exit Function
        If InStr(id, "SPOP-OPTION2") > 0 Or InStr(id, "BUTTON_2") > 0 Then BtnScore = 70: Exit Function
        If InStr(id, "TBAR[0]/BTN[12]") > 0 Then BtnScore = 60: Exit Function
    End If
End Function

' Bam nut phu hop tren popup w. Tra ve True neu da bam duoc.
Function PressPopupButton(session As Object, w As Object, ByVal positive As Boolean) As Boolean
    Dim btns As New Collection, txt As String, b As Object, best As Object, bestSc As Long, sc As Long
    CollectPopup w, txt, btns, 0
    For Each b In btns
        sc = BtnScore(b, positive)
        If sc > bestSc Then bestSc = sc: Set best = b
    Next b
    On Error Resume Next
    If Not best Is Nothing Then
        best.press
    Else
        w.sendVKey IIf(positive, 0, 12)     ' Enter = Tiep tuc / F12 = Huy
    End If
    PressPopupButton = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    WaitSap session
End Function

Sub ClosePopupCancel(session As Object)
    Dim w As Object, k As Long
    For k = 1 To 3
        Set w = PopupWnd(session)
        If w Is Nothing Then Exit Sub
        PressPopupButton session, w, False
    Next k
End Sub

' Nhu CheckPopup nhung dung cho PO: hop thoai THONG BAO (khong co nut No/Cancel,
' chi co Tiep tuc) thi tu bam Tiep tuc va chay tiep; hop thoai hoi Yes/No thi dung lai.
Sub CheckPopupPO(session As Object)
    Dim k As Long, w As Object, txt As String, btns As Collection, b As Object, isQuestion As Boolean
    WaitSap session
    For k = 1 To 4
        Set w = PopupWnd(session)
        If w Is Nothing Then Exit Sub
        txt = "": Set btns = New Collection
        CollectPopup w, txt, btns, 0
        isQuestion = False
        For Each b In btns
            If BtnScore(b, False) >= 70 Then isQuestion = True
        Next b
        If isQuestion Then Exit For
        On Error Resume Next
        w.sendVKey 0
        Err.Clear
        On Error GoTo 0
        WaitSap session
    Next k
    CheckPopup session
End Sub

' Hop thoai/Thong bao co noi PO dang co LOI (khong chi canh bao)?
Function IsFaultyMsg(ByVal txt As String) As Boolean
    Dim t As String
    t = LCase(txt)
    IsFaultyMsg = (InStr(t, "faulty") > 0 Or InStr(t, "error") > 0 Or InStr(t, "incorrect") > 0 Or _
                   InStr(t, "incomplete") > 0 Or InStr(t, "khong hop le") > 0)
End Function

' --------------------------------------------------------------------------
' v23: SAVE PO va tu xu ly hop thoai SAP hien ra truoc khi luu
'  - Hop thoai CANH BAO (vd "co thong bao, ban co muon luu?") -> bam Save/Yes
'  - Canh bao (chu vang) o thanh trang thai -> Enter roi bam Save lan nua
'  - Hop thoai noi PO con dong LOI -> hoi ban (mac dinh No) truoc khi luu de
'  - Tra ve So PO; khong luu duoc thi bao loi kem noi dung SAP
' --------------------------------------------------------------------------
Function SavePOWithPopups(session As Object, ByVal sup As String) As String
    Dim k As Long, w As Object, txt As String, mt As String, st As String, nSave As Long, num As String, nIdle As Long
    curStep = "bam Save PO"
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    nSave = 1
    For k = 1 To 15
        WaitSap session
        Set w = PopupWnd(session)
        If Not w Is Nothing Then
            txt = PopupText(w)
            curStep = "xu ly hop thoai SAP khi Save (" & Left(txt, 150) & ")"
            If IsFaultyMsg(txt) Then
                On Error Resume Next
                TramActivate
                On Error GoTo 0
                If Hoi("Khi Save PO " & sup & ", SAP bao:" & vbCrLf & vbCrLf & txt & vbCrLf & vbCrLf & _
                          "PO van con dong LOI. Van LUU DE (Save) PO nay?" & vbCrLf & _
                          "(No = huy Save, giu nguyen man hinh SAP de ban tu kiem tra)", _
                          vbYesNo + vbExclamation + vbDefaultButton2, "CreatePO " & MACRO_VERSION) = vbNo Then
                    PressPopupButton session, w, False
                    Err.Raise vbObjectError + 620, "SavePO", "CHUA SAVE - ban chon khong luu khi SAP bao: " & txt
                End If
            End If
            If Not PressPopupButton(session, w, True) Then
                Err.Raise vbObjectError + 621, "SavePO", "Khong bam duoc nut Save/Yes tren hop thoai SAP: " & txt
            End If
        Else
            st = "": mt = ""
            On Error Resume Next
            st = session.findById("wnd[0]/sbar").Text
            mt = session.findById("wnd[0]/sbar").MessageType
            On Error GoTo 0
            If mt = "S" Then
                num = ExtractDocNumber(st)
                If num <> "" Then
                    SavePOWithPopups = num
                    Exit Function
                End If
            ElseIf mt = "E" Or mt = "A" Then
                Err.Raise vbObjectError + 622, "SavePO", "SAP bao loi khi Save: " & st
            ElseIf mt = "W" Or mt = "I" Then
                ' canh bao o thanh trang thai -> xac nhan (Enter) roi Save lai
                If nSave >= 4 Then
                    Err.Raise vbObjectError + 623, "SavePO", "Da bam Save " & nSave & " lan, SAP van canh bao: " & st
                End If
                curStep = "xac nhan canh bao SAP khi Save (" & st & ")"
                session.findById("wnd[0]").sendVKey 0
                WaitSap session
                If PopupWnd(session) Is Nothing Then
                    session.findById("wnd[0]/tbar[0]/btn[11]").press
                    nSave = nSave + 1
                End If
            Else
                nIdle = nIdle + 1
                If nIdle > 5 Then Exit For
                Application.Wait Now + TimeSerial(0, 0, 1)
            End If
        End If
    Next k
    Err.Raise vbObjectError + 624, "SavePO", "Da bam Save nhung chua nhan duoc So PO. Thanh trang thai SAP: " & st
End Function

' Ban tong quat cua SavePOWithPopups, dung cho ca PR (v24-auto2)
Function SaveDocWithPopups(session As Object, ByVal nhan As String, Optional ByVal choKhongDoi As Boolean = False) As String
    Dim k As Long, w As Object, txt As String, mt As String, st As String, nSave As Long, num As String, nIdle As Long
    curStep = "bam Save " & nhan
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    nSave = 1
    For k = 1 To 15
        WaitSap session
        Set w = PopupWnd(session)
        If Not w Is Nothing Then
            txt = PopupText(w)
            curStep = "xu ly hop thoai SAP khi Save (" & Left(txt, 150) & ")"
            If IsFaultyMsg(txt) Then
                On Error Resume Next
                TramActivate
                On Error GoTo 0
                If Hoi("Khi Save " & nhan & ", SAP bao:" & vbCrLf & vbCrLf & txt & vbCrLf & vbCrLf & _
                          nhan & " van con dong LOI. Van LUU DE (Save)?" & vbCrLf & _
                          "(No = huy Save, giu nguyen man hinh SAP de ban tu kiem tra)", _
                          vbYesNo + vbExclamation + vbDefaultButton2, "Save " & nhan & " " & MACRO_VERSION) = vbNo Then
                    PressPopupButton session, w, False
                    Err.Raise vbObjectError + 620, "Save" & nhan, "CHUA SAVE - ban chon khong luu khi SAP bao: " & txt
                End If
            End If
            If Not PressPopupButton(session, w, True) Then
                Err.Raise vbObjectError + 621, "Save" & nhan, "Khong bam duoc nut Save/Yes tren hop thoai SAP: " & txt
            End If
        Else
            st = "": mt = ""
            On Error Resume Next
            st = session.findById("wnd[0]/sbar").Text
            mt = session.findById("wnd[0]/sbar").MessageType
            On Error GoTo 0
            If mt = "S" Then
                num = ExtractDocNumber(st)
                If num <> "" Then
                    SaveDocWithPopups = num
                    Exit Function
                End If
                ' ME22N: SAP bao "No data was changed" (khong co gi de luu) = du lieu tren SAP da dung
                If choKhongDoi And (InStr(1, st, "no data", vbTextCompare) > 0 Or InStr(1, st, "not changed", vbTextCompare) > 0) Then
                    SaveDocWithPopups = "(KHONG DOI)"
                    Exit Function
                End If
            ElseIf mt = "E" Or mt = "A" Then
                Err.Raise vbObjectError + 622, "Save" & nhan, "SAP bao loi khi Save: " & st
            ElseIf mt = "W" Or mt = "I" Then
                ' canh bao o thanh trang thai -> xac nhan (Enter) roi Save lai
                If nSave >= 4 Then
                    Err.Raise vbObjectError + 623, "Save" & nhan, "Da bam Save " & nSave & " lan, SAP van canh bao: " & st
                End If
                curStep = "xac nhan canh bao SAP khi Save (" & st & ")"
                session.findById("wnd[0]").sendVKey 0
                WaitSap session
                If PopupWnd(session) Is Nothing Then
                    session.findById("wnd[0]/tbar[0]/btn[11]").press
                    nSave = nSave + 1
                End If
            Else
                nIdle = nIdle + 1
                If nIdle > 5 Then Exit For
                Application.Wait Now + TimeSerial(0, 0, 1)
            End If
        End If
    Next k
    Err.Raise vbObjectError + 624, "Save" & nhan, "Da bam Save nhung chua nhan duoc so " & nhan & ". Thanh trang thai SAP: " & st
End Function

' True neu thanh trang thai SAP dang bao loi (loai E/A)
Function SbarIsError(session As Object) As Boolean
    Dim mt As String
    On Error Resume Next
    mt = session.findById("wnd[0]/sbar").MessageType
    On Error GoTo 0
    SbarIsError = (mt = "E" Or mt = "A")
End Function

' Neu thanh trang thai SAP bao LOI (chu do, loai E/A) thi dung lai kem noi dung
Sub CheckSbarError(session As Object)
    Dim t As String, mt As String
    On Error Resume Next
    mt = session.findById("wnd[0]/sbar").MessageType
    t = session.findById("wnd[0]/sbar").Text
    On Error GoTo 0
    If mt = "E" Or mt = "A" Then
        Err.Raise vbObjectError + 604, "SAP", "SAP bao loi: " & t
    End If
End Sub

' Tim chi so cot cua bang (table control) theo ten truong, vd "MEPO1211-BANFN". -1 neu khong co.
Function TblColIndex(TBL As Object, ByVal fieldName As String) As Long
    Dim c As Long, n As Long, nm As String
    TblColIndex = -1
    On Error Resume Next
    n = TBL.Columns.Count
    For c = 0 To n - 1
        nm = ""
        nm = TBL.GetCell(0, c + 0).Name
        If UCase(nm) = UCase(fieldName) Then
            TblColIndex = c
            Exit Function
        End If
    Next c
End Function

' Doc so SAP hien (vd "5.832" / "5,832" / "11,500") theo kieu dau thap phan nao
' cho ket qua gan voi gia tri mong muon nhat
Function ParseBest(ByVal s As String, ByVal expected As Double) As Double
    Dim a As Double, b As Double
    If s = "" Then ParseBest = -1: Exit Function
    a = ParseSapNum(s, gDecSep)
    b = ParseSapNum(s, IIf(gDecSep = ",", ".", ","))
    If Abs(b - expected) < Abs(a - expected) Then ParseBest = b Else ParseBest = a
End Function

' Suy ra so dong item PR khi cot P trong: thu tu dong trong cung SoPR x PR_ITEM_STEP
Function PrItemGuess(wsPR As Worksheet, ByVal r As Long) As Long
    Dim i As Long, n As Long, soPR As String
    soPR = Trim(CStr(wsPR.Range("L" & r).Value))
    For i = 3 To r
        If Trim(CStr(wsPR.Range("L" & i).Value)) = soPR Then n = n + 1
    Next i
    PrItemGuess = n * PR_ITEM_STEP
End Function


' --------------------------------------------------------------------------
' 3) NHAP KHO / GOODS RECEIPT (MIGO) - 1 chung tu cho moi PO da tao
'
' LOGIC (theo mo ta cua ban):
'  - Lay cac PO da co So PO (cot J sheet PO) va chua nhap kho.
'  - MIGO: Goods Receipt (A01) / Purchase Order (R01), go So PO, Enter.
'  - Tren bang dong (theo file record): tich o OK (thuc nhan) va go cot Text
'    cho TAT CA cac dong: "Mua hang phuc vu Bep thang 10" (thang = thang hien
'    tai luc chay). Mau noi dung o sheet Mapping o B19, sua duoc.
'  - Dung lai hoi truoc khi Post. Ghi So chung tu vat tu vao sheet MIGO.
'
' Sheet "MIGO" (macro TongHopMIGO tu tao lai, giu cot G/H da co):
'   A STT | B So_PO | C Nha_cung_cap | D So_dong | E Tong_tien | F Noi_dung_text
'   G So_chung_tu_MIGO | H Status
' --------------------------------------------------------------------------

' Mau noi dung text MIGO: lay o Mapping!B19; neu trong thi ghi mau mac dinh.
' {THANG} se duoc thay bang thang hien tai, vd 10.
Function MigoTextTemplate() As String
    Dim wsM As Worksheet, t As String
    Set wsM = ThisWorkbook.Sheets("Mapping")
    If Trim(CStr(wsM.Range("A19").Value)) = "" Then
        wsM.Range("A18").Value = "4) NOI DUNG TEXT KHI NHAP KHO (MIGO)"
        wsM.Range("A18").Font.Bold = True
        wsM.Range("A19").Value = "Noi_dung_text_MIGO"
        wsM.Range("C19").Value = "{THANG} = thang luc chay MIGO, vd 10"
    End If
    t = Trim(CStr(wsM.Range("B19").Value))
    If t = "" Then
        ' "Mua hang phuc vu Bep thang {THANG}" co dau tieng Viet
        t = "Mua h" & ChrW(224) & "ng ph" & ChrW(7909) & "c v" & ChrW(7909) & " B" & ChrW(7871) & "p th" & ChrW(225) & "ng {THANG}"
        wsM.Range("B19").Value = t
    End If
    MigoTextTemplate = t
End Function

Function MigoText() As String
    MigoText = Replace(MigoTextTemplate(), "{THANG}", CStr(Month(Date)))
End Function


' --------------------------------------------------------------------------
' TONG HOP MIGO: liet ke cac PO da tao (sheet PO cot J) can nhap kho
' --------------------------------------------------------------------------
Sub TongHopMIGO(Optional ByVal silent As Boolean = False)
    Dim wsPO As Worksheet, wsMG As Worksheet
    Set wsPO = ThisWorkbook.Sheets("PO")
    Set wsMG = ThisWorkbook.Sheets("MIGO")

    ' giu lai So chung tu + Status da co (theo So PO)
    Dim keep As Object, r As Long, lastR As Long, k As String
    Set keep = CreateObject("Scripting.Dictionary")
    If wsMG.Range("G1").Value = "So_chung_tu_MIGO" Then
        lastR = wsMG.Cells(wsMG.Rows.Count, "B").End(xlUp).Row
        For r = 3 To lastR
            k = Trim(CStr(wsMG.Cells(r, "B").Value))
            If k <> "" Then keep(k) = Array(wsMG.Cells(r, "G").Value, wsMG.Cells(r, "H").Value)
        Next r
    End If

    wsMG.Cells.Clear
    wsMG.Range("A1:H1").Value = Array("STT", "So_PO", "Nha_cung_cap", "So_dong", "Tong_tien", "Noi_dung_text", "So_chung_tu_MIGO", "Status")
    wsMG.Range("A2:H2").Value = Array("Tu dong", "Tu dong = sheet PO cot J", "Tu dong", "Tu dong", "Tu dong", _
                                      "Tu dong (mau o Mapping B19)", "<< Macro PostMIGO >>", "<< Macro PostMIGO >>")
    With wsMG.Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
    End With
    wsMG.Range("A2:H2").Font.Italic = True
    wsMG.Range("A2:H2").Font.Color = RGB(118, 118, 118)

    Dim txt As String, outR As Long, soPO As String, kv As Variant
    txt = MigoText()
    outR = 3
    lastR = wsPO.Cells(wsPO.Rows.Count, "B").End(xlUp).Row
    For r = 3 To lastR
        soPO = Trim(CStr(wsPO.Cells(r, "J").Value))
        If soPO <> "" Then
            wsMG.Cells(outR, "A").Value = outR - 2
            wsMG.Cells(outR, "B").NumberFormat = "@"
            wsMG.Cells(outR, "B").Value = soPO
            wsMG.Cells(outR, "C").Value = wsPO.Cells(r, "B").Value
            wsMG.Cells(outR, "D").Value = wsPO.Cells(r, "D").Value
            wsMG.Cells(outR, "E").Value = wsPO.Cells(r, "F").Value
            wsMG.Cells(outR, "F").Value = txt
            If keep.Exists(soPO) Then
                kv = keep(soPO)
                wsMG.Cells(outR, "G").NumberFormat = "@"
                wsMG.Cells(outR, "G").Value = kv(0)
                wsMG.Cells(outR, "H").Value = kv(1)
                If Trim(CStr(kv(0))) <> "" Then wsMG.Cells(outR, "F").Value = "(da nhap kho)"
                If Trim(CStr(kv(0))) = "DA NHAP TRUOC" Then wsMG.Cells(outR, "G").Interior.Color = RGB(255, 235, 156)
            End If
            outR = outR + 1
        End If
    Next r
    wsMG.Range("E3:E" & Application.Max(3, outR)).NumberFormat = "#,##0"
    wsMG.Columns("A:H").AutoFit

    If Not silent Then
        wsMG.Activate
        Hoi "Da liet ke " & (outR - 3) & " PO da tao (can nhap kho neu cot So_chung_tu_MIGO con trong).", _
               vbInformation, "TongHopMIGO " & MACRO_VERSION
    End If
End Sub


' --------------------------------------------------------------------------
' NHAP KHO (MIGO) cho cac PO chua co So chung tu MIGO
' --------------------------------------------------------------------------
Sub PostMIGO()
    TongHopMIGO True

    Dim wsMG As Worksheet
    Set wsMG = ThisWorkbook.Sheets("MIGO")

    Dim lastR As Long, r As Long, todo As New Collection, summary As String, v As Variant, txt As String
    txt = MigoText()
    lastR = wsMG.Cells(wsMG.Rows.Count, "B").End(xlUp).Row
    For r = 3 To lastR
        If Trim(CStr(wsMG.Cells(r, "B").Value)) <> "" And Trim(CStr(wsMG.Cells(r, "G").Value)) = "" Then
            todo.Add r
            summary = summary & "  - PO " & wsMG.Cells(r, "B").Value & " (" & wsMG.Cells(r, "C").Value & ", " & _
                      wsMG.Cells(r, "D").Value & " dong)" & vbCrLf
        End If
    Next r

    If todo.Count = 0 Then
        wsMG.Activate
        Hoi "Khong co PO nao can nhap kho (cac PO da co So chung tu MIGO, hoac chua tao PO).", vbInformation, "PostMIGO " & MACRO_VERSION
        Exit Sub
    End If

    If Hoi("Se nhap kho (MIGO) " & todo.Count & " PO:" & vbCrLf & summary & vbCrLf & _
              "Noi dung text: xem cot F sheet MIGO (thang " & Month(Date) & ")" & vbCrLf & vbCrLf & "Bat dau?", vbYesNo + vbQuestion, "PostMIGO " & MACRO_VERSION) = vbNo Then Exit Sub

    Dim session As Object
    Set session = GetSapSession()
    If session Is Nothing Then Exit Sub
    Set ctlCache = CreateObject("Scripting.Dictionary")

    Dim soPO As String, nTick As Long, ans As VbMsgBoxResult, docNo As String, msg As String
    Dim nSkip As Long, skipList As String
    For Each v In todo
        soPO = Trim(CStr(wsMG.Cells(v, "B").Value))
        On Error GoTo MigoError

        nTick = EnterMIGO(session, soPO, txt, CLng(Val(wsMG.Cells(v, "D").Value)))

        If nTick = -1 Then
            ' PO da duoc nhap kho truoc (vd MIGO thu cong) -> ghi chu, chuyen PO tiep theo
            wsMG.Cells(v, "G").Value = "DA NHAP TRUOC"
            wsMG.Cells(v, "H").Value = gMigoNote & " (" & Format(Now, "dd/mm/yyyy hh:nn") & ")"
            wsMG.Cells(v, "G").Interior.Color = RGB(255, 235, 156)
            nSkip = nSkip + 1
            skipList = skipList & "  - PO " & soPO & " (" & wsMG.Cells(v, "C").Value & ")" & vbCrLf
            GoTo NextMigo
        End If

        ' Dung lai xac nhan truoc khi Post
        On Error Resume Next
        TramActivate
        On Error GoTo MigoError
        ans = Hoi("MIGO cho PO " & soPO & " (" & wsMG.Cells(v, "C").Value & "):" & vbCrLf & _
                     "  - Da tich OK + go Text: " & nTick & " / " & wsMG.Cells(v, "D").Value & " dong" & vbCrLf & _
                     "  - Text: noi dung o cot F sheet MIGO" & vbCrLf & _
                     IIf(gMigoNote <> "", "  - CHU Y: " & gMigoNote & vbCrLf, "") & vbCrLf & _
                     "Kiem tra tren SAP (Alt+Tab) roi tra loi: POST nhap kho?", _
                     vbYesNo + IIf(nTick = CLng(Val(wsMG.Cells(v, "D").Value)), vbQuestion, vbExclamation) + vbDefaultButton2, _
                     "PostMIGO " & MACRO_VERSION)
        If ans = vbNo Then
            wsMG.Cells(v, "H").Value = "CHUA POST - dung theo yeu cau"
            Exit Sub
        End If

        curStep = "bam Post MIGO"
        session.findById("wnd[0]/tbar[0]/btn[11]").press
        CheckPopup session
        msg = session.findById("wnd[0]/sbar").Text
        docNo = ExtractDocNumber(msg)
        CheckSbarError session
        On Error GoTo 0

        wsMG.Cells(v, "G").NumberFormat = "@"
        wsMG.Cells(v, "G").Value = docNo
        wsMG.Cells(v, "H").Value = msg & IIf(gMigoNote <> "", " | " & gMigoNote, "")
        GoTo NextMigo
MigoError:
        wsMG.Cells(v, "H").Value = "LOI, buoc [" & curStep & "]: " & Err.Description
        Hoi "Loi khi nhap kho PO " & soPO & ":" & vbCrLf & Err.Description & vbCrLf & _
               "Buoc: " & curStep & vbCrLf & _
               "Macro da DUNG, man hinh SAP giu nguyen. Neu loi 'khong tim thay control': chay SAP_ChanDoan va gui file.", _
               vbCritical, "PostMIGO " & MACRO_VERSION
        Exit Sub
NextMigo:
    Next v

    TongHopMIGO True
    wsMG.Activate
    Hoi "Da xu ly xong MIGO. Kiem tra cot So_chung_tu_MIGO / Status." & vbCrLf & _
           IIf(nSkip > 0, vbCrLf & nSkip & " PO da nhap kho TRUOC do (bo qua, da ghi chu 'DA NHAP TRUOC'):" & vbCrLf & skipList, "") & _
           vbCrLf & "Nho Ctrl+S luu file.", vbInformation, "PostMIGO " & MACRO_VERSION
End Sub

' --------------------------------------------------------------------------
' Nhap 1 MIGO (CHUA Post) - theo file record cua ban:
'   1. /nmigo (Goods Receipt / Purchase Order), go So PO, Enter
'   2. Tren bang dong: moi dong tich o OK (GOITEM-TAKE_IT) + go cot Text (GOITEM-SGTXT)
' Tra ve so dong da xu ly.
' --------------------------------------------------------------------------
Function EnterMIGO(session As Object, ByVal soPO As String, ByVal txt As String, ByVal nExpected As Long) As Long
    Dim c As Object, TBL As Object, cOK As Long, cTxt As Long, nRows As Long, i As Long, cel As Object

    curStep = "mo MIGO"
    SapCuaSo session
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmigo"
    session.findById("wnd[0]").sendVKey 0
    CheckPopup session

    ' Dam bao dang o Goods Receipt / Purchase Order (SAP thuong nho san)
    curStep = "chon Goods Receipt / Purchase Order"
    Set c = FindCtlSoft(session, "cmbGODYNPRO-ACTION")
    If Not c Is Nothing Then
        If c.key <> MIGO_ACTION Then c.key = MIGO_ACTION
        CheckPopup session
    End If
    Set c = FindCtlSoft(session, "cmbGODYNPRO-REFDOC")
    If Not c Is Nothing Then
        If c.key <> MIGO_REFDOC Then c.key = MIGO_REFDOC
        CheckPopup session
    End If

    curStep = "go So PO " & soPO
    Set c = FindCtl(session, "ctxtGODYNPRO-PO_NUMBER")
    c.Text = soPO
    session.findById("wnd[0]").sendVKey 0
    CheckPopup session

    ' Doc thong bao SAP + bang dong de nhan biet PO DA NHAP KHO truoc do
    Dim sbT As String, sbType As String
    gMigoNote = ""
    On Error Resume Next
    sbT = session.findById("wnd[0]/sbar").Text
    sbType = session.findById("wnd[0]/sbar").MessageType
    On Error GoTo 0

    curStep = "tim cot OK / Text tren bang dong MIGO"
    Set TBL = FindCtlSoft(session, MIGO_TBL_ID)
    nRows = 0
    cOK = -1: cTxt = -1
    If Not TBL Is Nothing Then
        On Error Resume Next
        nRows = TBL.RowCount
        On Error GoTo 0
        cOK = TblColIndex(TBL, "GOITEM-TAKE_IT")
        cTxt = TblColIndex(TBL, "GOITEM-SGTXT")
    End If

    ' Khong con dong nao de nhan (bang trong / SAP bao loi) -> coi la DA NHAP KHO, ghi chu lai, bo qua
    ' Dau hieu: thong bao "does not contain any selectable items" (hoac tuong tu),
    ' bang trong, hoac ca 2 cot OK + Text deu khong doc duoc (bang khong co dong nao)
    Dim noItemMsg As Boolean, lt As String
    lt = LCase(sbT)
    noItemMsg = (InStr(lt, "selectable") > 0 Or InStr(lt, "does not contain") > 0 Or InStr(lt, "no items") > 0 _
                 Or InStr(lt, "kh" & ChrW(244) & "ng ch" & ChrW(7913) & "a") > 0 Or InStr(lt, "kh" & ChrW(244) & "ng c" & ChrW(243)) > 0)
    If TBL Is Nothing Or nRows = 0 Or noItemMsg Or (cOK < 0 And cTxt < 0) Then
        gMigoNote = "PO khong con dong nao de nhap kho (da MIGO truoc do?)" & IIf(sbT <> "", " | SAP bao: " & sbT, "")
        EnterMIGO = -1
        Exit Function
    End If
    If sbType = "E" Or sbType = "A" Then
        Err.Raise vbObjectError + 704, "SAP", "SAP bao loi: " & sbT
    End If
    If cOK < 0 Or cTxt < 0 Then
        Err.Raise vbObjectError + 701, "EnterMIGO", "Bang dong MIGO khong co cot OK (TAKE_IT) hoac Text (SGTXT)" & _
            IIf(sbT <> "", " | SAP bao: " & sbT, "")
    End If

    If nExpected > 0 And nRows > nExpected Then nRows = nExpected
    If nExpected > 0 And nRows < nExpected Then
        gMigoNote = "Chi con " & nRows & "/" & nExpected & " dong de nhan (mot phan da MIGO truoc do?)"
    End If

    For i = 0 To nRows - 1
        curStep = "tich OK + go Text dong " & (i + 1)
        Set cel = MigoRowCell(session, i, cOK)
        If Not cel.Selected Then cel.Selected = True
        Set cel = MigoRowCell(session, i, cTxt)
        cel.Text = txt
    Next i
    session.findById("wnd[0]").sendVKey 0
    CheckPopup session
    CheckSbarError session

    EnterMIGO = nRows
End Function

' Lay o (dong idx, cot col) cua bang dong MIGO, cuon neu can
Function MigoRowCell(session As Object, ByVal idx As Long, ByVal col As Long) As Object
    Dim TBL As Object, vis As Long, pos As Long, mx As Long, newPos As Long, tries As Long
    Set TBL = FindCtl(session, MIGO_TBL_ID)
    vis = TBL.VisibleRowCount
    pos = TBL.VerticalScrollbar.Position
    For tries = 1 To 4
        If idx >= pos And idx <= pos + vis - 1 Then Exit For
        mx = TBL.VerticalScrollbar.Maximum
        If idx < pos Then newPos = idx Else newPos = idx - vis + 1
        If newPos > mx Then newPos = mx
        If newPos < 0 Then newPos = 0
        If tries <= 2 Then
            TBL.VerticalScrollbar.Position = newPos
        Else
            On Error Resume Next
            TBL.GetCell(IIf(idx < pos, 0, vis - 1) + 0, col + 0).SetFocus
            On Error GoTo 0
            session.findById("wnd[0]").sendVKey IIf(idx < pos, 81, 82)
        End If
        WaitSap session
        Set TBL = FindCtl(session, MIGO_TBL_ID)
        vis = TBL.VisibleRowCount
        pos = TBL.VerticalScrollbar.Position
    Next tries
    If idx - pos < 0 Or idx - pos > vis - 1 Then
        Err.Raise vbObjectError + 703, "MigoRowCell", "Khong cuon bang dong MIGO toi dong " & (idx + 1)
    End If
    Set MigoRowCell = TBL.GetCell(idx - pos + 0, col + 0)
End Function


' --------------------------------------------------------------------------
' TIEN ICH: chay ca 3 buoc lien tiep (chi dung sau khi da test tung buoc rieng!)
' --------------------------------------------------------------------------
Sub RunAll()
    CreatePR
    CreatePO
    PostMIGO
End Sub

' ================================================================================
'  v23-Tram: CHAY TU DONG TU TRAM SAP (TramSAP.ps1) - KHONG HIEN HOP THOAI
'  Khi TRAM_AUTO = True, moi hop thoai Hoi() tu tra loi:
'   - Cau hoi Luu (Save) PR/PO hoac Post MIGO: chi tra loi YES khi TRAM_LUU_THAT = True
'     VA hop thoai khong o muc canh bao (lech SL/gia, thieu dong) -> neu canh bao thi NO (khong luu).
'   - Cau hoi "Bat dau nhap vao SAP?": YES.  OK/Cancel (can thao tac tay): CANCEL.
'  Moi noi dung hop thoai duoc ghi vao TRAM_LOG de gui ve giao dien.
'  Chay tay (Alt+F8) van hien hop thoai nhu v22.
' ================================================================================

Function Hoi(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As VbMsgBoxResult
    If Not TRAM_AUTO Then
        Hoi = VBA.Interaction.MsgBox(prompt, buttons, title)
        Exit Function
    End If
    Dim kind As Long, isSave As Boolean, canhBao As Boolean, ans As VbMsgBoxResult
    kind = buttons And 7
    isSave = (InStr(UCase(prompt), "(SAVE)") > 0) Or (InStr(UCase(prompt), "POST NHAP KHO") > 0)
    canhBao = ((buttons And vbExclamation) = vbExclamation) Or (InStr(UCase(prompt), "CANH BAO") > 0)
    Select Case kind
        Case vbYesNo, vbYesNoCancel
            If isSave Then
                If TRAM_LUU_THAT And Not canhBao Then ans = vbYes Else ans = vbNo
            Else
                ans = vbYes
            End If
        Case vbOKCancel, vbRetryCancel
            ans = vbCancel
        Case Else
            ans = vbOK
    End Select
    TRAM_LOG = TRAM_LOG & "[" & title & "] " & Replace(Replace(prompt, vbCrLf, " | "), vbLf, " | ") & _
               "  => " & Choose(ans, "OK", "Cancel", "Abort", "Retry", "Ignore", "Yes", "No") & vbLf
    Hoi = ans
End Function

Sub TramActivate()
    On Error Resume Next
    If Not TRAM_AUTO Then AppActivate Application.Caption
End Sub

' Xoa du lieu lan chay truoc: CongNo A3:K, cot macro ghi sheet PR (L,M,P..S), sheet PO
Function TramChuanBi() As String
    Dim wsC As Worksheet, wsPR As Worksheet
    Set wsC = ThisWorkbook.Sheets("CongNo")
    Set wsPR = ThisWorkbook.Sheets("PR")
    EnsurePRHeaders wsPR
    wsC.Range("A3:K3000").ClearContents
    wsC.Range("D3:D3000").NumberFormat = "@"
    wsC.Range("B3:B3000").NumberFormat = "@"
    wsPR.Range("L3:M3000").ClearContents
    wsPR.Range("P3:S3000").ClearContents
    On Error Resume Next
    ThisWorkbook.Sheets("PO").Range("A3:K3000").ClearContents
    On Error GoTo 0
    TramChuanBi = "OK"
End Function

' Ghi ma NCC SAP (LIFNR) vao Mapping cot E:F
Function TramDatMaNCC(ByVal nm As String, ByVal cd As String) As String
    Dim wsM As Worksheet
    Set wsM = ThisWorkbook.Sheets("Mapping")
    EnsureVendorTable wsM
    UpsertVendor wsM, Trim(nm), Trim(cd)
    TramDatMaNCC = "OK"
End Function

' Ghi So PR + so dong item cho dong r sheet PR (dung khi tao PO)
Function TramGanPR(ByVal r As Long, ByVal soPR As String, ByVal item As String) As String
    With ThisWorkbook.Sheets("PR")
        .Range("L" & r).NumberFormat = "@"
        .Range("L" & r).Value = soPR
        If item <> "" Then
            .Range("P" & r).NumberFormat = "@"
            .Range("P" & r).Value = item
        End If
    End With
    TramGanPR = "OK"
End Function

Function TramChayPR(ByVal luuThat As Boolean) As String
    TRAM_AUTO = True: TRAM_LUU_THAT = luuThat: TRAM_LOG = ""
    On Error Resume Next
    CreatePR
    If Err.Number <> 0 Then TRAM_LOG = TRAM_LOG & "LOI VBA: " & Err.Description & vbLf
    On Error GoTo 0
    TRAM_AUTO = False
    TramChayPR = TRAM_LOG
End Function

Function TramChayPO(ByVal luuThat As Boolean) As String
    TRAM_AUTO = True: TRAM_LUU_THAT = luuThat: TRAM_LOG = ""
    On Error Resume Next
    CreatePO
    If Err.Number <> 0 Then TRAM_LOG = TRAM_LOG & "LOI VBA: " & Err.Description & vbLf
    On Error GoTo 0
    TRAM_AUTO = False
    TramChayPO = TRAM_LOG
End Function

Function TramPhienBan() As String
    TramPhienBan = MACRO_VERSION
End Function


' ================================================================================
'  v23-Tram: SUA PO DA TAO (ME22N) - tang/giam so luong dong co san, them dong PR bo sung
'  Sheet "PO_SUA" (tram tu tao), du lieu tu dong 3, MOT PO moi lan chay:
'   A So_PO | B Loai (SUA/THEM) | C Dong_PO (EBELP, SUA) | D So_PR (THEM) | E Dong_PR (THEM)
'   F Ma_vat_tu | G SL_moi | H Don_gia | I Ten_hang | J Status (macro ghi) | K Dong_PO_moi (macro ghi)
'  SUA: tim dong theo EBELP, kiem tra dung ma vat tu, go SL moi.
'  THEM: dong trong dau tien, go So PR + dong PR, SL, gia (nhu ME21N).
'  Doc lai toan bo dong da dong vao: SL + gia phai BANG TUYET DOI cot G/H, lech thi KHONG luu.
' ================================================================================
Sub EnsurePOSua()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("PO_SUA")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "PO_SUA"
        ws.Range("A1:K1").Value = Array("So_PO", "Loai", "Dong_PO", "So_PR", "Dong_PR", "Ma_vat_tu", "SL_moi", "Don_gia", "Ten_hang", "Status", "Dong_PO_moi")
        ws.Range("A2:K2").Value = Array("Tram ghi", "SUA/THEM", "EBELP", "THEM", "THEM", "", "", "", "", "<< macro >>", "<< macro >>")
        ws.Range("A1:K1").Font.Bold = True
        ws.Columns("A:K").ColumnWidth = 14
    End If
    ws.Range("A:F").NumberFormat = "@"
End Sub

Function TramChuanBiPOSua() As String
    EnsurePOSua
    ThisWorkbook.Sheets("PO_SUA").Range("A3:K3000").ClearContents
    TramChuanBiPOSua = "OK"
End Function

Function TramChayPOSua(ByVal luuThat As Boolean) As String
    TRAM_AUTO = True: TRAM_LUU_THAT = luuThat: TRAM_LOG = ""
    On Error Resume Next
    SuaPO_ME22N
    If Err.Number <> 0 Then TRAM_LOG = TRAM_LOG & "LOI VBA: " & Err.Description & vbLf
    On Error GoTo 0
    TRAM_AUTO = False
    TramChayPOSua = TRAM_LOG
End Function

' Doc text o (dong idx, cot col) bang item PO; tra ve Chr(1) neu khong cuon toi duoc (het bang)
Function POCellTextSafe(session As Object, ByVal idx As Long, ByVal col As Long) As String
    Dim c As Object
    On Error Resume Next
    Set c = PORowCell(session, idx, col)
    If Err.Number <> 0 Or c Is Nothing Then
        POCellTextSafe = Chr(1)
    Else
        POCellTextSafe = Trim(c.Text)
    End If
    Err.Clear
End Function

Private Sub POSuaStatus(ws As Worksheet, ByVal lastR As Long, ByVal st As String)
    Dim r As Long
    For r = 3 To lastR
        If Trim(CStr(ws.Cells(r, "A").Value)) <> "" Then ws.Cells(r, "J").Value = st
    Next r
End Sub

Sub SuaPO_ME22N()
    EnsurePOSua
    Dim ws As Worksheet, lastR As Long, r As Long, soPO As String, n As Long
    Set ws = ThisWorkbook.Sheets("PO_SUA")
    lastR = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastR < 3 Then Hoi "Sheet PO_SUA khong co dong nao.", vbExclamation, "SuaPO": Exit Sub
    soPO = Trim(CStr(ws.Cells(3, "A").Value))
    For r = 3 To lastR
        If Trim(CStr(ws.Cells(r, "A").Value)) <> "" Then
            n = n + 1
            If Trim(CStr(ws.Cells(r, "A").Value)) <> soPO Then Hoi "Sheet PO_SUA chi duoc chua 1 PO.", vbExclamation, "SuaPO": Exit Sub
        End If
    Next r
    If Hoi("Sua PO " & soPO & ": " & n & " dong. Bat dau nhap vao SAP?", vbYesNo + vbQuestion, "SuaPO " & MACRO_VERSION) = vbNo Then Exit Sub

    Dim session As Object
    Set session = GetSapSession()
    If session Is Nothing Then POSuaStatus ws, lastR, "LOI: khong ket noi duoc SAP": Exit Sub
    Set ctlCache = CreateObject("Scripting.Dictionary")
    gDecSep = SAP_DECIMAL_SEP

    Dim tbl As Object, f As Object, w0 As String, cTmp As Object
    Dim cItem As Long, cBanfn As Long, cBnfpo As Long, cMat As Long, cQty As Long, cPrice As Long
    Dim rowOf As Object, idxOf As Object, i As Long, t As String, nUsed As Long, idx As Long, eb As Long
    Set rowOf = CreateObject("Scripting.Dictionary")
    Set idxOf = CreateObject("Scripting.Dictionary")

    On Error GoTo LoiSua
    curStep = "mo ME22N"
    SapCuaSo session
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme22n"
    session.findById("wnd[0]").sendVKey 0
    CheckPopup session
    If UCase(session.Info.Transaction) <> "ME22N" Then Err.Raise vbObjectError + 901, "SuaPO", "SAP khong mo duoc ME22N (dang o '" & session.Info.Transaction & "')"

    curStep = "chon PO khac (Shift+F5) " & soPO
    session.findById("wnd[0]").sendVKey 17
    WaitSap session
    Set f = Nothing
    On Error Resume Next
    Set f = SearchCtl(session.findById("wnd[1]"), "MEPO_SELECT-EBELN")
    On Error GoTo LoiSua
    If f Is Nothing Then Err.Raise vbObjectError + 902, "SuaPO", "Khong thay o nhap so PO trong hop 'Other Purchase Order'"
    f.Text = soPO
    session.findById("wnd[1]").sendVKey 0
    WaitSap session
    CheckPopup session
    CheckSbarError session
    w0 = session.findById("wnd[0]").Text
    If InStr(w0, soPO) = 0 Then Err.Raise vbObjectError + 903, "SuaPO", "Khong mo duoc PO " & soPO & " (tieu de: " & w0 & ")"
    If InStr(UCase(w0), "DISPLAY") > 0 Or InStr(UCase(w0), "HIEN THI") > 0 Then
        curStep = "chuyen sang che do sua (Ctrl+F1)"
        session.findById("wnd[0]").sendVKey 25
        WaitSap session
        CheckPopup session
    End If

    On Error Resume Next
    SetItemDetail session, False
    Err.Clear
    On Error GoTo LoiSua

    curStep = "tim cot bang item PO"
    Set tbl = FindCtl(session, PO_TBL_ID)
    cItem = TblColIndex(tbl, "MEPO1211-EBELP")
    cBanfn = TblColIndex(tbl, "MEPO1211-BANFN")
    cBnfpo = TblColIndex(tbl, "MEPO1211-BNFPO")
    cMat = TblColIndex(tbl, "MEPO1211-EMATN")
    cQty = TblColIndex(tbl, "MEPO1211-MENGE")
    cPrice = TblColIndex(tbl, "MEPO1211-NETPR")
    If cItem < 0 Or cMat < 0 Or cQty < 0 Or cPrice < 0 Then Err.Raise vbObjectError + 904, "SuaPO", "Bang item ME22N thieu cot Item / Material / PO Quantity / Net Price"

    curStep = "doc cac dong PO hien co"
    For i = 0 To 998
        t = POCellTextSafe(session, i, cItem)
        If t = "" Or t = Chr(1) Then Exit For
        rowOf(CLng(Val(t))) = i
    Next i
    nUsed = i

    ' --- SUA so luong dong co san ---
    For r = 3 To lastR
        If UCase(Trim(CStr(ws.Cells(r, "B").Value))) = "SUA" Then
            eb = CLng(Val(ws.Cells(r, "C").Value))
            curStep = "sua dong PO " & eb
            If Not rowOf.Exists(eb) Then Err.Raise vbObjectError + 905, "SuaPO", "PO " & soPO & " khong co dong " & eb
            idx = rowOf(eb)
            t = POCellTextSafe(session, idx, cMat)
            If UCase(t) <> UCase(Trim(CStr(ws.Cells(r, "F").Value))) Then Err.Raise vbObjectError + 906, "SuaPO", "Dong " & eb & " tren SAP la ma " & t & ", khong phai " & ws.Cells(r, "F").Value
            Set cTmp = PORowCell(session, idx, cQty)
            cTmp.Text = SapNumber(CDbl(ws.Cells(r, "G").Value))
            session.findById("wnd[0]").sendVKey 0
            CheckPopup session
            CheckSbarError session
            idxOf(r) = idx
        End If
    Next r

    ' --- THEM dong tu PR bo sung ---
    For r = 3 To lastR
        If UCase(Trim(CStr(ws.Cells(r, "B").Value))) = "THEM" Then
            curStep = "them dong PR " & ws.Cells(r, "D").Value & "/" & ws.Cells(r, "E").Value & " (" & ws.Cells(r, "F").Value & ")"
            If Trim(CStr(ws.Cells(r, "D").Value)) = "" Or Trim(CStr(ws.Cells(r, "E").Value)) = "" Then Err.Raise vbObjectError + 907, "SuaPO", "Dong them thieu So PR / dong PR"
            If cBanfn < 0 Or cBnfpo < 0 Then Err.Raise vbObjectError + 908, "SuaPO", "Bang item ME22N khong hien cot Purch.Req./Item of requisition"
            idx = nUsed
            Set cTmp = PORowCell(session, idx, cBanfn)
            cTmp.Text = Trim(CStr(ws.Cells(r, "D").Value))
            Set cTmp = PORowCell(session, idx, cBnfpo)
            cTmp.Text = Trim(CStr(ws.Cells(r, "E").Value))
            session.findById("wnd[0]").sendVKey 0
            CheckPopup session
            Set cTmp = PORowCell(session, idx, cQty)
            cTmp.Text = SapNumber(CDbl(ws.Cells(r, "G").Value))
            Set cTmp = PORowCell(session, idx, cPrice)
            cTmp.Text = SapNumber(CDbl(ws.Cells(r, "H").Value))
            session.findById("wnd[0]").sendVKey 0
            CheckPopup session
            CheckSbarError session
            t = POCellTextSafe(session, idx, cMat)
            If UCase(t) <> UCase(Trim(CStr(ws.Cells(r, "F").Value))) Then Err.Raise vbObjectError + 909, "SuaPO", "Dong PR " & ws.Cells(r, "D").Value & "/" & ws.Cells(r, "E").Value & " la ma " & t & ", khong phai " & ws.Cells(r, "F").Value
            ws.Cells(r, "K").NumberFormat = "@"
            ws.Cells(r, "K").Value = POCellTextSafe(session, idx, cItem)
            idxOf(r) = idx
            nUsed = nUsed + 1
        End If
    Next r

    ' --- Doc lai, chinh gia neu can, dem lech ---
    Dim dq As String, dp As String, q As Double, pr As Double, expQ As Double, expP As Double
    Dim nDiff As Long, rep As String, netOK As Boolean, netErr As String
    For r = 3 To lastR
        If idxOf.Exists(r) Then
            idx = idxOf(r)
            expQ = CDbl(ws.Cells(r, "G").Value)
            expP = CDbl(ws.Cells(r, "H").Value)
            curStep = "doc lai dong " & (idx + 1)
            ReadPORow session, idx, cQty, cPrice, expQ, expP, dq, dp, q, pr
            If Abs(pr - expP) > 0.5 And expP > 0 Then
                netOK = TrySetNetPrice(session, idx, cPrice, expP, dp, netErr)
                ReadPORow session, idx, cQty, cPrice, expQ, expP, dq, dp, q, pr
                If Not netOK Or Abs(pr - expP) > 0.5 Then
                    SetPriceByPBXX session, idx, expP
                    ReadPORow session, idx, cQty, cPrice, expQ, expP, dq, dp, q, pr
                End If
            End If
            If Abs(q - expQ) > 0.0005 Or (expP > 0 And Abs(pr - expP) > 0.5) Then
                nDiff = nDiff + 1
                rep = rep & "  - " & ws.Cells(r, "F").Value & ": SAP " & dq & " x " & dp & " / can " & expQ & " x " & expP & vbCrLf
                ws.Cells(r, "J").Value = "LECH: SAP " & dq & " x " & dp
            End If
        End If
    Next r

    ' --- Luu ---
    Dim ans As VbMsgBoxResult, msg As String
    TramActivate
    ans = Hoi("Sua PO " & soPO & ": " & n & " dong." & vbCrLf & _
              IIf(nDiff > 0, "CANH BAO - lech " & nDiff & " dong:" & vbCrLf & rep, "SL va gia BANG TUYET DOI du lieu gui di." & vbCrLf) & vbCrLf & _
              "Luu (Save) PO nay?", vbYesNo + IIf(nDiff > 0, vbExclamation, vbQuestion) + vbDefaultButton2, "SuaPO " & MACRO_VERSION)
    If ans = vbNo Then
        For r = 3 To lastR
            If Trim(CStr(ws.Cells(r, "J").Value)) = "" And Trim(CStr(ws.Cells(r, "A").Value)) <> "" Then ws.Cells(r, "J").Value = "DA DIEN, CHUA LUU" & IIf(nDiff > 0, " (lech " & nDiff & " dong)", "")
        Next r
        Exit Sub
    End If
    ' v23.1-Tram-gl.2: Save qua SaveDocWithPopups (xu ly popup + canh bao W nhu PR/PO).
    ' Chi ghi DA LUU khi thanh trang thai la S va co dung so PO (hoac "No data was changed");
    ' truoc day ghi DA LUU ca khi SAP chi canh bao (W) va PO chua duoc luu.
    Dim soLuu As String
    soLuu = SaveDocWithPopups(session, "PO " & soPO, True)
    msg = session.findById("wnd[0]/sbar").Text
    If soLuu <> soPO And soLuu <> "(KHONG DOI)" Then
        Err.Raise vbObjectError + 910, "SuaPO", "Save xong nhung SAP khong bao da luu PO " & soPO & " (thanh trang thai: " & msg & ")"
    End If
    POSuaStatus ws, lastR, "DA LUU: " & msg
    Exit Sub

LoiSua:
    Dim e As String
    e = "LOI buoc [" & curStep & "]: " & Err.Description
    On Error Resume Next
    e = e & " | SAP: " & session.findById("wnd[0]/sbar").Text
    POSuaStatus ws, lastR, e
    Hoi e & vbCrLf & "Man hinh SAP giu nguyen.", vbCritical, "SuaPO " & MACRO_VERSION
End Sub

' ================================================================================
'  v23-Tram: MIGO cuoi ky qua tram (dung module MIGO_TungPhan v2-Tram)
' ================================================================================
Function TramChuanBiMIGO() As String
    Application.Run "EnsureMigoTP"
    With ThisWorkbook.Sheets("MIGO_TP")
        .Range("A3:I3000").ClearContents
        .Range("A3:C3000").NumberFormat = "@"
        .Range("E3:E3000").NumberFormat = "@"
    End With
    TramChuanBiMIGO = "OK"
End Function

Function TramChayMIGO(ByVal luuThat As Boolean) As String
    TRAM_AUTO = True: TRAM_LUU_THAT = luuThat: TRAM_LOG = ""
    On Error Resume Next
    Application.Run "PostMIGO_TungPhan"
    If Err.Number <> 0 Then TRAM_LOG = TRAM_LOG & "LOI VBA: " & Err.Description & vbLf
    On Error GoTo 0
    TRAM_AUTO = False
    TramChayMIGO = TRAM_LOG
End Function


' ================================================================================
'  v23-Tram: TAO FILE DNTT qua tram (dung module DNTT_Automation v6-Tram)
'  File mau DNTT.xlsx / DNTT-CCDC.xlsx dat CUNG thu muc voi file nay; ket qua o DNTT_xuat\
' ================================================================================
Function TramChuanBiDNTT() As String
    Dim ws As Worksheet, wc As Worksheet, p As String
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("DNTT")
    Set wc = ThisWorkbook.Sheets("DNTT_CauHinh")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "DNTT"
        ws.Range("A1:P1").Value = Array("STT", "Nha_cung_cap", "So_PO", "So_tien_PO", "Thang", "Ky", "Noi_dung_thanh_toan", "Ngay_lap", "Han_thanh_toan", "So_bien_ban", "Kiem_tra_ky", "Sheet_mau", "File_mau", "Nguoi_thu_huong", "File_DNTT_da_tao", "Status")
        ws.Range("A1:P1").Font.Bold = True
    End If
    ws.Range("A3:P3000").ClearContents
    ws.Range("A3:P3000").Interior.Pattern = xlNone
    ws.Range("C3:C3000").NumberFormat = "@"
    ' Tram: file mau + thu muc xuat lay o thu muc cua file nay (khong dung duong dan cu)
    If Not wc Is Nothing Then
        Dim r As Long, k As String
        For r = 4 To 17
            k = Trim(CStr(wc.Cells(r, 1).Value))
            Select Case k
                Case "File_mau_DNTT_1"
                    p = ThisWorkbook.Path & "\DNTT.xlsx"
                    If Dir(p) = "" Then p = ""
                    wc.Cells(r, 2).Value = p
                Case "File_mau_DNTT_2"
                    p = ThisWorkbook.Path & "\DNTT-CCDC.xlsx"
                    If Dir(p) = "" Then p = ""
                    wc.Cells(r, 2).Value = p
                Case "File_cong_no"
                    wc.Cells(r, 2).Value = ""
                Case "Thu_muc_xuat"
                    wc.Cells(r, 2).Value = ThisWorkbook.Path & "\DNTT_xuat"
            End Select
        Next r
    End If
    TramChuanBiDNTT = "OK"
End Function

Function TramChayDNTT() As String
    TRAM_AUTO = True: TRAM_LUU_THAT = True: TRAM_LOG = ""
    On Error Resume Next
    Application.Run "TaoDNTT"
    If Err.Number <> 0 Then TRAM_LOG = TRAM_LOG & "LOI VBA: " & Err.Description & vbLf
    On Error GoTo 0
    TRAM_AUTO = False
    TramChayDNTT = TRAM_LOG
End Function

' ================================================================================
'  v23-Tram: CHAY NGAM + NHIEU CUA SO SAP SONG SONG
' ================================================================================
' Chay tay: phong to cua so SAP nhu cu. Chay qua tram: thu nho (khong chiem man hinh).
Sub SapCuaSo(session As Object)
    On Error Resume Next
    If TRAM_AUTO Then
        If TRAM_AN_SAP Then session.findById("wnd[0]").iconify
    Else
        session.findById("wnd[0]").maximize
    End If
End Sub

' Dat cua so SAP + che do an cho lan chay tiep theo (goi truoc TramChayPR/PO/...)
Function TramCauHinh(ByVal phien As Long, ByVal anSap As Boolean) As String
    TRAM_PHIEN = phien
    TRAM_AN_SAP = anSap
    TramCauHinh = "OK"
End Function

' Mo them cua so SAP (session) cho du n cua so; thu nho cac cua so neu anSap. Tra ve so cua so dang co.
Function TramDamBaoPhien(ByVal n As Long, ByVal anSap As Boolean) As Long
    Dim eng As Object, conn As Object, t As Single, k As Long, c As Long
    On Error Resume Next
    ' SAP gia lap dang BAT: coi nhu da co du n cua so
    k = 0
    k = Application.Run("SapGiaLap_DamBaoPhien", n, anSap)
    Err.Clear
    If k > 0 Then TramDamBaoPhien = k: Exit Function
    Set eng = GetObject("SAPGUI").GetScriptingEngine
    If eng Is Nothing Then Exit Function
    If eng.Children.Count = 0 Then Exit Function
    Set conn = eng.Children(0)
    c = conn.Children.Count
    Do While c < n
        conn.Children(0).createSession
        t = Timer
        Do While conn.Children.Count <= c And Timer - t < 15
            DoEvents
        Loop
        If conn.Children.Count <= c Then Exit Do
        c = conn.Children.Count
    Loop
    If anSap Then
        For k = 0 To conn.Children.Count - 1
            conn.Children(k + 0).findById("wnd[0]").iconify
        Next k
    End If
    TramDamBaoPhien = conn.Children.Count
End Function
