VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SapGL_Ctl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================================
'  SAP GIA LAP - 1 doi tuong SAP GUI Scripting (session, cua so, o nhap, nut,
'  combobox, tab, grid ALV, table control, o cua bang, thanh cuon, status bar...)
'  Cac macro goi y nhu goi SAP that: .findById, .Text, .key, .press, .sendVKey,
'  .modifyCell, .GetCell, .VerticalScrollbar.Position ...
'  Moi xu ly nam o module SapGiaLap (ham GL_...).
'
'  Rieng 3 ten .Select / .Type / .Close VBA KHONG cho dat lam ten ham, nen
'  macro goi qua SapSelect / SapType / SapClose (module SAP_Automation), cac ham
'  nay chuyen sang ChonTab / Loai / DongCuaSo khi dang gia lap.
' ================================================================================

Public Kind As String       ' SESSION, CONN, APP, INFO, WND, OKCD, SBAR, BTN, CMB, FLD, TAB, GRID, TBL, CELL, SCROLL, COLS, ENTRY, LBL, CONT
Public mId As String
Public mName As String
Public mType As String
Public mText As String
Public mKey As String
Public mRef As String        ' du lieu gan voi control (vd "PO:VENDOR", "ACC:KOSTL", ten bang...)
Public Hang As Long         ' o bang: dong TUYET DOI trong du lieu
Public HangHien As Long     ' o bang: dong tren man hinh (0..VisibleRowCount-1)
Public Cot As Long          ' o bang: chi so cot
Public Tooltip As String
Public mChangeable As Boolean
Public Kids As SapGL_List
Public Cha As Object

Private Sub Class_Initialize()
    Set Kids = New SapGL_List
    mChangeable = True
End Sub

' ---------------------------------------------------------------- nhan dang
Public Property Get LaGiaLap() As Boolean
    LaGiaLap = True
End Property

Public Property Get Id() As String
    Id = mId
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Get Loai() As String          ' = .Type cua SAP
    Loai = mType
End Property

Public Property Get Parent() As Object
    Set Parent = Cha
End Property

Public Function Children(Optional ByVal idx As Variant) As Object
    If IsMissing(idx) Then
        Set Children = GL_Children(Me)
    Else
        Set Children = GL_Children(Me).ElementAt(CLng(idx))
    End If
End Function

' ---------------------------------------------------------------- tim control
Public Function findById(ByVal sId As String, Optional ByVal bRaiseArg As Variant) As Object
    Dim bRaise As Boolean
    bRaise = True
    If Not IsMissing(bRaiseArg) Then bRaise = CBool(bRaiseArg)
    Set findById = GL_FindById(Me, sId, bRaise)
End Function

Public Function FindByName(ByVal sName As String, ByVal sType As String) As Object
    Set FindByName = GL_FindByName(Me, sName, sType)
End Function

Public Function FindAllByName(ByVal sName As String, ByVal sType As String) As Object
    Set FindAllByName = GL_FindAllByName(Me, sName, sType)
End Function

' ---------------------------------------------------------------- noi dung
Public Property Get Text() As String
    Text = GL_GetText(Me)
End Property

Public Property Let Text(ByVal v As String)
    GL_SetText Me, v
End Property

Public Property Get Value() As String
    Value = GL_GetText(Me)
End Property

Public Property Get Key() As String
    Key = GL_GetKey(Me)
End Property

Public Property Let Key(ByVal v As String)
    GL_SetKey Me, v
End Property

Public Property Get Entries() As Object
    Set Entries = GL_Entries(Me)
End Property

Public Property Get Changeable() As Boolean
    Changeable = GL_Changeable(Me)
End Property

Public Property Get Selected() As Boolean
    Selected = GL_GetSelected(Me)
End Property

Public Property Let Selected(ByVal v As Boolean)
    GL_SetSelected Me, v
End Property

Public Property Get IconName() As String
    IconName = GL_IconName(Me)
End Property

Public Property Get MessageType() As String
    MessageType = GL_MessageType(Me)
End Property

' ---------------------------------------------------------------- session / info
Public Property Get Busy() As Boolean
    Busy = False
End Property

Public Property Get Info() As Object
    Set Info = GL_Info()
End Property

Public Property Get Transaction() As String
    Transaction = GL_InfoVal("TRANSACTION")
End Property

Public Property Get SystemName() As String
    SystemName = GL_InfoVal("SYSTEMNAME")
End Property

Public Property Get Client() As String
    Client = GL_InfoVal("CLIENT")
End Property

Public Property Get User() As String
    User = GL_InfoVal("USER")
End Property

Public Property Get Program() As String
    Program = GL_InfoVal("PROGRAM")
End Property

Public Property Get MajorVersion() As Long
    MajorVersion = 7
End Property

Public Property Get MinorVersion() As Long
    MinorVersion = 70
End Property

' ---------------------------------------------------------------- hanh dong
Public Sub maximize()
    GL_Log "maximize", Me, ""
End Sub

Public Sub iconify()                          ' thu nho cua so (tram chay ngam)
    GL_Log "iconify", Me, ""
End Sub

Public Sub restore()
    GL_Log "restore", Me, ""
End Sub

Public Sub sendVKey(ByVal vKey As Long)
    GL_SendVKey Me, vKey
End Sub

Public Sub press()
    GL_Press Me
End Sub

Public Sub SetFocus()
    GL_SetFocus Me
End Sub

Public Sub ChonTab()                          ' = .Select cua SAP (tab)
    GL_SelectTab Me
End Sub

Public Sub DongCuaSo()                        ' = .Close cua SAP (cua so)
    GL_CloseWnd Me
End Sub

' ---------------------------------------------------------------- grid ALV (ME51N)
Public Sub modifyCell(ByVal Row As Long, ByVal Column As String, ByVal v As String)
    GL_ModifyCell Me, Row, Column, v
End Sub

Public Function getCellValue(ByVal Row As Long, ByVal Column As String) As String
    getCellValue = GL_GetCellValue(Me, Row, Column)
End Function

Public Property Let currentCellColumn(ByVal v As String)
    mKey = v
    GL_Log "currentCellColumn", Me, v
End Property

Public Property Get currentCellColumn() As String
    currentCellColumn = mKey
End Property

Public Property Get currentCellRow() As Long
    currentCellRow = Hang
End Property

Public Property Let currentCellRow(ByVal v As Long)
    Hang = v
End Property

Public Sub setCurrentCell(ByVal Row As Long, ByVal Column As String)
    Hang = Row
    mKey = Column
End Sub

Public Sub pressEnter()
    GL_GridEnter Me
End Sub

' ---------------------------------------------------------------- table control (ME21N / ME23N / MIGO / Conditions)
Public Property Get RowCount() As Long
    RowCount = GL_RowCount(Me)
End Property

Public Property Get VisibleRowCount() As Long
    VisibleRowCount = GL_VisibleRows(Me)
End Property

Public Property Get VerticalScrollbar() As Object
    Set VerticalScrollbar = GL_Scrollbar(Me)
End Property

Public Property Get Columns() As Object
    Set Columns = GL_Columns(Me)
End Property

Public Function GetCell(ByVal Row As Long, ByVal Column As Long) As Object
    Set GetCell = GL_GetCell(Me, Row, Column)
End Function

' ---------------------------------------------------------------- thanh cuon
Public Property Get Position() As Long
    Position = GL_ScrollPos(Me)
End Property

Public Property Let Position(ByVal v As Long)
    GL_SetScrollPos Me, v
End Property

Public Property Get Maximum() As Long
    Maximum = GL_ScrollMax(Me)
End Property

Public Property Get Minimum() As Long
    Minimum = 0
End Property

Public Property Get PageSize() As Long
    PageSize = GL_VisibleRows(Cha)
End Property
