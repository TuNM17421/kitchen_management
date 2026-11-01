VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SapGL_Dong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================================
'  SAP GIA LAP - 1 dong item (PR ME51N / PO ME21N-ME23N / dong MIGO)
'  Chi la noi chua du lieu, logic nam o module SapGiaLap.
' ================================================================================

' --- chung ---
Public Mat As String          ' ma vat tu
Public Menge As Double        ' so luong
Public Loi As String          ' loi SAP dang treo tren dong (rong = khong loi)
Public DaDoi As Boolean       ' dong vua bi sua tu lan Enter truoc
Public DaDoiGia As Boolean    ' vua sua SL / gia / dieu kien gia
' --- PR (ME51N) ---
Public Name1 As String
Public Preis As Double
Public Dispo As String
Public Ekorg As String
Public Ekgrp As String
Public Knttp As String
Public Kostl As String
Public Geber As String
Public Bnfpo As String
Public DaGoGia As Boolean     ' da go don gia (PREIS)
' --- PO (ME21N / ME23N) ---
Public Banfn As String        ' so PR go vao cot Purch.Req.
Public BanfnItem As String    ' dong item PR go vao cot Item of requisition
Public Ebelp As String
Public Netpr As Double
Public Peinh As Double
Public Loekz As String
Public DaLay As Boolean       ' da lay du lieu tu dong PR
Public KhoaGia As Boolean     ' o Net Price khong cho sua (gia lay tu dieu kien)
Public Txz01 As String
Public GiaPR As Double        ' don gia tren dong PR goc
Public SLNhap As Double       ' ME22N: so luong da nhap kho cua dong PO
Public SLPR As Double         ' so luong tren dong PR goc
' --- MIGO ---
Public SoPO As String
Public PoItem As String
Public Take As Boolean
Public Sgtxt As String
Public Meins As String
Public SLMo As Double         ' so luong con mo (con nhan duoc) khi mo PO trong MIGO

' --- bang Conditions cua dong PO ---
Private mK() As String
Private mA() As Double
Private mN As Long

Private Sub Class_Initialize()
    Peinh = 1
    Meins = "KG"
    mN = 0
End Sub

Public Function CondCount() As Long
    CondCount = mN
End Function

Public Sub CondThem(ByVal k As String, ByVal a As Double)
    mN = mN + 1
    ReDim Preserve mK(1 To mN)
    ReDim Preserve mA(1 To mN)
    mK(mN) = k
    mA(mN) = a
End Sub

Public Function CondK(ByVal i As Long) As String
    If i >= 1 And i <= mN Then CondK = mK(i)
End Function

Public Function CondA(ByVal i As Long) As Double
    If i >= 1 And i <= mN Then CondA = mA(i)
End Function

Public Sub CondDatK(ByVal i As Long, ByVal k As String)
    If i >= 1 And i <= mN Then mK(i) = k
End Sub

Public Sub CondDatA(ByVal i As Long, ByVal a As Double)
    If i >= 1 And i <= mN Then mA(i) = a
End Sub

Public Sub CondXoa(ByVal i As Long)
    Dim j As Long
    If i < 1 Or i > mN Then Exit Sub
    For j = i To mN - 1
        mK(j) = mK(j + 1)
        mA(j) = mA(j + 1)
    Next j
    mN = mN - 1
    If mN > 0 Then
        ReDim Preserve mK(1 To mN)
        ReDim Preserve mA(1 To mN)
    End If
End Sub
