VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SapGL_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================================
'  SAP GIA LAP - danh sach control (thay cho GuiComponentCollection cua SAP GUI)
'  Dung cho .Children / .Entries / .Columns / FindAllByName. Chi so bat dau tu 0
'  giong SAP GUI: .Count, .ElementAt(i), .Item(i)
' ================================================================================

Private mCol As Collection

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Public Sub Them(ByVal o As Object)
    mCol.Add o
End Sub

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Property Get Length() As Long
    Length = mCol.Count
End Property

Public Function ElementAt(ByVal i As Long) As Object
    If i < 0 Or i >= mCol.Count Then
        Err.Raise 614, "SAP GUI (gia lap)", "The enumerator of the collection cannot find an element with the specified index."
    End If
    Set ElementAt = mCol.Item(i + 1)
End Function

Public Function Item(ByVal i As Variant) As Object
    Set Item = ElementAt(CLng(i))
End Function

Public Function LaGiaLap() As Boolean
    LaGiaLap = True
End Function
