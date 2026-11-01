# Bien doi code VBA CHI DE CHAY THU trong LibreOffice (khong co Scripting.Dictionary)
import re
def find_close(s, i):
    d = 0; q = False
    while i < len(s):
        ch = s[i]
        if ch == '"': q = not q
        elif not q:
            if ch == '(': d += 1
            elif ch == ')':
                d -= 1
                if d == 0: return i
        i += 1
    return -1

def xform(src):
    # LibreOffice: End(xlUp) tinh theo sheet dang chon -> thay bang ham quet (chi khi chay thu)
    src = re.sub(r'(\w+)\.Cells\(\1\.Rows\.Count, ([^)]+)\)\.End\(xlUp\)\.Row', r'GLLastRow(\1, \2)', src)
    # LibreOffice khong co hang vbObjectError
    src = src.replace('vbObjectError', '1000')
    names = set(re.findall(r'Set (\w+) = CreateObject\("Scripting\.Dictionary"\)', src))
    src = src.replace('CreateObject("Scripting.Dictionary")', 'New GLDict')
    # LibreOffice: sau 'On Error Resume Next', 'On Error GoTo Nhan' khong bat lai duoc loi -> chen 'On Error GoTo 0' truoc
    src = re.sub(r'(?m)^(\s*)On Error GoTo (?!0\b)(\w+)', r'\1On Error GoTo 0: On Error GoTo \2', src)
    out = []
    for ln in src.split('\n'):
        # LibreOffice khong co FormatConditions (chi to mau) -> bo khi chay thu
        if '.FormatConditions' in ln or ln.strip().startswith('fc.Interior') or 'As FormatCondition' in ln:
            out.append("' [LO] " + ln); continue
        for nm in names:
            pos = 0
            while True:
                m = re.compile(r'(?<![\w.])' + nm + r'\(').search(ln, pos)
                if not m: break
                st = m.start(); op = m.end() - 1
                cl = find_close(ln, op)
                if cl < 0: break
                args = ln[op+1:cl]
                pre = ln[:st]
                rest = ln[cl+1:]
                if rest.startswith(' = ') and (pre.strip() == '' or pre.rstrip().endswith('Then') or pre.rstrip().endswith(':')):
                    head = pre + nm + '.SetItem ' + args + ', '
                    ln = head + rest[3:]
                    pos = len(head)
                    continue
                rep = nm + '.Item(' + args + ')'
                if rest.startswith('('):
                    c2 = find_close(rest, 0)
                    rep += '.Item(' + rest[1:c2] + ')'
                    rest = rest[c2+1:]
                ln = pre + rep + rest
                pos = len(pre) + len(rep)
        out.append(ln)
    return '\n'.join(out)

GLDICT = '''Option VBASupport 1
Option ClassModule
Private mK As Collection
Private mV As Collection
Private mKeys() As Variant
Private mN As Long
Private Sub Class_Initialize()
    Set mK = New Collection
    Set mV = New Collection
    mN = 0
End Sub
Private Function KS(k As Variant) As String
    KS = "k" & CStr(k)
End Function
Public Function Exists(k As Variant) As Boolean
    Dim x As Variant
    On Error GoTo No
    x = mK.Item(KS(k))
    Exists = True
    Exit Function
No:
    Exists = False
End Function
Public Sub Add(k As Variant, v As Variant)
    If Exists(k) Then Err.Raise 457, "GLDict", "This key is already associated with an element of this collection"
    mN = mN + 1
    ReDim Preserve mKeys(1 To mN)
    mKeys(mN) = k
    mK.Add mN, KS(k)
    mV.Add v, KS(k)
End Sub
Public Sub SetItem(k As Variant, v As Variant)
    If Exists(k) Then
        mV.Remove KS(k)
        mV.Add v, KS(k)
    Else
        Add k, v
    End If
End Sub
Public Function Item(k As Variant) As Variant
    If Not Exists(k) Then
        Item = Empty
        Exit Function
    End If
    If IsObject(mV.Item(KS(k))) Then
        Set Item = mV.Item(KS(k))
    Else
        Item = mV.Item(KS(k))
    End If
End Function
Public Property Get Count() As Long
    Count = mV.Count
End Property
Public Function Keys() As Variant
    Dim a() As Variant, i As Long, n As Long
    If mV.Count = 0 Then Keys = Array(): Exit Function
    ReDim a(0 To mV.Count - 1)
    For i = 1 To mN
        If Exists(mKeys(i)) Then
            If mK.Item(KS(mKeys(i))) = i Then a(n) = mKeys(i): n = n + 1
        End If
    Next i
    Keys = a
End Function
Public Sub Remove(k As Variant)
    mK.Remove KS(k)
    mV.Remove KS(k)
End Sub
Public Sub RemoveAll()
    Set mK = New Collection
    Set mV = New Collection
    mN = 0
End Sub
'''
