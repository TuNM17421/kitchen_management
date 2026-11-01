# Chay thu E2E qua cac ham Tram* (giong TramSAP.ps1): moi lenh mo ban sao MOI cua Template_PR_Tram.xlsm,
# cai module, chay, dong KHONG luu. DB gia lap dung chung qua file ngoai (o B15).
import sys, os, shutil, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lo import *
from xform import xform, GLDICT
PRJ = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
RUN = '/tmp/gltest/run'
TPL = PRJ + '/excel-mau/Template_PR_Tram.xlsm'

TESTGL = '''Option VBASupport 1
Function Setup(ByVal dbFile As String, ByVal kb As String) As String
    Dim ws As Worksheet, r As Long, a As Variant, x As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("SAP_GiaLap")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "SAP_GiaLap"
    End If
    ws.Range("B3").Value = "BAT"
    Set x = SapGiaLap_PhienNeuBat()
    ws.Range("B15").Value = dbFile
    For r = 19 To 40
        ws.Cells(r, 1).Value = ""
    Next r
    If kb <> "" Then
        For Each a In Split(kb, ",")
            ws.Cells(CLng(a), 1).Value = "x"
        Next a
    End If
    SapGiaLap_DatLai
    SapGiaLap_XoaLog
    Setup = "OK " & GL_VERSION & " / " & TramPhienBan()
End Function
Function LogText() As String
    Dim ws As Worksheet, r As Long, s As String, n As Long
    Set ws = ThisWorkbook.Worksheets("SAP_GiaLap_Log")
    n = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 3 To n
        s = s & ws.Cells(r, 3).Value & " | " & ws.Cells(r, 4).Value & " | " & ws.Cells(r, 5).Value & " | " & ws.Cells(r, 6).Value & " | " & ws.Cells(r, 7).Value & vbLf
    Next r
    LogText = s
End Function
Function Sess() As String
    Dim o As Object
    Set o = GetSapSession()
    Sess = o.Id & " | " & ThisWorkbook.Name & " | DBrows=" & ThisWorkbook.Worksheets("SAP_GiaLap_DB").Cells(ThisWorkbook.Worksheets("SAP_GiaLap_DB").Rows.Count, 1).End(xlUp).Row
End Function
Function GLLastRow(ws As Object, col As Variant) As Long
    Dim r As Long, c As Long
    c = ws.Cells(1, col).Column
    r = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    Do While r > 1
        If Trim(CStr(ws.Cells(r, c).Value)) <> "" Then Exit Do
        r = r - 1
    Loop
    GLLastRow = r
End Function
' Goi thang macro (khong qua TramChay* co On Error Resume Next: LibreOffice xu ly
' Resume Next cua ham goi khac Excel - nuot loi ben trong SuaPO_ME22N)
Function ChayPOSua(ByVal luu As Boolean) As String
    TRAM_AUTO = True: TRAM_LUU_THAT = luu: TRAM_LOG = ""
    SuaPO_ME22N
    TRAM_AUTO = False
    ChayPOSua = TRAM_LOG
End Function
Function ChayMIGO(ByVal luu As Boolean) As String
    TRAM_AUTO = True: TRAM_LUU_THAT = luu: TRAM_LOG = ""
    PostMIGO_TungPhan
    TRAM_AUTO = False
    ChayMIGO = TRAM_LOG
End Function
Function DamBao(ByVal n As Long) As Long
    DamBao = TramDamBaoPhien(n, True)
End Function
'''

ctx = start()

def fresh(name='Template_PR_Tram.xlsm'):
    os.makedirs(RUN, exist_ok=True)
    dst = os.path.join(RUN, name)
    shutil.copy(TPL, dst)
    doc = load(ctx, dst)
    lib = doc.BasicLibraries.getByName('VBAProject')
    for n in list(lib.getElementNames()):
        if not (n.startswith('Sheet') or n == 'ThisWorkbook'):
            lib.removeByName(n)
    mods = [
        ('SAP_Automation', xform(vba_src(PRJ + '/tram-sap/SAP_Automation_v23_Tram.bas')), False),
        ('MIGO_TungPhan', xform(vba_src(PRJ + '/tram-sap/MIGO_TungPhan_v2_Tram.bas')), False),
        ('SapGiaLap', xform(vba_src(PRJ + '/gia_lap/SapGiaLap.bas')), False),
        ('SapGL_Ctl', vba_src(PRJ + '/gia_lap/SapGL_Ctl.cls', True), True),
        ('SapGL_Dong', vba_src(PRJ + '/gia_lap/SapGL_Dong.cls', True), True),
        ('SapGL_List', vba_src(PRJ + '/gia_lap/SapGL_List.cls', True), True),
        ('GLDict', GLDICT, True),
        ('TestGL', TESTGL, False),
    ]
    for name, src, cls in mods:
        lib.insertByName(name, src)
    return doc

def R(doc, mod, fn, *a):
    return run(doc, 'VBAProject', mod, fn, a)

def sheet(doc, nm):
    return doc.Sheets.getByName(nm)

def fill(doc, nm, rows):
    ws = sheet(doc, nm)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = ws.getCellByPosition(j, 2 + i)
            if v is None: c.setString('')
            elif isinstance(v, (int, float)): c.setValue(v)
            else: c.setString(str(v))

def read(doc, nm, rng):
    ws = sheet(doc, nm)
    return [[c for c in r] for r in ws.getCellRangeByName(rng).getDataArray()]

def db(doc):
    ws = sheet(doc, 'SAP_GiaLap_DB')
    out = []
    for r in ws.getCellRangeByName('A2:P60').getDataArray():
        if r[0]: out.append(list(r[:11]) + [r[15]])
    return out

def show(title, x):
    print('\n=== ' + title)
    if isinstance(x, list):
        for r in x: print('   ', r)
    else:
        print('   ' + str(x).replace('\n', '\n    '))
