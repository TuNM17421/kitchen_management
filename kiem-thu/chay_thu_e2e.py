import sys, os; sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from e2e import *
import datetime
shutil.rmtree(RUN, ignore_errors=True)
OUT = {}
def lenh(title, kb, fn, dbf='gl_db.xls'):
    doc = fresh()
    R(doc, 'TestGL', 'Setup', dbf, kb)
    try:
        res = fn(doc)
    finally:
        log = R(doc, 'TestGL', 'LogText')
        dbr = db(doc)
        doc.close(True)
    show(title + ' / ket qua', res)
    show(title + ' / DB sau lenh', dbr)
    OUT[title] = (res, log)
    return res, log, dbr

CN = [
 [1, 'HD001', 'NCC A', 'VT001', 'Hang 1', 10, 50000, 'Ớt(RAU)', 'PR-T11-NV'],
 [2, 'HD001', 'NCC A', 'VT002', 'Hang 2', 5.5, 120000, 'Ớt(RAU)', 'PR-T11-NV'],
 [3, 'HD002', 'NCC B', 'VT003', 'Hang 3', 3, 80000, 'Chanh', 'PR-T11-NV'],
 [4, 'HD002', 'NCC B', 'VT004', 'Hang 4', 2, 30000, 'Chanh', 'PR-T11-NV'],
]
BS = [[1, 'HD003', 'NCC A', 'VT005', 'Hang 5', 4, 25000, 'Ớt(RAU)', 'PR-T11-BS']]

def cmd_pr(rows):
    def f(doc):
        R(doc, 'SAP_Automation', 'TramCauHinh', 0, True)
        R(doc, 'SAP_Automation', 'TramChuanBi')
        fill(doc, 'CongNo', rows); doc.calculateAll()
        lg = R(doc, 'SAP_Automation', 'TramChayPR', True)
        return read(doc, 'PR', 'L3:P%d' % (2 + len(rows))) + [lg[-300:]]
    return f

# 1. PR
r1, _, _ = lenh('1 PR', '', cmd_pr(CN))
soPR = str(int(r1[0][0])); items = [r[4] for r in r1[:4]]

# 2. PO
def f_po(doc):
    R(doc, 'SAP_Automation', 'TramCauHinh', 0, True)
    R(doc, 'SAP_Automation', 'TramChuanBi')
    fill(doc, 'CongNo', CN); doc.calculateAll()
    R(doc, 'SAP_Automation', 'TramDatMaNCC', 'NCC A', '2000001')
    R(doc, 'SAP_Automation', 'TramDatMaNCC', 'NCC B', '2000002')
    for i in range(4):
        R(doc, 'SAP_Automation', 'TramGanPR', 3 + i, soPR, items[i])
    doc.calculateAll()
    lg = R(doc, 'SAP_Automation', 'TramChayPO', True)
    return read(doc, 'PO', 'B3:K4') + [lg[-400:]]
r2, _, _ = lenh('2 PO', '', f_po)
poA = [r for r in r2[:2] if r[0] == 'NCC A'][0][8]; poB = [r for r in r2[:2] if r[0] == 'NCC B'][0][8]
poA = str(int(poA)) if isinstance(poA, float) else poA
poB = str(int(poB)) if isinstance(poB, float) else poB
print('PO A', poA, 'PO B', poB)

# 3. PR bo sung
r3, _, _ = lenh('3 PR bo sung', '', cmd_pr(BS))
soBS = str(int(r3[0][0])); itBS = r3[0][4]

def posua(rows, luu):
    def f(doc):
        R(doc, 'SAP_Automation', 'TramChuanBiPOSua')
        fill(doc, 'PO_SUA', rows)
        lg = R(doc, 'TestGL', 'ChayPOSua', luu)
        return read(doc, 'PO_SUA', 'A3:K%d' % (2 + len(rows))) + [lg[-600:]]
    return f
SUA = [[poA, 'SUA', '10', '', '', 'VT001', 12, 50000, 'Hang 1'],
       [poA, 'THEM', '', soBS, itBS, 'VT005', 4, 25000, 'Hang 5']]
# 4a. ME22N chay thu (khong luu)
lenh('4a ME22N khong luu', '', posua(SUA, False))
# 4b. ME22N: PO mo o che do hien thi (kich ban dong 33) + luu that
r4, log4, _ = lenh('4b ME22N luu that + mo o che do hien thi', '33', posua(SUA, True))
print('\n'.join(l for l in log4.split('\n') if 'Ctrl+F1' in l or 'Doi che do' in l or 'Mo PO' in l or 'SAP bao' in l))
# 4c. PO dang bi khoa
lenh('4c ME22N PO bi khoa', '34', posua([[poB, 'SUA', '10', '', '', 'VT003', 4, 80000, 'Hang 3']], True))
# 4d. canh bao W khi Save
r4d, log4d, db4d = lenh('4d ME22N canh bao W khi Save', '35', posua([[poB, 'SUA', '10', '', '', 'VT003', 4, 80000, 'Hang 3']], True))

# 4e. SAP bao loi E khi Save -> macro phai ghi LOI, khong duoc ghi DA LUU
lenh('4e ME22N loi E khi Save', '36', posua([[poB, 'SUA', '10', '', '', 'VT003', 3, 80000, 'Hang 3']], True))
def migo(rows, luu, phien=0):
    def f(doc):
        R(doc, 'SAP_Automation', 'TramCauHinh', phien, True)
        R(doc, 'SAP_Automation', 'TramChuanBiMIGO')
        fill(doc, 'MIGO_TP', rows)
        lg = R(doc, 'TestGL', 'ChayMIGO', luu)
        return read(doc, 'MIGO_TP', 'A3:I%d' % (2 + len(rows))) + [lg[-500:]]
    return f
# 5a. MIGO chay thu
M1 = [['P01', poA, '25.10.2026', '10', 'VT001', 4, 'Nhan ky 1'], ['P01', poA, '25.10.2026', '20', 'VT002', 5.5, 'Nhan ky 1']]
lenh('5a MIGO khong luu', '', migo(M1, False))
# 5b. MIGO luu that, nhap 1 phan dong 10, bo dong 30 (dong them tu ME22N)
r5, log5, _ = lenh('5b MIGO luu that (nhap 1 phan)', '', migo(M1, True))
# 5c. Nhan vuot SL con mo (dong 10 con 8, nhap 9)
lenh('5c MIGO nhan vuot', '', migo([['P02', poA, '26.10.2026', '10', 'VT001', 9, 'Nhan ky 2']], True))
# 5d. Ngay sai dinh dang
lenh('5d MIGO ngay sai', '', migo([['P03', poA, '2026-10-26', '10', 'VT001', 1, 'x']], True))
# 5e. Hai phieu, cua so SAP thu 2 (so chung tu dai rieng)
lenh('5e MIGO 2 phieu, cua so 2', '', migo([['P04', poA, '27.10.2026', '10', 'VT001', 8, 'Nhan ky 2'],
                                          ['P05', poA, '27.10.2026', '30', 'VT005', 4, 'Nhan ky 2']], True, 1))
# 7. Sheet CongNo bi chen 1 dong o tren (lech dong so voi sheet PR) -> CNRow phai doc dung dong
def chen_dong_cn(doc):
    sheet(doc, 'CongNo').Rows.insertByIndex(2, 1)    # chen dong 3: du lieu CongNo xuong dong 4..7
    sheet(doc, 'CongNo').getCellRangeByName('B3').setString('(dong ghi chu ai do chen vao)')
    doc.calculateAll()
def cmd_pr_lech(rows):
    def f(doc):
        R(doc, 'SAP_Automation', 'TramCauHinh', 0, True)
        R(doc, 'SAP_Automation', 'TramChuanBi')
        fill(doc, 'CongNo', rows); chen_dong_cn(doc)
        lg = R(doc, 'SAP_Automation', 'TramChayPR', True)
        return [[sheet(doc, 'PR').getCellRangeByName('C3').getFormula()]] + read(doc, 'PR', 'L3:P%d' % (2 + len(rows))) + [lg[-300:]]
    return f
r7, _, _ = lenh('7a PR, CongNo lech 1 dong', '', cmd_pr_lech(CN), 'gl_db7.xls')
soPR7 = str(int(r7[1][0])); items7 = [r[4] for r in r7[1:5]]
def f_po_lech(doc):
    R(doc, 'SAP_Automation', 'TramCauHinh', 0, True)
    R(doc, 'SAP_Automation', 'TramChuanBi')
    fill(doc, 'CongNo', CN); chen_dong_cn(doc)
    R(doc, 'SAP_Automation', 'TramDatMaNCC', 'NCC A', '2000001')
    R(doc, 'SAP_Automation', 'TramDatMaNCC', 'NCC B', '2000002')
    for i in range(4):
        R(doc, 'SAP_Automation', 'TramGanPR', 3 + i, soPR7, items7[i])
    doc.calculateAll()
    lg = R(doc, 'SAP_Automation', 'TramChayPO', True)
    return read(doc, 'PO', 'B3:K4') + [lg[-400:]]
lenh('7b PO, CongNo lech 1 dong', '', f_po_lech, 'gl_db7.xls')

# 6. TramDamBaoPhien
doc = fresh(); R(doc, 'TestGL', 'Setup', 'gl_db.xls', ''); show('6 TramDamBaoPhien(3)', R(doc, 'TestGL', 'DamBao', 3)); doc.close(True)
json.dump({k: [v[0], v[1]] for k, v in OUT.items()}, open('/tmp/gltest/t2.json', 'w'), ensure_ascii=False, indent=1, default=str)
