"""Soạn gói dữ liệu test (thực đơn + định mức + lượng khách) cho Bảng điều hành, tháng 11/2026."""
import json, random, sys, datetime as dt, unicodedata, re
DATA = sys.argv[1]; OUT = sys.argv[2]
IT = {x[0]: x for x in json.load(open(DATA, encoding='utf8'))['items']}
rnd = random.Random(20261101)

def slug(s):
    s = unicodedata.normalize('NFD', str(s).lower())
    s = re.sub('[̀-ͯ]', '', s).replace('đ', 'd')
    return re.sub('[^a-z0-9]+', '-', s).strip('-')[:120] or 'mon'

# món -> [(mã, định lượng / suất theo đơn vị tính của mã)]
DM = {
 'Thịt kho trứng':        [('THP00423',0.08),('THP07943',1),('THP01229',0.004)],
 'Gà rang gừng':          [('THP00332',0.12),('THP01821',0.005),('THP01229',0.003)],
 'Cá basa sốt cà chua':   [('THP03146',0.10),('THP00155',0.04),('THP06600',0.003)],
 'Sườn ram mặn ngọt':     [('THP00399',0.12),('THP01821',0.005),('THP01229',0.003)],
 'Bò xào súp lơ':         [('THP00374',0.06),('THP02912',0.06),('THP01829',0.003)],
 'Đậu phụ sốt cà chua':   [('THP02258',0.12),('THP00155',0.05),('THP06600',0.003)],
 'Tôm rim thịt':          [('THP01901',0.05),('THP00423',0.04),('THP01229',0.003)],
 'Cá diêu hồng chiên xù': [('THP00493',0.10)],
 'Chả lá lốt':            [('THP02654',0.08),('THP01821',0.004)],
 'Mực xào cần tỏi':       [('THP00547',0.07),('THP01829',0.004),('THP02052',0.03)],
 'Gà luộc lá chanh':      [('THP02965',0.12)],
 'Trứng đúc thịt':        [('THP07943',2),('THP02654',0.03)],
 'Canh bí đao nấu tôm':   [('THP00149',0.08),('THP06722',0.01)],
 'Canh cải ngọt thịt băm':[('THP00225',0.06),('THP02654',0.015)],
 'Canh chua cá':          [('THP03860',0.05),('THP00155',0.03),('THP00301',0.01)],
 'Canh rau ngót':         [('THP00230',0.06),('THP02654',0.01)],
 'Canh khoai tây cà rốt': [('THP00195',0.05),('THP00159',0.03),('THP00352',0.03)],
 'Rau muống xào tỏi':     [('THP00242',0.10),('THP01829',0.004)],
 'Bắp cải luộc':          [('THP00213',0.10)],
 'Cải chíp xào nấm':      [('THP00218',0.08),('THP00132',0.02)],
 'Su su xào':             [('THP03644',0.10),('THP01829',0.003)],
 'Bí ngồi xào':           [('THP00146',0.10)],
 'Chuối tiêu':            [('THP00626',1)],
 'Dưa hấu':               [('THP00660',0.15)],
 'Thanh long':            [('THP00693',0.15)],
 'Cam sành':              [('THP00637',0.15)],
 'Sữa chua không đường':  [('THP01122',1)],
 'Phở bò':                [('THP03014',0.15),('THP00374',0.05),('THP06600',0.005),('THP00381',0.05)],
 'Bún gà':                [('THP00102',0.15),('THP02965',0.06),('THP06600',0.005)],
 'Bánh mì trứng':         [('THP00066',1),('THP07943',1)],
 'Cháo thịt bằm':         [('THP08254',0.04),('THP02654',0.04)],
 'Bánh cuốn chả':         [('THP00088',0.20),('THP00114',0.03)],
 'Súp gà ngô':            [('THP02036',0.05),('THP00159',0.02),('THP07943',0.5)],
}
for m, nl in DM.items():
    for ma, _ in nl:
        assert ma in IT, (m, ma)

MAN = ['Thịt kho trứng','Gà rang gừng','Cá basa sốt cà chua','Sườn ram mặn ngọt','Bò xào súp lơ','Đậu phụ sốt cà chua',
       'Tôm rim thịt','Cá diêu hồng chiên xù','Chả lá lốt','Mực xào cần tỏi','Gà luộc lá chanh','Trứng đúc thịt']
CANH = ['Canh bí đao nấu tôm','Canh cải ngọt thịt băm','Canh chua cá','Canh rau ngót','Canh khoai tây cà rốt']
RAU = ['Rau muống xào tỏi','Bắp cải luộc','Cải chíp xào nấm','Su su xào','Bí ngồi xào']
TM = ['Chuối tiêu','Dưa hấu','Thanh long','Cam sành']
SANG = ['Phở bò','Bún gà','Bánh mì trứng','Cháo thịt bằm','Bánh cuốn chả','Súp gà ngô']

def iso(d): return d.isoformat()
def week_items(nhom, mon, k):
    """k = chỉ số tuần để xoay vòng món."""
    items = []
    for i in range(7):
        d = iso(mon + dt.timedelta(days=i)); s = k * 7 + i
        if nhom == 'NV':
            for j, b in enumerate(['Trưa', 'Tối']):
                t = s * 2 + j
                for x in range(3):
                    items.append({'d': d, 'b': b, 'loai': 'Mặn', 'mon': MAN[(t * 3 + x) % len(MAN)], 'tl': 67})
                items += [{'d': d, 'b': b, 'loai': 'Canh', 'mon': CANH[t % 5]},
                          {'d': d, 'b': b, 'loai': 'Rau', 'mon': RAU[(t + 2) % 5]},
                          {'d': d, 'b': b, 'loai': 'Tráng miệng', 'mon': TM[t % 4]}]
        else:
            items += [{'d': d, 'b': 'Sáng', 'loai': 'Chọn 1', 'mon': SANG[s % 6]},
                      {'d': d, 'b': 'Sáng', 'loai': 'Chọn 1', 'mon': SANG[(s + 3) % 6]},
                      {'d': d, 'b': 'Sáng', 'loai': 'Tráng miệng', 'mon': TM[(s + 1) % 4]}]
            for j, b in enumerate(['Trưa', 'Tối']):
                t = s * 2 + j + 5
                items += [{'d': d, 'b': b, 'loai': 'Mặn', 'mon': MAN[t % len(MAN)]},
                          {'d': d, 'b': b, 'loai': 'Canh', 'mon': CANH[t % 5]},
                          {'d': d, 'b': b, 'loai': 'Rau', 'mon': RAU[(t + 1) % 5]},
                          {'d': d, 'b': b, 'loai': 'Tráng miệng', 'mon': TM[t % 4]}]
            items.append({'d': d, 'b': 'Phụ', 'loai': '', 'mon': 'Sữa chua không đường'})
    return items

NAM, THANG = 2026, 11
first = dt.date(NAM, THANG, 1); last = dt.date(NAM, THANG + 1, 1) - dt.timedelta(days=1)
mon0 = first - dt.timedelta(days=first.weekday())
docs = {}
k = 0; w = mon0
while w <= last:
    for g in ('NV', 'NTC'):
        docs[f'thucdon/{g}_{iso(w)}'] = {'nhom': g, 'tuan': iso(w), 'items': week_items(g, w, k), 'nguon': 'Gói test 11/2026'}
    w += dt.timedelta(days=7); k += 1

NV_TB = {1:[181,30],2:[160,27],3:[156,30],4:[189,27],5:[185,30],6:[112,26],0:[38,25]}
NT = {'TH':[26,29,24,22],'VIP':[1,2,2,3],'AN':[10,11,11,22]}
def jit(v): return max(0, round(v * rnd.uniform(0.85, 1.15)))
days = {}
d = first
while d <= last:
    wd = (d.weekday() + 1) % 7
    days[iso(d)] = [jit(x) for x in NV_TB[wd] + NT['TH'] + NT['VIP'] + NT['AN']]
    d += dt.timedelta(days=1)

dinhmuc = {slug(m): {'mon': m, 'nl': [{'ma': a, 'q': q} for a, q in nl], 'nguon': 'Excel tuần 1', 'ghichu': 'gói test 11/2026'} for m, nl in DM.items()}
P = {'ten': 'Gói test tháng 11/2026 (dữ liệu giả)', 'docs': docs, 'dinhmuc': dinhmuc, 'khach': {'thang': f'{NAM}-{THANG:02d}', 'days': days}}
json.dump(P, open(OUT, 'w', encoding='utf8'), ensure_ascii=False, indent=1)
print('tuan', k, 'docs', len(docs), 'mon', len(dinhmuc), 'ngay', len(days))
