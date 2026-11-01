/**
 * TramSAP_API.gs – cổng nối giữa giao diện "Điều hành PR trước" và Trạm SAP (máy công ty).
 * Dán vào: file PRTruoc_DuLieu → Tiện ích mở rộng → Apps Script → thay toàn bộ Code.gs → Lưu.
 * Triển khai: Triển khai → Tùy chọn triển khai mới → Loại: Ứng dụng web
 *             Thực thi với tư cách: Tôi · Người có quyền truy cập: Bất kỳ ai → Triển khai → chép URL /exec.
 * Chỉ máy có đúng TOKEN mới gọi được. Không chia sẻ TOKEN.
 */
const TOKEN = 'DAN_MA_BAO_MAT_CUA_BAN';
const CH = 45000; // độ dài tối đa mỗi ô JSON (giống giao diện)
const XUAT_FOLDER = '1EUf-O4drs1z2U_RtJpvM16ODTnOW3kJ3'; // thư mục XuatSAP của dự án (file ĐNTT trạm tạo được chép lên đây)

function doGet(e)  { return handle_(e && e.parameter || {}); }
function doPost(e) { let b = {}; try { b = JSON.parse(e.postData.contents || '{}'); } catch (x) {} return handle_(b); }

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function handle_(p) {
  if (p.token !== TOKEN) return out_({ ok: false, loi: 'Sai token' });
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return out_({ ok: false, loi: 'Bận, thử lại' });
  try {
    ping_(p);
    if (p.a === 'ping') return out_({ ok: true });
    if (p.a === 'poll') return out_({ ok: true, lenh: takeNext_(p) });
    if (p.a === 'file') {
      const blob = Utilities.newBlob(Utilities.base64Decode(String(p.b64 || '')), String(p.mime || 'application/octet-stream'), String(p.name || 'file'));
      const f = DriveApp.getFolderById(XUAT_FOLDER).createFile(blob);
      return out_({ ok: true, id: f.getId(), url: f.getUrl() });
    }
    if (p.a === 'tiendo') { patch_(p.id, { tiendo: String(p.tiendo || '') }); return out_({ ok: true }); }
    if (p.a === 'ketqua') {
      patch_(p.id, { tt: p.tt === 'xong' ? 'xong' : 'loi', xongLuc: Date.now(), ketqua: p.ketqua || null, log: String(p.log || '').slice(0, 30000), tiendo: '' });
      return out_({ ok: true });
    }
    return out_({ ok: false, loi: 'Lệnh không rõ: ' + p.a });
  } catch (err) {
    return out_({ ok: false, loi: String(err) });
  } finally { lock.releaseLock(); }
}

function sh_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.getRange(1, 1, 1, 3).setValues([['id', 'cap_nhat', 'du_lieu']]); }
  return s;
}
function now_() { return Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss'); }

function ping_(p) {
  const c = CacheService.getScriptCache();
  if (p.a === 'poll' && c.get('pinged')) return;
  c.put('pinged', '1', 45);
  const s = sh_('tram');
  s.getRange(1, 1, 2, 5).setValues([
    ['may', 'ping_ts', 'ping', 'phien_ban', 'luu_that'],
    ["'" + String(p.may || ''), Date.now(), "'" + now_(), "'" + String(p.ver || ''), String(p.luu || '')]
  ]);
}

function rows_() {
  const s = sh_('lenh'), n = s.getLastRow();
  if (n < 2) return [];
  const w = Math.max(3, s.getLastColumn());
  return s.getRange(1, 1, n, w).getDisplayValues().map((r, i) => ({ row: i + 1, id: r[0], parts: r.slice(2).filter(x => x !== '') }))
    .filter(x => x.id && x.id !== 'id' && x.parts.length);
}
function read_(x) { try { return JSON.parse(x.parts.join('')); } catch (e) { return null; } }
function write_(row, id, doc) {
  const s = sh_('lenh'), json = JSON.stringify(doc), vals = ["'" + id, "'" + now_()];
  for (let i = 0; i < json.length; i += CH) vals.push("'" + json.slice(i, i + CH));
  const w = Math.max(vals.length, s.getLastColumn());
  while (vals.length < w) vals.push('');
  if (s.getMaxColumns() < w) s.insertColumnsAfter(s.getMaxColumns(), w - s.getMaxColumns());
  s.getRange(row, 1, 1, w).setValues([vals]);
}

function takeNext_(p) {
  const list = rows_().map(x => ({ x, d: read_(x) })).filter(o => o.d && o.d.tt === 'cho').sort((a, b) => (a.d.taoLuc || 0) - (b.d.taoLuc || 0));
  if (!list.length) return null;
  const o = list[0];
  o.d.tt = 'dang'; o.d.nhanLuc = Date.now(); o.d.may = String(p.may || '');
  write_(o.x.row, o.x.id, o.d);
  return Object.assign({ id: o.x.id }, o.d);
}

function patch_(id, ch) {
  const x = rows_().find(r => r.id === id);
  if (!x) throw new Error('Không thấy lệnh ' + id);
  const d = read_(x) || {};
  Object.assign(d, ch);
  write_(x.row, id, d);
}
