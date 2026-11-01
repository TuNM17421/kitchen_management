// Kiểm thử trang trong trình duyệt (Playwright): tính dự trù và phân bổ PR trên gói test phải ra đúng số.
// Chạy: node tests/web/smoke.cjs   (cần đã build: python3 web/build.py && python3 web/build.py --demo)
const { chromium } = require('playwright');
const path = require('node:path');

const ROOT = path.resolve(__dirname, '..', '..');
const url = f => 'file://' + path.join(ROOT, 'web', 'dist', f);
let fail = 0;
const check = (ok, msg) => { console.log((ok ? 'OK   ' : 'FAIL ') + msg); if (!ok) fail++; };

(async () => {
const browser = await chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {});
async function open(file) {
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
  const errs = [];
  page.on('pageerror', e => errs.push(e.message));
  // Thư viện xlsx từ CDN không cần cho các bước này: chặn để test chạy được khi không có mạng.
  await page.route('https://cdnjs.cloudflare.com/**', r => r.fulfill({ status: 200, contentType: 'text/javascript', body: '' }));
  await page.goto(url(file));
  return { page, errs };
}

{ // Bản thường: chạy thử, không lỗi JS
  const { page, errs } = await open('dieu-hanh.html');
  await page.waitForTimeout(1500);
  check(await page.isVisible('#previewNote'), 'bản thường: hiện thông báo chạy thử');
  check((await page.innerText('#conn')).includes('Chạy thử'), 'bản thường: trạng thái "Chạy thử"');
  check(errs.length === 0, 'bản thường: không lỗi JS ' + JSON.stringify(errs));
  await page.close();
}
{ // Bản demo: tự nạp gói test, tính dự trù, phân bổ PR
  const { page, errs } = await open('dieu-hanh-demo.html');
  await page.waitForFunction(() => /Lần tính gần nhất/.test(document.body.innerText), null, { timeout: 15000 });
  const t = await page.innerText('body');
  check(/Nạp gói “Gói test tháng 11\/2026/.test(t), 'demo: đã nạp gói test');
  check(/Lần tính gần nhất: 10\.042 suất, trong đó 10\.042 suất có thực đơn \+ định mức/.test(t), 'demo: 10.042 suất đều có thực đơn + định mức');
  check(t.includes('684.546.740 đ'), 'demo: ngân sách dự tính 684.546.740 đ');
  check(/188 mã hàng · 18 NCC/.test(t), 'demo: 188 mã hàng, 18 NCC');
  await page.getByText('Lập PR', { exact: true }).first().click();
  await page.waitForTimeout(800);
  const t2 = await page.innerText('body');
  check(/Mã Nhân viên \(be1\)\s*350,4 tr/.test(t2), 'demo: PR mã Nhân viên 350,4 tr');
  check(/Mã Bệnh nhân \(be2\)\s*334,1 tr/.test(t2), 'demo: PR mã Bệnh nhân 334,1 tr');
  check(errs.length === 0, 'demo: không lỗi JS ' + JSON.stringify(errs));
  await page.close();
}
await browser.close();
console.log(fail ? `\n${fail} kiểm tra lỗi` : '\nTất cả kiểm tra đạt');
process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
