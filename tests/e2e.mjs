// 실제 브라우저에서 샘플 파일을 올려 처음부터 내려받기까지 확인한다.
// 실행: npm run e2e  (Playwright와 Chromium이 필요하다)
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';
import { readFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

const require = createRequire(import.meta.url);
const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const XLSX = require(join(root, 'app/vendor/xlsx.full.min.js'));
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const shots = process.env.SHOTS_DIR || mkdtempSync(join(tmpdir(), 'om-shots-'));

const PORT = 8765;
const server = spawn('python3', ['-m', 'http.server', String(PORT), '--bind', '127.0.0.1', '--directory', join(root, 'app')], { stdio: 'ignore' });
await new Promise(r => setTimeout(r, 800));
const BASE = `http://127.0.0.1:${PORT}/index.html`;
// 파일 경로 대신 내용을 넘긴다(일부 환경에서 Playwright가 한글 경로 파일을 넘기지 못한다).
const sample = n => ({ name: n, mimeType: 'application/octet-stream', buffer: readFileSync(join(root, 'samples', n)) });
const SAMPLE_FILES = ['스마트스토어_주문.xlsx', '쿠팡_DeliveryList.xlsx', '기타몰_주문내역.xlsx', '자사몰_주문.csv'].map(sample);

let failed = false;
// 로캘이 없는 환경에서는 Chromium이 한글 파일 이름을 'download'로 바꿔 버려 UTF-8 로캘을 지정한다.
const browser = await chromium.launch({ env: { ...process.env, LANG: 'C.UTF-8', LC_ALL: 'C.UTF-8' } });
try {
  const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1200, height: 900 } });
  const page = await ctx.newPage();
  const problems = [];
  page.on('pageerror', e => problems.push('pageerror: ' + e.message));
  page.on('console', m => { if (m.type() === 'error') problems.push('console: ' + m.text()); });
  page.on('request', r => { if (!r.url().startsWith(`http://127.0.0.1:${PORT}/`) && !r.url().startsWith('blob:') && !r.url().startsWith('data:')) problems.push('external request: ' + r.url()); });

  await page.goto(BASE);
  await page.setInputFiles('#file-input', SAMPLE_FILES);
  await page.waitForSelector('#s2:not(.hidden)');
  const badges = await page.$$eval('#file-list .badge', els => els.map(e => e.textContent.trim()));
  assert.deepEqual(badges, ['✓ 스마트스토어', '✓ 쿠팡', '추측함 · 확인 필요', '열 연결 필요']);
  assert.equal(await page.textContent('#s2-orders'), '12');
  assert.match(await page.textContent('#s2-notices'), /같은 주문 1건/);
  await page.screenshot({ path: join(shots, 'desk-s2.png'), fullPage: true });

  // 자사몰 CSV(EUC-KR)의 열을 직접 연결한다.
  await page.click('#file-list li:nth-child(4) [data-map]');
  await page.waitForSelector('#s21:not(.hidden)');
  assert.match(await page.textContent('#map-preview'), /오세린/); // 글자가 깨지지 않았는지
  assert.equal(await page.inputValue('#m-phone'), '핸드폰');
  for (const [k, v] of Object.entries({ orderNo: '코드', receiver: '고객', address: '배송처', product: '아이템', qty: '개수', memo: '비고' })) {
    await page.selectOption('#m-' + k, v);
  }
  assert.equal(await page.inputValue('#map-market'), '자사몰'); // 파일 이름에서 미리 채워진다
  await page.screenshot({ path: join(shots, 'desk-s21.png'), fullPage: true });
  await page.click('#s21-ok');
  await page.waitForSelector('#s2:not(.hidden)');
  assert.equal(await page.textContent('#s2-orders'), '14');

  await page.click('#s2-next');
  await page.waitForSelector('#s3:not(.hidden)');
  assert.equal(await page.textContent('#s3-boxes'), '8');
  assert.equal(await page.textContent('#s3-groups'), '6');
  assert.match(await page.textContent('#tabs'), /기타몰 2/);

  // 상품명 줄임말 설정
  await page.click('#alias-open');
  const aliasInputs = await page.$$('#alias-list input');
  const before = await page.textContent('#alias-notice');
  await aliasInputs[0].fill('현미 1kg');
  await aliasInputs[0].dispatchEvent('change');
  assert.notEqual(await page.textContent('#alias-notice'), before);
  await page.screenshot({ path: join(shots, 'desk-s3.png'), fullPage: true });

  // 한 주문 빼기 → 박스 수 감소(오세린은 단독 주문)
  const row = page.locator('#order-table tr', { hasText: '오세린' });
  await row.locator('input[type="checkbox"]').uncheck();
  assert.equal(await page.textContent('#s3-boxes'), '7');

  await page.click('#s3-next');
  await page.waitForSelector('#s4:not(.hidden)');
  assert.equal(await page.textContent('#ex-boxes'), '7');
  await page.screenshot({ path: join(shots, 'desk-s4.png'), fullPage: true });

  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('[data-ex="courier"]')]);
  assert.match(dl.suggestedFilename(), /^택배업로드_\d{4}-\d{2}-\d{2}\.xlsx$/);
  const wb = XLSX.read(readFileSync(await dl.path()), { type: 'buffer' });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false });
  assert.equal(rows.length, 1 + 7);
  const seojun = rows.find(r => r[0] === '이서준');
  assert.equal(seojun[1], '010-9876-1234');
  assert.match(seojun[4], /현미 1kg x1/);
  assert.match(seojun[8], /쿠팡:31000124/);
  assert.ok(!rows.some(r => r[0] === '오세린'));

  const [dl2] = await Promise.all([page.waitForEvent('download'), page.click('[data-ex="all"]')]);
  const wb2 = XLSX.read(readFileSync(await dl2.path()), { type: 'buffer' });
  assert.equal(XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[0]], { header: 1 }).length, 1 + 13);

  // 새로 열어도 열 연결이 기억된다.
  await page.goto(BASE);
  await page.setInputFiles('#file-input', [sample('자사몰_주문.csv')]);
  await page.waitForSelector('#s2:not(.hidden)');
  assert.equal(await page.textContent('#file-list .badge'), '✓ 자사몰');

  // 잘못된 파일
  await page.click('#s2-back');
  await page.setInputFiles('#file-input', [{ name: '깨진파일.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from('not an excel file') }]);
  await page.waitForSelector('#s2:not(.hidden)');
  const badText = await page.textContent('#file-list');
  assert.match(badText, /읽을 수 없음|주문표 없음/);

  assert.deepEqual(problems, []);
  console.log('desktop flow ok');

  // 휴대폰 화면: 가로 스크롤 없음, 카드 목록 표시
  const mob = await browser.newPage({ viewport: { width: 390, height: 844 } });
  const mobProblems = [];
  mob.on('pageerror', e => mobProblems.push(e.message));
  await mob.goto(BASE);
  await mob.click('#btn-sample');
  await mob.waitForSelector('#s2:not(.hidden)');
  await mob.screenshot({ path: join(shots, 'mob-s2.png'), fullPage: true });
  await mob.click('#s2-next');
  await mob.waitForSelector('#order-cards .ocard');
  await mob.screenshot({ path: join(shots, 'mob-s3.png'), fullPage: true });
  assert.ok(await mob.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth));
  assert.deepEqual(mobProblems, []);
  console.log('mobile flow ok');

  // 파일을 더블클릭해 여는 경우(file://)에도 동작하는지
  const local = await browser.newPage();
  const localProblems = [];
  local.on('pageerror', e => localProblems.push(e.message));
  local.on('console', m => { if (m.type() === 'error') localProblems.push(m.text()); });
  await local.goto('file://' + join(root, 'app/index.html'));
  await local.click('#btn-sample');
  await local.waitForSelector('#s2:not(.hidden)', { timeout: 3000 });
  assert.deepEqual(localProblems, []);
  console.log('file:// ok');
  console.log('screenshots:', shots);
} catch (e) {
  failed = true;
  console.error(e);
} finally {
  await browser.close();
  server.kill();
}
process.exit(failed ? 1 : 0);
