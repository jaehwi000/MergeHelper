// app/js/samples.js의 샘플 주문으로 samples/ 폴더에 실제 xlsx·csv 파일을 만든다.
// 실행: npm run samples
import { createRequire } from 'node:module';
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const XLSX = require(join(root, 'app/vendor/xlsx.full.min.js'));
await import(join(root, 'app/js/samples.js'));
const out = join(root, 'samples');
mkdirSync(out, { recursive: true });

for (const s of globalThis.OM.SAMPLES) {
  if (s.csv) {
    // 국내 쇼핑몰 CSV에 흔한 EUC-KR로 저장해 인코딩 처리를 함께 검증한다.
    writeFileSync(join(out, s.name), encodeEucKr(s.csv));
  } else {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s.aoa), 'Sheet1');
    writeFileSync(join(out, s.name), XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
  }
}
// 2단계 체험용: 택배사에서 받는 결과 파일 흉내. 택배사마다 열 이름이 달라서 일부러 다른 이름을 쓰고,
// 원주문번호는 앞쪽 절반에만 남겨 이름·연락처로 찾는 경우도 함께 보여 준다. 전화번호는 가운데를 가린다.
{
  const OM = globalThis.OM;
  for (const f of ['fields', 'markets', 'normalize', 'decode', 'parse', 'merge', 'exporters', 'workbook']) {
    await import(join(root, `app/js/core/${f}.js`));
  }
  const files = OM.SAMPLES.map(s => {
    const book = s.csv
      ? OM.readWorkbook(XLSX, s.name, new TextEncoder().encode(s.csv))
      : OM.readWorkbook(XLSX, s.name, XLSX.write(sheetBook(s.aoa), { type: 'buffer', bookType: 'xlsx' }));
    return OM.parseWorkbook(book, {});
  });
  const { orders } = OM.merge.combine(files);
  const { boxes } = OM.merge.group(orders);
  const rows = [['운송장번호', '접수일자', '받는분성명', '받는분전화번호', '받는분주소', '품목명', '고객주문번호']];
  boxes.forEach((list, i) => {
    const o = list[0];
    const phone = OM.norm.phoneDisplay(o.phone).replace(/^(\d{3})-(\d{3,4})-/, '$1-****-');
    const ref = i < boxes.length / 2 ? list.map(x => `${x.market}:${x.orderNo}`).join(' ') : '';
    rows.push([`6894${String(12345670 + i * 7).padStart(8, '0')}`, '2026-10-15', o.receiver, phone, o.address, list.map(x => x.product).join(', '), ref]);
  });
  writeFileSync(join(out, '택배사_송장결과.xlsx'), XLSX.write(sheetBook(rows), { type: 'buffer', bookType: 'xlsx' }));
}

console.log('samples written to', out);

function sheetBook(aoa) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Sheet1');
  return wb;
}

// Node에는 EUC-KR 인코더가 없어서 디코더로 역방향 표를 만든다.
function encodeEucKr(text) {
  const dec = new TextDecoder('euc-kr');
  const map = new Map();
  for (let a = 0x81; a <= 0xfe; a++) {
    for (let b = 0x41; b <= 0xfe; b++) {
      const ch = dec.decode(new Uint8Array([a, b]));
      if (ch.length === 1 && ch !== '�' && !map.has(ch)) map.set(ch, [a, b]);
    }
  }
  const bytes = [];
  for (const ch of text) {
    const c = ch.codePointAt(0);
    if (c < 0x80) bytes.push(c);
    else if (map.has(ch)) bytes.push(...map.get(ch));
    else throw new Error('EUC-KR로 바꿀 수 없는 글자: ' + ch);
  }
  return Uint8Array.from(bytes);
}
