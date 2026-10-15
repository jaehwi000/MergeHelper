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
console.log('samples written to', out);

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
