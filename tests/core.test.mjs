import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const XLSX = require(join(root, 'app/vendor/xlsx.full.min.js'));
for (const f of ['fields', 'markets', 'normalize', 'decode', 'parse', 'merge', 'exporters', 'workbook', 'tracking']) {
  await import(join(root, `app/js/core/${f}.js`));
}
const { OM } = globalThis;

function sheetOf(name) {
  return OM.readWorkbook(XLSX, name, readFileSync(join(root, 'samples', name))).sheets[0].aoa;
}

test('정규화: 휴대폰 번호 앞자리 0 복원, 주소 시·도 표기 통일', () => {
  assert.equal(OM.norm.phone('1055550101'), '01055550101');
  assert.equal(OM.norm.phone('010-5555-0101'), '01055550101');
  assert.equal(OM.norm.phoneDisplay('01098761234'), '010-9876-1234');
  assert.equal(OM.norm.address('서울특별시 마포구 월드컵로 10, 101동'), OM.norm.address('서울 마포구 월드컵로 10 101동'));
  assert.equal(OM.norm.address('경기도 성남시 분당구 판교역로 20'), OM.norm.address('경기 성남시 분당구 판교역로 20'));
  assert.equal(OM.norm.address('부산광역시 해운대구 센텀중앙로 5 (우동)'), OM.norm.address('부산 해운대구 센텀중앙로 5'));
  assert.equal(OM.norm.qty('2개'), 2);
  assert.equal(OM.norm.qty(''), null);
});

test('스마트스토어 파일을 인식하고 연락처1이 비면 연락처2를 쓴다', () => {
  const p = OM.parse.parseSheet(sheetOf('스마트스토어_주문.xlsx'), {});
  assert.equal(p.status, 'ok');
  assert.equal(p.market, '스마트스토어');
  assert.equal(p.orders.length, 6);
  const yuna = p.orders.find(o => o.receiver === '최유나');
  assert.equal(yuna.phone, '010-3141-5926');
  // 구매자 정보가 아니라 수취인 정보를 읽어야 한다.
  assert.equal(p.orders[1].receiver, '이서준');
});

test('쿠팡 파일을 인식하고 한 주문번호의 여러 상품 줄을 모두 읽는다', () => {
  const p = OM.parse.parseSheet(sheetOf('쿠팡_DeliveryList.xlsx'), {});
  assert.equal(p.status, 'ok');
  assert.equal(p.market, '쿠팡');
  assert.equal(p.orders.length, 5);
  assert.equal(p.orders.filter(o => o.orderNo === '31000125').length, 2);
});

test('제목 줄이 있는 기타 형식은 머리글 줄을 찾아 열 이름으로 추측한다', () => {
  const p = OM.parse.parseSheet(sheetOf('기타몰_주문내역.xlsx'), {});
  assert.equal(p.status, 'guessed');
  assert.equal(p.headers[0], '주문No'); // 제목 줄을 건너뛰고 머리글을 찾는다
  assert.deepEqual(
    { orderNo: p.mapping.orderNo, receiver: p.mapping.receiver, phone: p.mapping.phone, address: p.mapping.address, product: p.mapping.product, qty: p.mapping.qty },
    { orderNo: '주문No', receiver: '수령인', phone: '수령인 휴대폰', address: '배송지 주소', product: '품명', qty: '주문수량' },
  );
  assert.equal(p.orders.length, 2);
});

test('EUC-KR CSV는 글자가 깨지지 않고, 모르는 형식은 열 연결 후 읽힌다', () => {
  const aoa = sheetOf('자사몰_주문.csv');
  const p = OM.parse.parseSheet(aoa, {});
  assert.equal(p.status, 'unknown');
  assert.equal(p.headers[2], '고객');
  assert.ok(p.missing.includes('주문번호'));
  const mapping = { orderNo: '코드', receiver: '고객', phone: '핸드폰', address: '배송처', product: '아이템', qty: '개수', memo: '비고' };
  const done = OM.parse.applyMapping(aoa, p, '자사몰', mapping);
  assert.equal(done.status, 'ok');
  assert.equal(done.orders.length, 2);
  assert.equal(done.orders[1].memo, '경비실');

  // 저장한 매핑은 같은 양식 파일에 자동 적용된다.
  const again = OM.parse.parseSheet(aoa, { [p.signature]: { market: '자사몰', mapping } });
  assert.equal(again.status, 'ok');
  assert.equal(again.market, '자사몰');
});

function loadAll() {
  const files = ['스마트스토어_주문.xlsx', '쿠팡_DeliveryList.xlsx', '기타몰_주문내역.xlsx', '자사몰_주문.csv'].map(n => {
    const aoa = sheetOf(n);
    let parsed = OM.parse.parseSheet(aoa, {});
    if (parsed.status === 'unknown') {
      parsed = OM.parse.applyMapping(aoa, parsed, '자사몰', { orderNo: '코드', receiver: '고객', phone: '핸드폰', address: '배송처', product: '아이템', qty: '개수', memo: '비고' });
    }
    return { name: n, parsed };
  });
  return OM.merge.combine(files);
}

test('합치기: 중복 줄 제거, 마켓을 넘나드는 합포장 묶기', () => {
  const { orders, duplicates } = loadAll();
  assert.equal(duplicates, 1);
  assert.equal(orders.length, 14);
  const g = OM.merge.group(orders);
  assert.equal(g.groups, 6);
  assert.equal(g.boxes.length, 8);
  const seojun = orders.filter(o => o.receiver === '이서준');
  assert.equal(new Set(seojun.map(o => g.groupNo[o.id])).size, 1);
  assert.notEqual(g.groupNo[seojun[0].id], undefined);
});

test('묶음 풀기와 보내지 않기를 반영한다', () => {
  const { orders } = loadAll();
  const minji = orders.filter(o => o.receiver === '김민지');
  const g = OM.merge.group(orders, new Set([orders.find(o => o.receiver === '오세린').id]), new Set([minji[0].id]));
  assert.equal(g.groups, 5);
  assert.equal(g.boxes.length, 8); // 오세린 빠짐(-1), 김민지 분리(+1)
});

test('내보내기: 택배 업로드는 박스 단위, 피킹 리스트는 줄임말 기준 합계', () => {
  const { orders } = loadAll();
  const g = OM.merge.group(orders);
  const aliases = {};
  for (const o of orders) if (/현미/.test(o.product)) aliases[OM.merge.itemKey(o)] = '현미 1kg';
  const courier = OM.exporters.courierRows(g.boxes, aliases);
  assert.equal(courier.length, 1 + 8);
  const seojun = courier.find(r => r[0] === '이서준');
  assert.equal(seojun[1], '010-9876-1234');
  assert.match(seojun[4], /현미 1kg x1 \/ 잡곡 혼합 10곡/);
  assert.equal(seojun[5], 2);
  assert.match(seojun[8], /스마트스토어:2026101512346/);
  assert.match(seojun[8], /쿠팡:31000124/);

  const pick = OM.exporters.pickList(orders, new Set(), aliases);
  assert.deepEqual(pick[0], ['현미 1kg', 8]);
  const total = pick.reduce((s, [, q]) => s + q, 0);
  assert.equal(total, orders.reduce((s, o) => s + o.qty, 0));
});

test('암호가 걸린 파일은 encrypted 오류로 알려 준다', () => {
  const cfb = XLSX.CFB.utils.cfb_new();
  XLSX.CFB.utils.cfb_add(cfb, '/EncryptionInfo', Buffer.from([4, 0, 4, 0, 0x40, 0, 0, 0]));
  XLSX.CFB.utils.cfb_add(cfb, '/EncryptedPackage', Buffer.alloc(16));
  const bytes = XLSX.CFB.write(cfb, { type: 'buffer' });
  assert.throws(() => OM.readWorkbook(XLSX, '잠긴파일.xlsx', bytes), e => e.code === 'encrypted');
});

test('큰 숫자 주문번호가 지수 표기로 바뀌지 않는다', () => {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['상품주문번호', '수취인명'], [2026101512345678, '홍길동']]), 'S');
  const book = OM.readWorkbook(XLSX, 'n.xlsx', XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
  assert.equal(book.sheets[0].aoa[1][0], '2026101512345678');
});

// ---------- 2단계: 송장번호 되돌려 넣기 ----------

function loadFiles() {
  return ['스마트스토어_주문.xlsx', '쿠팡_DeliveryList.xlsx', '기타몰_주문내역.xlsx'].map(n => ({ name: n, sheet: { aoa: sheetOf(n) }, parsed: OM.parse.parseSheet(sheetOf(n), {}) }));
}

test('택배사 결과: 원주문번호로 먼저 찾고, 없으면 이름+가려진 전화번호 뒤 4자리로 찾는다', () => {
  const files = loadFiles();
  const { orders } = OM.merge.combine(files);
  const g = OM.merge.group(orders);
  const result = sheetOf('택배사_송장결과.xlsx');
  const m = OM.tracking.matchResult(g.boxes, result);
  assert.equal(m.found, true);
  assert.equal(m.rows, 7);
  assert.equal(m.unmatched.length, 0);
  assert.equal(m.assigned.size, 7);
  // 원주문번호가 없는 줄(정하은)도 이름+전화 뒤 4자리로 찾았다.
  const haeun = g.boxes.find(l => l[0].receiver === '정하은');
  assert.equal(m.assigned.get(OM.tracking.boxKey(haeun)), '689412345698');
});

test('택배사 결과: 찾지 못한 줄과 같은 박스에 다른 송장이 온 경우를 알려 준다', () => {
  const { orders } = OM.merge.combine(loadFiles());
  const g = OM.merge.group(orders);
  const aoa = [
    ['운송장번호', '받는분', '받는분 전화번호', '원주문번호'],
    ['1111-2222-3333', '김민지', '010-2233-5678', ''],
    ['999', '김민지', '010-2233-5678', ''],
    ['555', '모르는사람', '010-0000-0000', ''],
    ['', '송장 없음', '010-1111-1111', ''],
  ];
  const m = OM.tracking.matchResult(g.boxes, aoa);
  assert.equal(m.rows, 3);
  const minji = g.boxes.find(l => l[0].receiver === '김민지');
  assert.equal(m.assigned.get(OM.tracking.boxKey(minji)), '111122223333');
  assert.equal(m.conflicts.length, 1);
  assert.equal(m.unmatched.length, 1);
  assert.equal(m.unmatched[0].receiver, '모르는사람');
});

test('송장 열이 없는 파일은 결과를 찾지 못했다고 알린다', () => {
  const m = OM.tracking.matchResult([], [['이름', '전화'], ['a', 'b']]);
  assert.equal(m.found, false);
});

test('원본 파일에 송장 채우기: 마켓 열에 넣고, 없으면 열을 추가하고, 묶음의 모든 줄과 중복 줄에도 넣는다', () => {
  const files = loadFiles();
  const { orders } = OM.merge.combine(files);
  const excluded = new Set([orders.find(o => o.receiver === '윤재희').id]);
  const g = OM.merge.group(orders, excluded);
  const m = OM.tracking.matchResult(g.boxes, sheetOf('택배사_송장결과.xlsx'));
  const byLine = OM.tracking.trackingByLine(g.boxes, m.assigned);

  // 쿠팡: 기존 '택배사'·'운송장번호' 열에 넣는다.
  const cp = files[1];
  const outCp = OM.tracking.fillFile(cp.sheet.aoa, cp.parsed, byLine, 'CJ대한통운');
  const h = outCp.aoa[0];
  assert.equal(h.length, cp.sheet.aoa[0].length);
  const doyun = outCp.aoa.filter(r => r[12] === '한도윤');
  assert.equal(doyun.length, 2);
  assert.ok(doyun.every(r => r[h.indexOf('운송장번호')] === '689412345705' && r[h.indexOf('택배사')] === 'CJ대한통운'));
  assert.equal(outCp.filled, 5);

  // 스마트스토어: 열이 없으면 맨 뒤에 추가. 중복 줄에도 같은 송장.
  const ss = files[0];
  const outSs = OM.tracking.fillFile(ss.sheet.aoa, ss.parsed, byLine, 'CJ대한통운');
  const hs = outSs.aoa[0];
  assert.deepEqual(hs.slice(-2), ['택배사', '송장번호']);
  const t = hs.indexOf('송장번호');
  assert.equal(outSs.aoa[1][t], outSs.aoa[6][t]);
  assert.equal(outSs.filled, 6);

  // 보내지 않기로 한 주문은 비워 둔다.
  const etc = files[2];
  const outEtc = OM.tracking.fillFile(etc.sheet.aoa, etc.parsed, byLine, 'CJ대한통운');
  const he = outEtc.aoa[etc.parsed.headerRow];
  const jaehee = outEtc.aoa.find(r => r.includes('윤재희'));
  assert.equal(jaehee[he.indexOf('송장번호')] || '', '');
  assert.equal(outEtc.filled, 1);
  // 원본 표는 바뀌지 않는다.
  assert.equal(ss.sheet.aoa[0].length, 14);
});
