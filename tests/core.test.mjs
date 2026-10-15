import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const XLSX = require(join(root, 'app/vendor/xlsx.full.min.js'));
for (const f of ['fields', 'markets', 'normalize', 'decode', 'parse', 'merge', 'exporters', 'workbook']) {
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
