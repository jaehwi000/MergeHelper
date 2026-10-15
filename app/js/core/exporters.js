// 내려받을 표(2차원 배열)를 만든다. 파일로 쓰는 일은 화면 코드가 SheetJS로 한다.
(function (OM) {
  const N = () => OM.norm;

  function displayName(o, aliases) {
    const alias = aliases && aliases[OM.merge.itemKey(o)];
    if (alias) return alias;
    return o.option ? `${o.product} / ${o.option}` : o.product;
  }

  // 택배사 업로드용(기본 양식): 한 줄 = 한 박스.
  function courierRows(boxes, aliases) {
    const head = ['받는분', '받는분 전화번호', '우편번호', '주소', '품목', '수량', '배송메모', '마켓', '원주문번호'];
    const rows = boxes.map(list => {
      const o = list[0];
      return [
        o.receiver,
        N().phoneDisplay(o.phone),
        o.zip,
        o.address,
        list.map(x => `${displayName(x, aliases)} x${x.qty}`).join(' / '),
        list.reduce((s, x) => s + x.qty, 0),
        [...new Set(list.map(x => x.memo).filter(Boolean))].join(' / '),
        [...new Set(list.map(x => x.market))].join(', '),
        list.map(x => `${x.market}:${x.orderNo}`).join(' '),
      ];
    });
    return [head, ...rows];
  }

  // 피킹 리스트: 상품(줄임말 기준)별 수량 합계, 많은 순.
  function pickList(orders, excluded, aliases) {
    const m = new Map();
    for (const o of orders) {
      if (excluded && excluded.has(o.id)) continue;
      const k = displayName(o, aliases);
      m.set(k, (m.get(k) || 0) + o.qty);
    }
    return [...m.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'ko'));
  }

  function pickRows(orders, excluded, aliases) {
    return [['상품', '수량'], ...pickList(orders, excluded, aliases)];
  }

  // 통합 주문표: 주문 줄 하나씩, 원래 정보를 모두 남긴다.
  function allRows(orders, excluded, aliases, groupNo) {
    const head = ['마켓', '주문번호', '수취인', '연락처', '우편번호', '주소', '상품명', '옵션', '줄임말', '수량', '배송메모', '결제금액', '합포장 묶음'];
    const rows = orders.filter(o => !(excluded && excluded.has(o.id))).map(o => [
      o.market, o.orderNo, o.receiver, N().phoneDisplay(o.phone), o.zip, o.address,
      o.product, o.option, (aliases && aliases[OM.merge.itemKey(o)]) || '', o.qty, o.memo, o.amount,
      (groupNo && groupNo[o.id]) ? `묶음 ${groupNo[o.id]}` : '',
    ]);
    return [head, ...rows];
  }

  OM.exporters = { displayName, courierRows, pickList, pickRows, allRows };
})(globalThis.OM = globalThis.OM || {});
