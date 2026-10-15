// 여러 파일의 주문을 합치고, 중복을 빼고, 합포장 묶음을 만든다.
(function (OM) {
  const N = () => OM.norm;

  // 같은 주문 줄인지 판단하는 키. 쿠팡처럼 한 주문번호에 상품 줄이 여러 개인 경우를 위해 상품·옵션까지 포함한다.
  function lineKey(o) {
    return [o.market, o.orderNo, o.product, o.option].join('\u0001');
  }

  function shipKey(o) {
    return [N().name(o.receiver), N().phone(o.phone), N().address(o.address)].join('\u0001');
  }

  // 상품명 줄임말을 찾을 때 쓰는 키(상품명 + 옵션).
  function itemKey(o) {
    return o.product + '\u0001' + (o.option || '');
  }

  // fileResults: [{ parsed: { status, orders } }] → { orders, duplicates }
  function combine(fileResults) {
    const seen = new Set();
    const orders = [];
    let duplicates = 0;
    for (const fr of fileResults) {
      if (!fr.parsed || fr.parsed.status === 'unknown' || fr.parsed.status === 'empty') continue;
      for (const o of fr.parsed.orders) {
        const k = lineKey(o);
        if (seen.has(k)) { duplicates++; continue; }
        seen.add(k);
        orders.push({ ...o, id: k });
      }
    }
    return { orders, duplicates };
  }

  // excluded: 보내지 않을 주문 id, unmerged: 묶음에서 뺀 주문 id
  function group(orders, excluded, unmerged) {
    excluded = excluded || new Set();
    unmerged = unmerged || new Set();
    const boxes = new Map();
    for (const o of orders) {
      if (excluded.has(o.id)) continue;
      const k = unmerged.has(o.id) ? 'solo\u0001' + o.id : shipKey(o);
      if (!boxes.has(k)) boxes.set(k, []);
      boxes.get(k).push(o);
    }
    const groupNo = {};
    const memoConflict = new Set();
    let n = 0;
    for (const list of boxes.values()) {
      if (list.length < 2) continue;
      n++;
      list.forEach(o => { groupNo[o.id] = n; });
      const memos = new Set(list.map(o => o.memo).filter(Boolean));
      if (memos.size > 1) memoConflict.add(n);
    }

    // 주소는 같은데 이름·연락처가 달라 자동으로 묶지 않은 주문(가족 주문 등). 확인만 권한다.
    const byAddr = new Map();
    for (const list of boxes.values()) {
      const a = N().address(list[0].address);
      if (!byAddr.has(a)) byAddr.set(a, 0);
      byAddr.set(a, byAddr.get(a) + 1);
    }
    const sameAddress = new Set();
    for (const list of boxes.values()) {
      if (byAddr.get(N().address(list[0].address)) > 1) list.forEach(o => sameAddress.add(o.id));
    }

    return { boxes: [...boxes.values()], groupNo, groups: n, memoConflict, sameAddress };
  }

  OM.merge = { combine, group, lineKey, shipKey, itemKey };
})(globalThis.OM = globalThis.OM || {});
