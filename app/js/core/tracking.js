// 2단계: 택배사 결과 파일의 송장번호를 박스에 연결하고, 마켓별 원본 주문 파일에 되돌려 넣는다.
(function (OM) {
  const N = () => OM.norm;

  const COLS = {
    tracking: ['운송장번호', '송장번호', '운송장', '송장', '등기번호', 'trackingno', 'trackingnumber', 'invoiceno'],
    ref: ['원주문번호', '고객주문번호', '참조번호', '고객참조번호', '주문번호'],
    receiver: ['받는분성명', '받는분', '받는사람', '수하인명', '수하인', '수취인명', '수취인', '수령인명', '수령인'],
    phone: ['받는분전화번호', '받는분휴대폰', '받는분연락처', '수하인휴대폰', '수하인전화번호', '수취인전화번호', '수취인연락처', '수령인휴대폰', '휴대폰', '전화번호', '연락처'],
    address: ['받는분주소', '수하인주소', '수취인주소', '수령인주소', '배송지주소', '주소'],
  };
  // 마켓 규칙이 없는 파일에 송장을 넣을 때 쓰는 열 이름.
  const FILE_COLS = {
    tracking: ['송장번호', '운송장번호', '운송장', '송장'],
    courier: ['택배사', '택배사명', '배송업체', '배송사'],
  };
  const SENDER_WORDS = ['보내는', '송하인', '발송인', '보낸'];

  const boxKey = list => list.map(o => o.id).sort().join('\u0002');
  const normTracking = v => N().str(v).replace(/[\s\-]/g, '');

  // 마스킹된 번호(010-****-5678)는 뒤 4자리로만 비교한다.
  function phoneKeys(p) {
    const raw = N().str(p);
    const d = N().phone(raw);
    if (!d) return [];
    if (raw.includes('*')) return ['L4:' + d.slice(-4)];
    return ['F:' + d, 'L4:' + d.slice(-4)];
  }

  // 결과 파일에서 머리글 줄과 열을 찾는다. 송장번호 열과, 박스를 찾을 단서(원주문번호 또는 받는분)가 있어야 한다.
  function readResult(aoa) {
    for (let r = 0; r < Math.min(aoa.length, 15); r++) {
      const norm = (aoa[r] || []).map(h => N().header(h));
      const notSender = i => !SENDER_WORDS.some(w => norm[i].includes(w));
      const cols = {};
      const used = new Set();
      for (const key of ['tracking', 'ref', 'receiver', 'phone', 'address']) {
        const i = OM.parse.findColumn(norm, COLS[key], j => !used.has(j) && notSender(j));
        if (i >= 0) { cols[key] = i; used.add(i); }
      }
      if (cols.tracking !== undefined && (cols.ref !== undefined || cols.receiver !== undefined)) {
        return { headerRow: r, cols, headers: (aoa[r] || []).map(h => N().str(h)) };
      }
    }
    return null;
  }

  // boxes: OM.merge.group(...).boxes
  // 반환: { found, assigned: Map(boxKey → 송장번호), rows, unmatched: [{ row, receiver, tracking, reason }], conflicts }
  function matchResult(boxes, aoa) {
    const info = readResult(aoa);
    const assigned = new Map();
    const unmatched = [];
    const conflicts = [];
    if (!info) return { found: false, assigned, rows: 0, unmatched, conflicts };

    const byRef = new Map();      // '마켓:주문번호' 또는 '주문번호' → Set(boxKey)
    const byPerson = new Map();   // 이름 + 전화 키 → Set(boxKey)
    const addrOf = new Map();     // boxKey → 정규화 주소
    const add = (m, k, v) => { if (!m.has(k)) m.set(k, new Set()); m.get(k).add(v); };
    for (const list of boxes) {
      const k = boxKey(list);
      addrOf.set(k, N().address(list[0].address));
      for (const o of list) {
        add(byRef, `${o.market}:${o.orderNo}`, k);
        add(byRef, String(o.orderNo), k);
      }
      for (const pk of phoneKeys(list[0].phone)) add(byPerson, N().name(list[0].receiver) + '|' + pk, k);
    }

    const { headerRow, cols } = info;
    let rows = 0;
    for (let r = headerRow + 1; r < aoa.length; r++) {
      const row = aoa[r] || [];
      const tracking = normTracking(row[cols.tracking]);
      if (!tracking) continue;
      rows++;
      const receiver = N().str(row[cols.receiver]);
      let cand = new Set();

      if (cols.ref !== undefined) {
        for (const token of N().str(row[cols.ref]).split(/[\s,\/]+/).filter(Boolean)) {
          (byRef.get(token) || []).forEach(k => cand.add(k));
        }
      }
      if (!cand.size && cols.receiver !== undefined && cols.phone !== undefined) {
        for (const pk of phoneKeys(row[cols.phone])) {
          const hit = byPerson.get(N().name(receiver) + '|' + pk);
          if (hit) { cand = new Set(hit); break; }
        }
      }
      // 같은 사람의 박스가 여러 개면 주소로 좁힌다.
      if (cand.size > 1 && cols.address !== undefined) {
        const a = N().address(row[cols.address]);
        const narrowed = [...cand].filter(k => addrOf.get(k) === a);
        if (narrowed.length) cand = new Set(narrowed);
      }

      if (cand.size === 1) {
        const k = [...cand][0];
        const prev = assigned.get(k);
        if (prev && prev !== tracking) conflicts.push({ row: r + 1, receiver, tracking, existing: prev });
        else assigned.set(k, tracking);
      } else {
        unmatched.push({ row: r + 1, receiver, tracking, reason: cand.size ? '박스가 여러 개라 하나로 정하지 못함' : '맞는 주문을 찾지 못함' });
      }
    }
    return { found: true, assigned, rows, unmatched, conflicts, columns: info.headers };
  }

  // 박스별 송장번호 → 주문 줄(lineKey)별 송장번호
  function trackingByLine(boxes, assigned) {
    const m = new Map();
    for (const list of boxes) {
      const t = assigned.get(boxKey(list));
      if (t) list.forEach(o => m.set(o.id, t));
    }
    return m;
  }

  // 원본 주문 파일(aoa)의 각 주문 줄에 택배사·송장번호를 채운 새 표를 만든다.
  // parsed: 파일 해석 결과(orders에 row 번호가 있다). 반환: { aoa, filled, total, columns }
  function fillFile(aoa, parsed, byLine, courierName) {
    const out = aoa.map(r => [...r]);
    const header = out[parsed.headerRow];
    const norm = header.map(h => N().header(h));
    const market = OM.MARKETS.find(m => m.id === parsed.marketId);

    const locate = (named, syns) => {
      if (named) {
        const i = norm.indexOf(N().header(named));
        if (i >= 0) return i;
        header.push(named);
        norm.push(N().header(named));
        return header.length - 1;
      }
      const i = OM.parse.findColumn(norm, syns);
      if (i >= 0) return i;
      header.push(syns[0]);
      norm.push(N().header(syns[0]));
      return header.length - 1;
    };
    const courierCol = locate(market && market.shipping && market.shipping.courier, FILE_COLS.courier);
    const trackingCol = locate(market && market.shipping && market.shipping.tracking, FILE_COLS.tracking);

    let filled = 0;
    const width = header.length;
    for (const o of parsed.orders) {
      const t = byLine.get(OM.merge.lineKey(o));
      const row = out[o.row - 1];
      while (row.length < width) row.push('');
      if (!t) continue;
      row[courierCol] = courierName;
      row[trackingCol] = t;
      filled++;
    }
    return { aoa: out, filled, total: parsed.orders.length, columns: { courier: header[courierCol], tracking: header[trackingCol] } };
  }

  OM.tracking = { readResult, matchResult, trackingByLine, fillFile, boxKey, normTracking };
})(globalThis.OM = globalThis.OM || {});
