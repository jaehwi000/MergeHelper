// 시트(2차원 배열)를 표준 주문 목록으로 바꾼다.
(function (OM) {
  const N = () => OM.norm;
  const HEADER_SCAN_ROWS = 15;

  // 열 이름 목록의 지문. 같은 양식의 파일이면 같은 값이 나와 저장된 매핑을 다시 쓸 수 있다.
  function signature(headers) {
    return headers.map(h => N().header(h)).filter(Boolean).join('|');
  }

  function knownHeaderSet() {
    const set = new Set();
    Object.values(OM.SYNONYMS).flat().forEach(h => set.add(N().header(h)));
    OM.MARKETS.forEach(m => Object.values(m.fields).flat(2).forEach(h => set.add(N().header(h))));
    return set;
  }

  // 위쪽 몇 줄 중 알려진 열 이름이 가장 많이 들어 있는 줄을 머리글로 본다(제목 줄이 있는 파일 대비).
  function findHeaderRow(aoa, extraKnown) {
    const known = knownHeaderSet();
    (extraKnown || []).forEach(h => known.add(N().header(h)));
    let best = -1, bestScore = 0;
    for (let r = 0; r < Math.min(aoa.length, HEADER_SCAN_ROWS); r++) {
      const score = (aoa[r] || []).filter(c => known.has(N().header(c))).length;
      if (score > bestScore) { best = r; bestScore = score; }
    }
    return bestScore >= 3 ? best : -1;
  }

  function detectMarket(headers) {
    const set = new Set(headers.map(h => N().header(h)));
    return OM.MARKETS.find(m => m.signature.every(h => set.has(N().header(h)))) || null;
  }

  // 마켓 규칙의 대안 목록을 실제 열 번호로 바꾼다. 없는 열을 쓰는 대안은 버린다.
  function compileSpec(fields, headers) {
    const idx = new Map();
    headers.forEach((h, i) => { const k = N().header(h); if (k && !idx.has(k)) idx.set(k, i); });
    const spec = {};
    for (const [key, alts] of Object.entries(fields)) {
      const compiled = [];
      for (const alt of alts) {
        const names = Array.isArray(alt) ? alt : [alt];
        const cols = names.map(n => idx.get(N().header(n)));
        if (cols.every(c => c !== undefined)) compiled.push(cols);
      }
      if (compiled.length) spec[key] = compiled;
    }
    return spec;
  }

  // 알려진 마켓이 아닐 때 열 이름으로 항목을 추측한다. 반환값: { field: 열 이름 }
  function guessMapping(headers) {
    const norm = headers.map(h => N().header(h));
    const used = new Set();
    const result = {};
    const buyerish = h => OM.BUYER_WORDS.some(w => h.includes(w));
    for (const f of OM.FIELDS) {
      const syns = OM.SYNONYMS[f.key].map(s => N().header(s));
      const ok = i => !used.has(i) && norm[i] && !(['receiver', 'phone', 'address'].includes(f.key) && buyerish(norm[i]));
      let found = -1;
      for (const s of syns) {
        found = norm.findIndex((h, i) => ok(i) && h === s);
        if (found >= 0) break;
      }
      if (found < 0) {
        for (const s of syns) {
          if (s.length < 2) continue;
          found = norm.findIndex((h, i) => ok(i) && h.includes(s));
          if (found >= 0) break;
        }
      }
      if (found >= 0) { used.add(found); result[f.key] = headers[found]; }
    }
    return result;
  }

  function mappingToFields(mapping) {
    const fields = {};
    for (const [k, h] of Object.entries(mapping)) if (h) fields[k] = [h];
    return fields;
  }

  function missingRequired(spec) {
    return OM.FIELDS.filter(f => f.req && !spec[f.key]).map(f => f.label);
  }

  function readValue(row, alts) {
    if (!alts) return '';
    for (const cols of alts) {
      const v = cols.map(c => N().str(row[c])).filter(Boolean).join(' ');
      if (v) return v;
    }
    return '';
  }

  function rowsToOrders(aoa, headerRow, spec, marketName) {
    const orders = [];
    let badQty = 0;
    for (let r = headerRow + 1; r < aoa.length; r++) {
      const row = aoa[r] || [];
      const o = { market: marketName };
      for (const f of OM.FIELDS) o[f.key] = readValue(row, spec[f.key]);
      // 주문번호·수취인·상품명이 모두 비어 있으면 빈 줄이나 합계 줄로 보고 건너뛴다.
      if (!o.orderNo && !o.receiver && !o.product) continue;
      const q = N().qty(o.qty);
      if (q === null) { badQty++; o.qty = 1; o.qtyWarning = true; } else o.qty = q;
      o.row = r + 1;
      orders.push(o);
    }
    return { orders, badQty };
  }

  // 시트 하나를 해석한다.
  // savedMappings: { [signature]: { market: '이름', mapping: { field: 열 이름 } } }
  // 반환 status: ok(바로 사용) | guessed(추측함, 확인 권장) | unknown(필수 항목을 못 찾음) | empty(머리글 없음)
  function parseSheet(aoa, savedMappings) {
    let headerRow = findHeaderRow(aoa, Object.values(savedMappings || {}).flatMap(s => Object.values(s.mapping)));
    // 아는 열 이름이 하나도 없으면 값이 두 칸 이상 있는 첫 줄을 머리글로 보고 열 연결로 넘긴다.
    if (headerRow < 0) headerRow = aoa.findIndex(r => (r || []).filter(c => N().str(c)).length >= 2);
    if (headerRow < 0) return { status: 'empty', headers: [], orders: [] };
    const headers = (aoa[headerRow] || []).map(h => N().str(h));
    const sig = signature(headers);
    const base = { headers, headerRow, signature: sig, preview: aoa.slice(headerRow + 1, headerRow + 4) };

    const market = detectMarket(headers);
    const saved = savedMappings && savedMappings[sig];
    let spec, marketName, status, mapping = null;

    if (saved) {
      spec = compileSpec(mappingToFields(saved.mapping), headers);
      marketName = saved.market;
      mapping = saved.mapping;
      status = 'ok';
    } else if (market) {
      spec = compileSpec(market.fields, headers);
      marketName = market.name;
      status = 'ok';
    } else {
      mapping = guessMapping(headers);
      spec = compileSpec(mappingToFields(mapping), headers);
      marketName = '';
      status = 'guessed';
    }

    const missing = missingRequired(spec);
    if (missing.length) {
      if (!mapping) mapping = guessMapping(headers);
      return { ...base, status: 'unknown', market: marketName, mapping, missing, orders: [] };
    }
    const { orders, badQty } = rowsToOrders(aoa, headerRow, spec, marketName || '기타');
    return { ...base, status, market: marketName, mapping, missing: [], orders, badQty };
  }

  // 사용자가 열 연결을 마친 뒤 다시 해석한다.
  function applyMapping(aoa, parsed, marketName, mapping) {
    const spec = compileSpec(mappingToFields(mapping), parsed.headers);
    const missing = missingRequired(spec);
    if (missing.length) return { ...parsed, status: 'unknown', missing, mapping, orders: [] };
    const { orders, badQty } = rowsToOrders(aoa, parsed.headerRow, spec, marketName);
    return { ...parsed, status: 'ok', market: marketName, mapping, missing: [], orders, badQty };
  }

  OM.parse = { parseSheet, applyMapping, guessMapping, detectMarket, findHeaderRow, signature };
})(globalThis.OM = globalThis.OM || {});
