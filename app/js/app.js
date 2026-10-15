// 화면 동작. 데이터 처리는 js/core/*.js에 있고, 여기서는 상태를 들고 화면을 그린다.
(function (OM, XLSX) {
  const MAX_FILE_MB = 30;
  const $ = id => document.getElementById(id);
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  // 설정만 브라우저 저장소에 남긴다. 사생활 보호 모드 등에서 실패해도 동작은 계속된다.
  const store = {
    get(key, def) { try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : def; } catch (e) { return def; } },
    set(key, val) { try { localStorage.setItem(key, JSON.stringify(val)); } catch (e) { /* 저장 실패는 무시 */ } },
  };

  const state = {
    files: [],          // { id, name, size, error?, book?, sheet?, parsed? }
    filter: '전체',
    excluded: new Set(),
    unmerged: new Set(),
    mapIdx: -1,
    savedMappings: store.get('om.mappings', {}),
    aliases: store.get('om.aliases', {}),
  };
  let nextId = 1;

  function toast(msg) {
    const t = $('toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toast.timer);
    toast.timer = setTimeout(() => t.classList.remove('show'), 2600);
  }

  // ---------- 화면 전환 ----------
  let step = 1;
  function go(s) {
    step = s;
    ['s1', 's2', 's21', 's3', 's4'].forEach(id => $(id).classList.add('hidden'));
    $(s === 2.1 ? 's21' : 's' + s).classList.remove('hidden');
    const cur = Math.floor(s);
    document.querySelectorAll('#steps li').forEach(li => {
      const n = +li.dataset.s;
      li.className = n < cur ? 'done' : n === cur ? 'on' : '';
    });
    ({ 2: renderS2, 2.1: renderMap, 3: renderS3, 4: renderS4 })[s]?.();
    window.scrollTo(0, 0);
  }

  // ---------- 파일 읽기 ----------
  function readAsArrayBuffer(file) {
    if (file.arrayBuffer) return file.arrayBuffer();
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(r.result);
      r.onerror = () => reject(r.error);
      r.readAsArrayBuffer(file);
    });
  }

  // '기타몰_주문내역.xlsx' → '기타몰'
  function marketFromFileName(fileName) {
    return fileName.replace(/\.[^.]+$/, '').replace(/[_\-\s]*(주문내역|주문목록|주문|orders?)[_\-\s\d()]*$/i, '').trim() || '기타';
  }

  function addBook(name, size, book) {
    const { sheet, parsed } = OM.parseWorkbook(book, state.savedMappings);
    // 열 이름으로 추측한 파일은 마켓 이름을 파일 이름에서 가져온다.
    if (parsed.status === 'guessed' && !parsed.market) {
      parsed.market = marketFromFileName(name);
      parsed.orders.forEach(o => { o.market = parsed.market; });
    }
    state.files.push({ id: nextId++, name, size, book, sheet, parsed });
  }

  async function addFiles(fileList) {
    const list = [...fileList];
    if (!list.length) return;
    for (const file of list) {
      if (state.files.some(f => f.name === file.name && f.size === file.size)) {
        toast(`${file.name}은(는) 이미 올렸어요.`);
        continue;
      }
      if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
        state.files.push({ id: nextId++, name: file.name, size: file.size, error: 'type' });
        continue;
      }
      if (file.size > MAX_FILE_MB * 1024 * 1024) {
        state.files.push({ id: nextId++, name: file.name, size: file.size, error: 'size' });
        continue;
      }
      try {
        const buf = await readAsArrayBuffer(file);
        addBook(file.name, file.size, OM.readWorkbook(XLSX, file.name, buf));
      } catch (e) {
        state.files.push({ id: nextId++, name: file.name, size: file.size, error: e.code || 'unreadable' });
      }
    }
    if (step === 2) renderS2(); else go(2);
  }

  function loadSamples() {
    resetAll();
    for (const s of OM.SAMPLES) {
      let bytes;
      if (s.csv) {
        bytes = new TextEncoder().encode(s.csv);
      } else {
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(s.aoa), 'Sheet1');
        bytes = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
      }
      addBook(s.name, bytes.byteLength || bytes.length, OM.readWorkbook(XLSX, s.name, bytes));
    }
    go(2);
  }

  function resetAll() {
    state.files = [];
    state.excluded.clear();
    state.unmerged.clear();
    state.filter = '전체';
  }

  function combined() {
    return OM.merge.combine(state.files.filter(f => f.parsed));
  }

  // ---------- S2 파일 확인 ----------
  const ERRORS = {
    encrypted: ['암호 걸림', '암호가 걸린 파일은 브라우저에서 열 수 없어요. 엑셀에서 파일을 열고 암호를 지운 뒤(파일 → 정보 → 통합 문서 보호 → 암호 설정에서 비우기) 다시 저장해 올려 주세요.'],
    unreadable: ['읽을 수 없음', '엑셀 파일이 손상되었거나 지원하지 않는 형식이에요. 엑셀에서 열어 xlsx로 다시 저장해 보세요.'],
    type: ['지원 안 함', '엑셀(xlsx, xls) 또는 csv 파일만 올릴 수 있어요.'],
    size: ['너무 큼', `${MAX_FILE_MB}MB보다 큰 파일은 올릴 수 없어요. 기간을 나눠서 내려받아 주세요.`],
  };

  function fileStatus(f) {
    if (f.error) return { cls: 'err', label: ERRORS[f.error][0], help: ERRORS[f.error][1] };
    const p = f.parsed;
    if (p.status === 'empty') return { cls: 'err', label: '주문표 없음', help: '주문 목록으로 보이는 표를 찾지 못했어요.' };
    if (p.status === 'unknown') return { cls: 'err', label: '열 연결 필요', help: `찾지 못한 항목: ${p.missing.join(', ')}`, action: '열 직접 연결하기' };
    if (p.status === 'guessed') return { cls: 'warn', label: '추측함 · 확인 필요', help: '열 이름으로 추측해서 읽었어요. 맞는지 한 번 확인해 주세요.', action: '열 확인하기' };
    return { cls: 'ok', label: '✓ ' + p.market, action: p.mapping ? '열 다시 연결' : null };
  }

  function renderS2() {
    const { orders, duplicates } = combined();
    $('s2-files').textContent = state.files.length;
    $('s2-orders').textContent = orders.length;
    $('file-list').innerHTML = state.files.map((f, i) => {
      const st = fileStatus(f);
      const count = f.parsed && f.parsed.orders ? `${f.parsed.orders.length}건` : '';
      const qtyNote = f.parsed && f.parsed.badQty ? `<p class="help">수량을 읽지 못한 ${f.parsed.badQty}줄은 1개로 처리했어요.</p>` : '';
      return `<li>
        <div class="head">
          <span class="badge ${st.cls}">${esc(st.label)}</span>
          <span class="fname">${esc(f.name)}</span>
          <span class="muted small">${count}</span>
          ${st.action ? `<button class="link" data-map="${i}">${esc(st.action)}</button>` : ''}
          <button class="link" data-del="${i}" aria-label="${esc(f.name)} 빼기">✕</button>
        </div>
        ${st.help ? `<p class="help">${esc(st.help)}</p>` : ''}${qtyNote}
      </li>`;
    }).join('');

    const notices = [];
    if (duplicates) notices.push(`<div class="notice warn">⚠ <div>같은 주문 <b>${duplicates}건</b>이 두 번 들어 있어요. 한 건만 남기고 합칩니다.</div></div>`);
    const blocked = state.files.filter(f => f.error || (f.parsed && (f.parsed.status === 'unknown' || f.parsed.status === 'empty'))).length;
    if (blocked) notices.push(`<div class="notice err">! <div>읽지 못한 파일 ${blocked}개는 빼고 진행합니다.</div></div>`);
    $('s2-notices').innerHTML = notices.join('');
    $('s2-next').disabled = orders.length === 0;
  }

  // ---------- S2-1 열 연결 ----------
  function renderMap() {
    const f = state.files[state.mapIdx];
    const p = f.parsed;
    const headers = [...new Set(p.headers.filter(Boolean))];
    const current = p.mapping || OM.parse.guessMapping(p.headers);
    $('map-title').textContent = `${f.name}의 열을 알려 주세요`;
    $('map-grid').innerHTML = OM.FIELDS.map(fd => `
      <label class="${fd.req ? 'req' : ''}" for="m-${fd.key}">${esc(fd.label)}</label>
      <select id="m-${fd.key}"><option value="">선택 안 함</option>${headers.map(h => `<option value="${esc(h)}" ${current[fd.key] === h ? 'selected' : ''}>${esc(h)}</option>`).join('')}</select>`).join('');
    $('map-preview').innerHTML = `<tr>${p.headers.map(h => `<th>${esc(h)}</th>`).join('')}</tr>` +
      (p.preview || []).map(r => `<tr>${p.headers.map((_, i) => `<td>${esc(r[i])}</td>`).join('')}</tr>`).join('');
    $('map-market').value = p.market || marketFromFileName(f.name);
  }

  function confirmMap() {
    const f = state.files[state.mapIdx];
    const mapping = {};
    for (const fd of OM.FIELDS) {
      const v = $('m-' + fd.key).value;
      if (v) mapping[fd.key] = v;
      else if (fd.req) { toast(`'${fd.label}' 열을 골라 주세요.`); $('m-' + fd.key).focus(); return; }
    }
    const market = $('map-market').value.trim() || '기타';
    f.parsed = OM.parse.applyMapping(f.sheet.aoa, f.parsed, market, mapping);
    if ($('map-remember').checked) {
      state.savedMappings[f.parsed.signature] = { market, mapping };
      store.set('om.mappings', state.savedMappings);
      // 같은 양식의 다른 파일에도 바로 적용한다.
      state.files.forEach(o => {
        if (o !== f && o.parsed && o.parsed.signature === f.parsed.signature && o.parsed.status !== 'ok') {
          o.parsed = OM.parse.applyMapping(o.sheet.aoa, o.parsed, market, mapping);
        }
      });
    }
    toast(`${f.parsed.orders.length}건을 읽었어요.`);
    go(2);
  }

  // ---------- S3 주문 정리 ----------
  const colorCls = n => `c${((n - 1) % 6) + 1}`;
  const name = o => OM.exporters.displayName(o, state.aliases);
  const last4 = p => OM.norm.phone(p).slice(-4);

  function renderS3() {
    const { orders } = combined();
    const g = OM.merge.group(orders, state.excluded, state.unmerged);
    const live = orders.filter(o => !state.excluded.has(o.id));
    $('s3-orders').textContent = live.length;
    $('s3-boxes').textContent = g.boxes.length;
    $('s3-groups').textContent = g.groups;

    const markets = ['전체', ...new Set(orders.map(o => o.market))];
    if (!markets.includes(state.filter)) state.filter = '전체';
    $('tabs').innerHTML = markets.map(m => {
      const c = m === '전체' ? orders.length : orders.filter(o => o.market === m).length;
      return `<button class="${state.filter === m ? 'on' : ''}" data-tab="${esc(m)}">${esc(m)} ${c}</button>`;
    }).join('');

    // 묶음끼리 붙여서 보여 준다. 나머지는 원래 순서.
    const pos = new Map(orders.map((o, i) => [o.id, i]));
    const sorted = [...orders].sort((a, b) => (g.groupNo[a.id] || 1e9) - (g.groupNo[b.id] || 1e9) || pos.get(a.id) - pos.get(b.id));
    const shown = sorted.filter(o => state.filter === '전체' || o.market === state.filter);

    const groupCell = o => {
      const n = g.groupNo[o.id];
      if (n) {
        const memo = g.memoConflict.has(n) ? ' <span class="badge warn" title="배송메모가 서로 달라요. 송장에는 함께 적힙니다.">메모 다름</span>' : '';
        return `<span class="grp ${colorCls(n)}"><i></i>묶음 ${n}</span> <button class="link small" data-split="${esc(o.id)}">풀기</button>${memo}`;
      }
      if (state.unmerged.has(o.id)) return `<button class="link small" data-join="${esc(o.id)}">다시 묶기</button>`;
      if (g.sameAddress.has(o.id) && !state.excluded.has(o.id)) return '<span class="badge warn" title="주소는 같지만 이름이나 연락처가 달라 자동으로 묶지 않았어요.">주소 같음</span>';
      return '<span class="muted">—</span>';
    };

    $('order-table').innerHTML = `<thead><tr><th></th><th>보냄</th><th>마켓</th><th>수취인</th><th>연락처</th><th>주소</th><th>상품</th><th class="num">수량</th><th>합포장</th></tr></thead><tbody>` +
      shown.map(o => {
        const n = g.groupNo[o.id], off = state.excluded.has(o.id);
        return `<tr class="${off ? 'off' : ''}">
          <td class="bar ${n ? colorCls(n) : ''}"></td>
          <td><input type="checkbox" data-send="${esc(o.id)}" ${off ? '' : 'checked'} aria-label="${esc(o.receiver)} 주문 보내기"></td>
          <td>${esc(o.market)}</td>
          <td>${esc(o.receiver)}</td>
          <td>…${esc(last4(o.phone))}</td>
          <td class="addr" title="${esc(o.address)}">${esc(o.address)}</td>
          <td class="prod wrap-cell" title="${esc(o.product)}${o.option ? ' / ' + esc(o.option) : ''}">${esc(name(o))}${state.aliases[OM.merge.itemKey(o)] ? '' : ' <span class="badge warn">원래 이름</span>'}</td>
          <td class="num">${o.qty}${o.qtyWarning ? ' <span class="badge warn" title="수량을 읽지 못해 1개로 처리했어요.">?</span>' : ''}</td>
          <td>${groupCell(o)}</td>
        </tr>`;
      }).join('') + '</tbody>';

    $('order-cards').innerHTML = shown.map(o => {
      const n = g.groupNo[o.id], off = state.excluded.has(o.id);
      return `<div class="ocard ${n ? colorCls(n) : ''} ${off ? 'off' : ''}">
        <div class="row between"><b>${esc(o.receiver)}</b><span class="muted small">${esc(o.market)} · …${esc(last4(o.phone))}</span></div>
        <div>${esc(name(o))} × ${o.qty}</div>
        <div class="row between small">${groupCell(o)}
          <label class="row"><input type="checkbox" data-send="${esc(o.id)}" ${off ? '' : 'checked'}> 보냄</label></div>
      </div>`;
    }).join('');

    const items = uniqueItems(orders);
    const missing = items.filter(it => !state.aliases[it.key]).length;
    $('alias-notice').innerHTML = missing
      ? `✏️ <div>상품명 줄임말 <b>${missing}개</b>가 아직 없어요. 송장과 피킹 리스트에 긴 원래 이름이 찍힙니다. <button class="link" id="alias-open">설정하기</button></div>`
      : `✓ <div>모든 상품에 줄임말이 있어요. <button class="link" id="alias-open">수정하기</button></div>`;
  }

  function uniqueItems(orders) {
    const m = new Map();
    for (const o of orders) {
      const key = OM.merge.itemKey(o);
      if (!m.has(key)) m.set(key, { key, product: o.product, option: o.option, markets: new Set() });
      m.get(key).markets.add(o.market);
    }
    return [...m.values()];
  }

  function renderAlias() {
    const items = uniqueItems(combined().orders);
    $('alias-list').innerHTML = items.map((it, i) => `
      <div class="alias-item">
        <label for="alias-${i}" class="small">${esc(it.product)}${it.option ? ` <span class="muted">/ ${esc(it.option)}</span>` : ''} <span class="muted">· ${esc([...it.markets].join(', '))}</span></label>
        <input type="text" id="alias-${i}" data-alias="${esc(it.key)}" value="${esc(state.aliases[it.key] || '')}" placeholder="짧은 이름" maxlength="40">
      </div>`).join('');
  }

  // ---------- S4 내려받기 ----------
  function renderS4() {
    const { orders } = combined();
    const g = OM.merge.group(orders, state.excluded, state.unmerged);
    const pick = OM.exporters.pickList(orders, state.excluded, state.aliases);
    $('ex-boxes').textContent = g.boxes.length;
    $('ex-orders').textContent = orders.filter(o => !state.excluded.has(o.id)).length;
    $('ex-kinds').textContent = pick.length;
    $('ex-qty').textContent = pick.reduce((s, [, q]) => s + q, 0);
    $('pick').innerHTML = pick.map(([k, q]) => `<li><span>${esc(k)}</span><b>×${q}</b></li>`).join('');
  }

  function saveXlsx(fileName, sheetName, rows) {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = rows[0].map((_, c) => ({ wch: Math.min(60, Math.max(6, ...rows.map(r => String(r[c] ?? '').length * 1.6))) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, fileName);
    toast(`${fileName} 파일을 내려받았어요.`);
  }

  function exportFile(kind) {
    const d = new Date();
    const stamp = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    const { orders } = combined();
    const g = OM.merge.group(orders, state.excluded, state.unmerged);
    if (kind === 'courier') return saveXlsx(`택배업로드_${stamp}.xlsx`, '택배업로드', OM.exporters.courierRows(g.boxes, state.aliases));
    if (kind === 'pick') return saveXlsx(`피킹리스트_${stamp}.xlsx`, '피킹리스트', OM.exporters.pickRows(orders, state.excluded, state.aliases));
    if (kind === 'all') return saveXlsx(`통합주문표_${stamp}.xlsx`, '통합주문표', OM.exporters.allRows(orders, state.excluded, state.aliases, g.groupNo));
    if (kind === 'print') {
      document.body.classList.add('print-pick');
      window.print();
      document.body.classList.remove('print-pick');
    }
  }

  // ---------- 이벤트 ----------
  $('file-input').addEventListener('change', e => {
    addFiles(e.target.files);
    e.target.value = '';
  });

  // 파일을 화면 어디에 놓아도 받는다. 놓을 곳을 벗어나 브라우저가 파일을 열어 버리는 일을 막는다.
  const drop = $('drop');
  window.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('over'); });
  window.addEventListener('dragleave', e => { if (!e.relatedTarget) drop.classList.remove('over'); });
  window.addEventListener('drop', e => {
    e.preventDefault();
    drop.classList.remove('over');
    if (e.dataTransfer && e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
  });

  document.addEventListener('change', e => {
    const t = e.target;
    if (t.dataset.send !== undefined) {
      t.checked ? state.excluded.delete(t.dataset.send) : state.excluded.add(t.dataset.send);
      renderS3();
    } else if (t.dataset.alias !== undefined) {
      const v = t.value.trim();
      if (v) state.aliases[t.dataset.alias] = v; else delete state.aliases[t.dataset.alias];
      store.set('om.aliases', state.aliases);
      renderS3();
    }
  });

  document.addEventListener('click', e => {
    const t = e.target.closest('button');
    if (!t) return;
    const ds = t.dataset;
    switch (t.id) {
      case 'btn-sample': return loadSamples();
      case 's2-back': case 's4-reset': resetAll(); return go(1);
      case 's2-next': return go(3);
      case 's21-cancel': return go(2);
      case 's21-ok': return confirmMap();
      case 's3-back': return go(2);
      case 's3-next': return go(4);
      case 's4-back': return go(3);
      case 'alias-open':
        $('alias-card').classList.remove('hidden');
        renderAlias();
        return $('alias-card').scrollIntoView({ behavior: 'smooth', block: 'start' });
      case 'alias-close': return $('alias-card').classList.add('hidden');
    }
    if (ds.map !== undefined) { state.mapIdx = +ds.map; go(2.1); }
    else if (ds.del !== undefined) { state.files.splice(+ds.del, 1); state.files.length ? renderS2() : go(1); }
    else if (ds.tab !== undefined) { state.filter = ds.tab; renderS3(); }
    else if (ds.split !== undefined) { state.unmerged.add(ds.split); renderS3(); toast('묶음에서 뺐어요. 따로 발송합니다.'); }
    else if (ds.join !== undefined) { state.unmerged.delete(ds.join); renderS3(); }
    else if (ds.ex !== undefined) exportFile(ds.ex);
  });

  go(1);
})(globalThis.OM, globalThis.XLSX);
