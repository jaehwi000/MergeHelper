// 파일 바이트 → 시트별 2차원 배열. 화면과 테스트가 같은 코드로 파일을 읽도록 한 곳에 둔다.
(function (OM) {
  // 숫자 셀은 문자열로 바꾼다. 큰 주문번호가 지수 표기(2.02E+12)로 바뀌지 않도록 서식 대신 원래 값을 쓴다.
  function cellText(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') return Number.isInteger(v) ? v.toFixed(0) : String(v);
    return String(v);
  }

  // 반환: { sheets: [{ name, aoa }], encoding } / 실패 시 Error(code: 'encrypted' | 'unreadable')
  OM.readWorkbook = function (XLSX, fileName, bytes) {
    let wb, encoding = null;
    try {
      if (/\.(csv|txt)$/i.test(fileName)) {
        const d = OM.decodeText(bytes);
        encoding = d.encoding;
        wb = XLSX.read(d.text, { type: 'string', raw: true });
      } else {
        wb = XLSX.read(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes), { type: 'array' });
      }
    } catch (e) {
      const err = new Error(e && e.message);
      err.code = /password|encrypt/i.test(String(e && e.message)) ? 'encrypted' : 'unreadable';
      throw err;
    }
    const sheets = wb.SheetNames.map(name => ({
      name,
      aoa: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '', blankrows: false })
        .map(row => row.map(cellText)),
    }));
    return { sheets, encoding };
  };

  // 여러 시트 중 주문표로 읽히는 첫 시트를 고른다.
  OM.parseWorkbook = function (book, savedMappings) {
    let first = null;
    for (const s of book.sheets) {
      const parsed = OM.parse.parseSheet(s.aoa, savedMappings);
      if (parsed.status === 'ok' || parsed.status === 'guessed') return { sheet: s, parsed };
      if (!first && parsed.status !== 'empty') first = { sheet: s, parsed };
    }
    return first || { sheet: book.sheets[0] || { name: '', aoa: [] }, parsed: { status: 'empty', headers: [], orders: [] } };
  };
})(globalThis.OM = globalThis.OM || {});
