// CSV 글자 인코딩 판별. 국내 쇼핑몰·엑셀에서 저장한 CSV는 EUC-KR(CP949)인 경우가 많다.
(function (OM) {
  OM.decodeText = function (bytes) {
    const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    try {
      return { text: new TextDecoder('utf-8', { fatal: true }).decode(u8).replace(/^﻿/, ''), encoding: 'utf-8' };
    } catch (e) {
      return { text: new TextDecoder('euc-kr').decode(u8), encoding: 'euc-kr' };
    }
  };
})(globalThis.OM = globalThis.OM || {});
