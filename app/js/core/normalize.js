// 합포장 판단과 열 이름 비교를 위한 정규화 함수.
(function (OM) {
  const REGIONS = [
    [/^서울(특별시|시)?/, '서울'], [/^부산(광역시|시)?/, '부산'], [/^대구(광역시|시)?/, '대구'],
    [/^인천(광역시|시)?/, '인천'], [/^광주(광역시|시)?/, '광주'], [/^대전(광역시|시)?/, '대전'],
    [/^울산(광역시|시)?/, '울산'], [/^세종(특별자치시|시)?/, '세종'], [/^경기(도)?/, '경기'],
    [/^강원(특별자치도|도)?/, '강원'], [/^(충청북도|충북)/, '충북'], [/^(충청남도|충남)/, '충남'],
    [/^(전라북도|전북특별자치도|전북)/, '전북'], [/^(전라남도|전남)/, '전남'],
    [/^(경상북도|경북)/, '경북'], [/^(경상남도|경남)/, '경남'], [/^제주(특별자치도|도)?/, '제주'],
  ];

  const str = v => (v === null || v === undefined ? '' : String(v)).trim();

  OM.norm = {
    str,

    // 열 이름 비교용: 공백·줄바꿈 제거, 영문 소문자.
    header(h) {
      return str(h).replace(/\s+/g, '').toLowerCase();
    },

    // 숫자만 남긴다. 엑셀이 숫자로 저장해 앞자리 0이 빠진 휴대폰 번호(1012345678)는 0을 되살린다.
    phone(p) {
      let d = str(p).replace(/\D/g, '');
      if (/^1\d{8,9}$/.test(d)) d = '0' + d;
      return d;
    },

    // 화면 표시용 휴대폰 번호(010-1234-5678).
    phoneDisplay(p) {
      const d = OM.norm.phone(p);
      if (/^01\d{8,9}$/.test(d)) return d.replace(/^(\d{3})(\d{3,4})(\d{4})$/, '$1-$2-$3');
      return str(p);
    },

    // 같은 주소를 같은 키로: 시·도 표기 통일, 괄호 속 참고항목 제거, 공백·쉼표·하이픈 제거.
    address(a) {
      let s = str(a).replace(/\([^)]*\)/g, ' ').replace(/\s+/g, ' ').trim();
      for (const [re, to] of REGIONS) {
        if (re.test(s)) { s = s.replace(re, to); break; }
      }
      return s.replace(/[\s,\-]/g, '');
    },

    name(n) {
      return str(n).replace(/\s+/g, '');
    },

    qty(q) {
      const n = Number(str(q).replace(/[^\d.\-]/g, ''));
      return Number.isFinite(n) && n > 0 ? Math.round(n) : null;
    },
  };
})(globalThis.OM = globalThis.OM || {});
