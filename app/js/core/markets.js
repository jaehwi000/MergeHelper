// 마켓별 주문 파일 인식 규칙.
// 주의: 헤더 이름은 공개 자료를 바탕으로 한 추정이며 실제 다운로드 파일로 검증이 필요하다.
// 인식에 실패해도 열 연결(매핑) 화면으로 진행할 수 있다.
//
// fields 값은 "대안 목록"이다. 대안은 열 이름 하나이거나, 여러 열을 공백으로 이어 붙일 열 이름 배열이다.
// 행마다 앞에서부터 값이 비어 있지 않은 첫 대안을 쓴다.
(function (OM) {
  OM.MARKETS = [
    {
      id: 'smartstore',
      name: '스마트스토어',
      signature: ['상품주문번호', '수취인명'],
      fields: {
        orderNo: ['상품주문번호'],
        receiver: ['수취인명'],
        phone: ['수취인연락처1', '수취인연락처2'],
        address: ['통합배송지', ['기본배송지', '상세배송지']],
        zip: ['우편번호'],
        product: ['상품명'],
        option: ['옵션정보'],
        qty: ['수량'],
        memo: ['배송메세지', '배송메시지'],
        amount: ['최종 상품별 총 주문금액'],
      },
      // 송장번호를 되돌려 넣을 열. 파일에 없으면 맨 뒤에 추가한다.
      shipping: { courier: '택배사', tracking: '송장번호' },
    },
    {
      id: 'coupang',
      name: '쿠팡',
      signature: ['묶음배송번호', '수취인이름'],
      fields: {
        orderNo: ['주문번호'],
        receiver: ['수취인이름'],
        phone: ['수취인전화번호'],
        address: ['수취인 주소'],
        zip: ['우편번호'],
        product: ['등록상품명', '노출상품명(옵션명)'],
        option: ['등록옵션명'],
        qty: ['구매수(수량)'],
        memo: ['배송메세지'],
        amount: ['결제액'],
      },
      shipping: { courier: '택배사', tracking: '운송장번호' },
    },
  ];
})(globalThis.OM = globalThis.OM || {});
