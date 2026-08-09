/* ============================================================================
 * HAIMS PLUS 회의록 PPT 생성 스크립트  (모비스 회의록 양식)
 * ----------------------------------------------------------------------------
 * 사용법
 *   1) node 환경에 pptxgenjs 설치되어 있어야 함  (npm i pptxgenjs)
 *   2) 아래 [1] 회의 정보 / [2] 회의 내용 / [3] Action Item 영역만 수정
 *   3) node meeting_minutes_template.js
 *
 * 레이아웃 규칙 (수정 불필요)
 *   - A4 세로 (8.2699 x 11.6889 inch)
 *   - 바깥 페이지 테두리 없음. 문서 전체가 하나의 연속된 표 구조
 *   - 회의내용 본문은 표가 아닌 "라벨 : 설명" 텍스트 배치
 *   - 내용이 넘치면 자동으로 다음 장 생성 (상단에 '회 의 내 용' 헤더 반복)
 *   - Action Item 은 No / 내용 / 담당 / 상태 4열 표
 *
 * 작성 함수
 *   section("1. 주요 설명 사항")        → 대분류
 *   sub("1) 소제목")                    → 중분류
 *   item("라벨:", "설명 문장")          → 라벨-설명 한 줄
 * ==========================================================================*/

const pptxgen = require("pptxgenjs");

/* ==========================================================================
 * [1] 회의 정보  ─ 매번 수정
 * ========================================================================*/
const MEETING = {
  주제: "국내유통망 차세대 시스템 – 물류부문 (입고 : 청구·직입고·미수령, PDA, 직송)",
  주관팀: "모비스서비스 PI 추진팀",
  작성자: "이정재 위원",
  장소: "위워크선릉3 14B",
  회의일: "2026.08.07 (금)",
  시간: "09:28 ~ 10:34    10:50 ~ 11:33    13:21 ~ 14:31    14:49 ~ 14:57",
  // 회의구분: 정보전달 / 이해조정 / 의견교환 / 문제해결 / 기타 중 택1
  회의구분: "정보전달",
  참석자: [
    "[모비스서비스PI 추진팀] 박지호 책임",
    "[모비스부품운영팀] 김창준 책임",
    "[수행사 기존/물류/품질] 이정재, 오국환, 김진래, 박주일, 조동명",
    "[수행사 데이터아키텍쳐] 이선실",
  ],
  // 2쪽 이후 우측 상단 러닝 헤더
  러닝헤더: "국내유통망 차세대 시스템 – 물류부문 / 2026.08.07 (금)",
  출력파일: "HAIMS_PLUS_회의록_260807.pptx",
};

/* ==========================================================================
 * [3] Action Item  ─ [No, 내용, 담당, 상태]
 *     상태 예시: 결정 / 검토 / 재검토 / 미결 / 확인 / 진행 / 협의
 * ========================================================================*/
const ACTIONS = [
  ["1", "미수령 화면 용어 변경 (저장 → 조정 / 삭제 → 통제 취소)", "수행사", "결정"],
  ["2", "분류 후(저장대기) 상태의 미수령 등록 기능 도입 여부 결정", "수행사 / 모비스", "미결"],
  ["3", "미수령 팩스 발송 가능 시점(저장 전까지) 내부 협의", "모비스", "미결"],
  ["4", "분류 전 정렬입고저장 폐지 여부 및 대안 검토", "수행사 / 모비스", "재검토"],
  ["5", "직불 처리 옵션 3종 개발 및 오픈 시 기본값 고정", "수행사", "결정"],
  ["6", "직입고 리스트 일괄 등록 폐지에 따른 현장 영향 검토", "수행사", "검토"],
  ["7", "직불 판매 결과 메시지(거래처·건수) 표시 기준 확정", "수행사", "검토"],
  ["8", "직송입고 납입처리와 입고처리 통합 가능 여부 확인", "수행사", "확인"],
  ["9", "파렛트·케이스 단위 일괄 도착 보고 지원 여부 확인", "수행사", "확인"],
  ["10", "입고 테이블 분리 및 헤더·디테일 구조 설계", "수행사 / DA", "진행"],
  ["11", "주문번호 테이블 신설 및 청구·조치 연계 구조 정의", "수행사 / DA", "진행"],
  ["12", "파일럿 검증 범위(주문번호 적용 화면) 재협의", "모비스 / 수행사", "협의"],
];

/* ==========================================================================
 *  ↓↓↓  아래 렌더링 엔진은 수정 불필요  ↓↓↓
 * ========================================================================*/
const FONT = "맑은 고딕";
const BLACK = "000000";
const LINE = "555555";
const GRAY_FILL = "EFEFEF";
const GRAY_TXT = "666666";
const WHITE = "FFFFFF";

const PAGE_W = 8.2699, PAGE_H = 11.6889;
const TX = 0.42, TW = 7.43, TR = TX + TW;
const BOX_BOTTOM = 11.02;
const PAGE_NUM_Y = 11.13;
const PAD_TOP = 0.16;

const pres = new pptxgen();
pres.defineLayout({ name: "A4P", width: PAGE_W, height: PAGE_H });
pres.layout = "A4P";

let pageNum = 0;

function cell(slide, x, y, w, h, opts) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: opts && opts.fill ? opts.fill : WHITE },
    line: { color: LINE, width: 0.75 },
  });
}
function cellText(slide, x, y, w, h, text, o) {
  slide.addText(text, {
    x, y, w, h,
    align: (o && o.align) || "center",
    valign: (o && o.valign) || "middle",
    fontFace: FONT,
    fontSize: (o && o.size) || 10,
    bold: !!(o && o.bold),
    color: (o && o.color) || BLACK,
  });
}
function addPageNum(slide) {
  pageNum += 1;
  slide.addText(String(pageNum), {
    x: TX, y: PAGE_NUM_Y, w: TW, h: 0.22, align: "center",
    fontFace: FONT, fontSize: 8.5, color: GRAY_TXT,
  });
}

/* ---------- 1쪽 : 제목 + 헤더 표 ---------- */
const s1 = pres.addSlide();
s1.background = { color: WHITE };
s1.addText("회 의 록", {
  x: TX, y: 0.32, w: TW, h: 0.5,
  fontFace: FONT, fontSize: 22, bold: true, color: BLACK, charSpacing: 3,
});
addPageNum(s1);

const R1 = 0.92, RH = 0.42;
const C_LBL = 1.12;
const C_DIV = TR - 1.6, C_DIV_W = 1.6;
const C_MID_L = 3.52, C_MID_LW = 1.02;

// 1행 : 주제
cell(s1, TX, R1, C_LBL, RH, { fill: GRAY_FILL });
cellText(s1, TX, R1, C_LBL, RH, "주 제", { bold: true, size: 10.5 });
cell(s1, TX + C_LBL, R1, C_DIV - TX - C_LBL, RH);
cellText(s1, TX + C_LBL + 0.1, R1, C_DIV - TX - C_LBL - 0.2, RH, MEETING.주제,
  { align: "left", size: 8.5 });

// 회의구분 (1~3행 병합)
cell(s1, C_DIV, R1, C_DIV_W, RH * 3);
cellText(s1, C_DIV, R1 + 0.02, C_DIV_W, 0.3, "회 의 구 분", { bold: true, size: 10.5 });
const KINDS = ["정보전달", "이해조정", "의견교환", "문제해결", "기 타"];
let ky = R1 + 0.36;
KINDS.forEach((k) => {
  s1.addText(k, { x: C_DIV + 0.1, y: ky, w: 1.0, h: 0.175, align: "left", valign: "middle", fontFace: FONT, fontSize: 8, color: BLACK });
  if (k.replace(/\s/g, "") === MEETING.회의구분.replace(/\s/g, "")) {
    s1.addText("V", { x: C_DIV + 1.05, y: ky, w: 0.45, h: 0.175, align: "center", valign: "middle", fontFace: FONT, fontSize: 9, bold: true, color: BLACK });
  }
  ky += 0.165;
});

// 2행 : 주관팀 / 작성자
const R2 = R1 + RH;
cell(s1, TX, R2, C_LBL, RH, { fill: GRAY_FILL });
cellText(s1, TX, R2, C_LBL, RH, "주 관 팀", { bold: true, size: 10.5 });
cell(s1, TX + C_LBL, R2, C_MID_L - TX - C_LBL, RH);
cellText(s1, TX + C_LBL, R2, C_MID_L - TX - C_LBL, RH, MEETING.주관팀, { size: 9.5 });
cell(s1, C_MID_L, R2, C_MID_LW, RH, { fill: GRAY_FILL });
cellText(s1, C_MID_L, R2, C_MID_LW, RH, "작 성 자", { bold: true, size: 10 });
cell(s1, C_MID_L + C_MID_LW, R2, C_DIV - C_MID_L - C_MID_LW, RH);
cellText(s1, C_MID_L + C_MID_LW, R2, C_DIV - C_MID_L - C_MID_LW, RH, MEETING.작성자, { size: 9.5 });

// 3행 : 장소 / 회의일
const R3 = R2 + RH;
cell(s1, TX, R3, C_LBL, RH, { fill: GRAY_FILL });
cellText(s1, TX, R3, C_LBL, RH, "장 소", { bold: true, size: 10.5 });
cell(s1, TX + C_LBL, R3, C_MID_L - TX - C_LBL, RH);
cellText(s1, TX + C_LBL, R3, C_MID_L - TX - C_LBL, RH, MEETING.장소, { size: 9.5 });
cell(s1, C_MID_L, R3, C_MID_LW, RH, { fill: GRAY_FILL });
cellText(s1, C_MID_L, R3, C_MID_LW, RH, "회 의 일", { bold: true, size: 10 });
cell(s1, C_MID_L + C_MID_LW, R3, C_DIV - C_MID_L - C_MID_LW, RH);
cellText(s1, C_MID_L + C_MID_LW, R3, C_DIV - C_MID_L - C_MID_LW, RH, MEETING.회의일, { size: 9.5 });

// 4행 : 시간
const R4 = R3 + RH, R4H = 0.5;
cell(s1, TX, R4, C_LBL, R4H, { fill: GRAY_FILL });
cellText(s1, TX, R4, C_LBL, R4H, "시 간", { bold: true, size: 10.5 });
cell(s1, TX + C_LBL, R4, TR - TX - C_LBL, R4H);
cellText(s1, TX + C_LBL, R4, TR - TX - C_LBL, R4H, MEETING.시간, { size: 9.5 });

// 5행 : 참석자 헤더
const R5 = R4 + R4H, R5H = 0.32;
cell(s1, TX, R5, TW, R5H, { fill: GRAY_FILL });
cellText(s1, TX, R5, TW, R5H, "참 석 자", { bold: true, size: 11 });

// 6행 : 참석자 내용 (인원수에 맞춰 높이 자동)
const ATT_LH = 0.245;
const R6 = R5 + R5H;
const R6H = Math.max(0.6, MEETING.참석자.length * ATT_LH + 0.16);
cell(s1, TX, R6, TW, R6H);
let ay = R6 + 0.06;
MEETING.참석자.forEach((t) => {
  s1.addText(t, { x: TX + 0.12, y: ay, w: TW - 0.3, h: ATT_LH, valign: "middle", fontFace: FONT, fontSize: 9.5, color: BLACK });
  ay += ATT_LH;
});

// 7행 : 회의내용 헤더
const R7 = R6 + R6H, R7H = 0.32;
cell(s1, TX, R7, TW, R7H, { fill: GRAY_FILL });
cellText(s1, TX, R7, TW, R7H, "회 의 내 용", { bold: true, size: 11 });

/* ---------- 내용 흐름 엔진 ---------- */
let cur = { slide: s1, boxTop: R7 + R7H, y: R7 + R7H + PAD_TOP };

function closeBox(bottomY) {
  if (cur.boxTop === null) return;
  cur.slide.addShape("rect", {
    x: TX, y: cur.boxTop, w: TW, h: bottomY - cur.boxTop,
    fill: { type: "none" }, line: { color: LINE, width: 0.75 },
  });
}
function newContentPage() {
  closeBox(BOX_BOTTOM);
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addPageNum(s);
  cell(s, TX, 0.42, TW, 0.32, { fill: GRAY_FILL });
  cellText(s, TX, 0.42, TW, 0.32, "회 의 내 용", { bold: true, size: 11 });
  cur = { slide: s, boxTop: 0.74, y: 0.74 + PAD_TOP };
}
function ensure(h) {
  if (cur.y + h > BOX_BOTTOM - 0.12) newContentPage();
}
function estLines(text, cpl) {
  return Math.max(1, Math.ceil(text.length / cpl));
}

const X_SEC = TX + 0.16;
const X_SUB = TX + 0.38;
const X_LBL = TX + 0.6, W_LBL = 1.38;
const X_DSC = TX + 2.06, W_DSC = TR - (TX + 2.06) - 0.16;

function section(t) {
  ensure(0.42);
  cur.slide.addText(t, { x: X_SEC, y: cur.y, w: TW - 0.3, h: 0.26, fontFace: FONT, fontSize: 11.5, bold: true, color: BLACK });
  cur.y += 0.38;
}
function sub(t) {
  // 소제목만 페이지 끝에 홀로 남지 않도록 다음 항목 공간까지 확보
  ensure(0.34 + 0.62);
  cur.slide.addText(t, { x: X_SUB, y: cur.y, w: TW - 0.5, h: 0.24, fontFace: FONT, fontSize: 10.5, bold: true, color: BLACK });
  cur.y += 0.31;
}
function item(label, desc) {
  const h = Math.max(0.24, estLines(desc, 44) * 0.207 + 0.03);
  ensure(h + 0.08);
  cur.slide.addText(label, { x: X_LBL, y: cur.y, w: W_LBL, h, fontFace: FONT, fontSize: 10, bold: true, color: BLACK, valign: "top", align: "right" });
  cur.slide.addText(desc, { x: X_DSC, y: cur.y, w: W_DSC, h, fontFace: FONT, fontSize: 10, color: BLACK, valign: "top", lineSpacingMultiple: 1.16 });
  cur.y += h + 0.14;
}

/* ==========================================================================
 * [2] 회의 내용  ─ 매번 수정
 *     section() → sub() → item() 순으로 작성
 * ========================================================================*/
section("1. 주요 설명 사항");

sub("1) 입고 기본 개념 및 CWSF 구조");
item("기본 구조:", "대리점·지원센터가 청구하면 모비스가 조치하고, 그 조치 증표를 기준으로 입고를 잡음. 모비스 조치번호와 지원센터 판매번호가 CWSF 테이블에 집계되며, 청구입고 화면은 이 테이블을 조회하는 화면임.");
item("매칭 규칙:", "조치번호 1건에 입고번호 1건이 1:1로 매칭됨. 다만 하나의 청구번호에 조치번호가 복수로 생성될 수 있어(모비스 재고 부족으로 분할 조치) 청구 1건에 입고번호가 여러 개 생길 수 있음.");
item("연쇄 구조:", "모비스 조치 → 지원센터 CWSF 적재 → 지원센터 입고·직불 → 판매 처리 → 대리점 CWSF 적재 순으로 동일 프로세스가 반복됨. 테이블은 SWR(대리점 청구)과 지원센터 청구가 분리되어 있음.");
item("코드 체계:", "업체코드는 4자리(모비스=0001)이며 지원센터도 거래처 코드로 개별 등록해야 정산이 가능함. 계열코드는 향후 HK → HKC 로 표기 변경 예정.");

sub("2) 청구입고(웹) : 주문번호 매칭에 따른 직불 처리");
item("매칭 O:", "잔여 주문수량(자녀 주문수량)이 있는 건은 웹에서 일괄 자동 판매(직불) 처리됨. 처리 대상과 수량이 명확하므로 자동화 가능.");
item("매칭 X:", "주문번호가 매칭되지 않은 건은 저장 처리만 가능하며, 총주문수량은 참고 정보로만 표시됨. 어느 주문에 배분할지 시스템이 판단할 수 없어 웹에서는 판매 기능을 제공하지 않음.");
item("설계 사유:", "웹은 한 번에 30~40건을 일괄 처리하는 경우가 많아 건별 팝업 방식은 처리 지연과 오류를 유발함. 실제 주문번호가 물린 건은 전체의 약 20% 수준임.");
item("수량 개념:", "주문수량은 해당 주문번호의 잔여 수량, 총주문수량은 그 부품에 걸린 전체 잔여 주문수량의 합계이며 둘 다 실시간 조회값임.");

sub("3) 미수령 처리");
item("처리 방식:", "미수령 등록 시 해당 수량은 통제재고로 잡히고, 미수령 화면에서 저장 시 통제를 해제하는 조정이 이루어짐. 관할 사업소 통보를 위한 팩스 발송 기능이 화면에 포함됨.");
item("범위 한계:", "입고번호가 생성되는 시점에 입고·직불·미수령 수량이 모두 확정되므로, 이후 미수령을 해제해도 직불을 추가로 떨어뜨리지 않고 가용재고 증가로만 처리함.");
item("용어 변경:", "미수령 화면의 '저장'은 실제로 조정 행위이므로 '조정'으로, '삭제'는 '통제 취소'로 명칭을 변경하기로 함.");

sub("4) 직입고 (청구 없는 입고)");
item("용도:", "타 메이커 부품이나 타 대리점 재고 구매 등 청구 절차 없이 재고를 확보한 경우에 사용하며, 임의 업체코드를 등록해 입고를 잡음. 구입 단가는 대리점이 직접 입력함.");
item("현행 문제:", "엑셀과 매크로로 입고번호 하나에 수백 건을 한 번에 쌓아 저장하는 사례가 다수 있음.");
item("입고판매:", "직입고는 주문번호가 물려 있지 않으므로 해당 부품의 총주문수량을 조회하여 판매 대상을 판단함. 1회 처리는 최대 5라인으로 제한하고 라인마다 팝업으로 판매처를 선택하도록 함.");

sub("5) PDA 입고분류 · 저장 체계 재편");
item("단계 구분:", "분류 단계에서 입고번호가 생성되며 '입고대기 → 저장대기'로 상태가 전환되고, 저장 단계에서 로케이션 가용재고로 확정됨. 분류 시점에도 직불 처리는 함께 수행되어야 함.");
item("화면 분리:", "기존에는 분류 여부와 무관하게 모든 상태값을 불러와 저장할 수 있었으나, 앞으로는 분류한 건은 '분류 후 정렬입고저장'에서만, 분류하지 않은 건은 '일반 입고저장'에서만 처리하도록 분리함.");
item("일괄 처리:", "분류가 끝난 건은 직불이 이미 종료되었으므로 일괄 선택 저장을 허용함. 팝업은 자동으로 띄우지 않고 대상 건수만 표시하되, 수량이 맞지 않을 때 대상 내역을 열어 확인할 수 있도록 함.");
item("필드 전환:", "대리점은 '주문수량' 기준, 지원센터는 '청구수량' 기준으로 화면 필드와 메시지가 전환되어야 함(프로세스와 CWSF 유입 구조는 동일).");

sub("6) PDA 직불 처리 방식");
item("원바이원:", "주문번호가 매칭된 건은 조치 증표 단위로 한 건씩 처리하며, 처리 결과(판매 수량·저장 수량)를 팝업 메시지로 표시함. 'ALL' 기능으로 다음 건이 자동 조회되도록 하여 반복 스캔 부담을 줄임.");
item("일괄:", "주문번호가 없는 건은 일괄 선택 후 저장할 수 있도록 함.");
item("부분 판매:", "조치 수량보다 주문 수량이 많은 경우는 선택할 수 없도록 막고, 부분 판매 개념은 두지 않으며 남는 수량은 저장 처리하는 방향으로 정리함.");
item("증표 관계:", "입고증표 1건에 판매증표는 여러 건이 생성될 수 있으므로, 입고 테이블에는 판매번호를 두지 않고 직불 수량만 관리함. 판매번호 확인은 판매 테이블에서 입고번호로 역조회함.");

sub("7) 직송입고 및 ASM 납입처리");
item("배경:", "협력·생산업체에서 지원센터로 물건이 직송된 경우 모비스가 실물을 본 적이 없어 조치 증표가 생성되지 않으며, 이 상태로는 입고를 잡을 수 없음.");
item("처리 방식:", "지원센터가 ASM 번호(업체 납품 번호)를 조회하여 도착 수량을 입력하고 대신 납입 처리를 수행하면, 조치 증표가 생성되고 CWSF까지 내려와 정상 입고가 가능해짐.");
item("검토 사항:", "납입 처리와 입고 처리를 두 단계로 나눌 필요가 있는지, 한 화면에서 입고까지 완결할 수 있는지 추가 확인이 필요함. 파렛트·케이스 단위 일괄 도착 보고 가능 여부도 함께 확인하기로 함.");

sub("8) 도착처리 및 피킹일자 조건");
item("현황:", "CWSF에 조치 증표가 내려와도 관할 사업소의 피킹일자가 찍혀야만 입고를 잡을 수 있음. 물류센터에서만 조치된 건은 조회는 되지만 입고가 불가함.");
item("처리 방식:", "물건이 이미 도착한 예외 상황을 위해 파렛트 단위로 조회하여 피킹일자를 강제로 기록하는 화면을 별도로 둠. 조치가 완료된 건이므로 수량 변경은 불가하고 상태값만 갱신함.");
item("인터페이스:", "UWSF → CWSF 데이터 연계는 실시간(수분 단위) 인터페이스로 동작하는 것으로 정리함.");

sub("9) 테이블 및 번호 체계");
item("테이블 분리:", "기존에는 단일 전표 테이블에서 앞자리 구분값으로 입고·판매·조정을 함께 관리하였으나, 필드명이 판매 기준으로 되어 있어 혼선이 큼. 이번에 입고를 별도 테이블로 분리함.");
item("헤더-디테일:", "직입고는 입고번호 하나에 수십~수백 품목이 걸릴 수 있으므로, 헤더 테이블에서 항목 수를 관리하고 디테일 테이블에서 부품별로 관리하는 구조로 설계함(판매 테이블과 동일 방식).");
item("주문 테이블:", "주문번호 테이블이 신설되며, 청구 테이블이 주문번호를 물고 조치 테이블이 청구번호를 물어 입고 시점에 잔여 주문수량을 판단하는 구조임.");

sub("10) 현행 무재고 판매 · 역조정 프로세스");
item("현행 동작:", "재고가 0인 상태에서 판매하기 위해 무재고 조정으로 재고를 늘려 판매하고, 이후 입고가 되면 가장 오래된 건부터 자동으로 역조정되어 재고가 다시 차감됨.");
item("시사점:", "현행에서도 사실상 직불에 해당하는 행위가 재고상으로 이루어지고 있으나 작업자는 이를 인지하지 못함. 긴급건은 아예 저장하지 않고 별도로 빼두는 방식으로 운영되고 있음.");

section("2. 협의 및 결정 사항");

sub("1) 분류 후 미수령 등록 (미결)");
item("현황:", "입고번호가 생성된 이후에는 미수령을 등록할 수 없어, 대리점들이 분류 기능을 기피하는 주요 원인이 되고 있음. 분류 시점에 수량을 정확히 셀 수 없다는 것이 현장 의견임.");
item("논의 내용:", "저장대기 상태의 수량을 미수령으로 전환할 수 있는 기능을 두자는 안이 제시됨. PDA에서는 어렵더라도 웹에서 저장 전 상태에 한해 미수령 등록을 허용하는 방안을 검토함.");
item("전제 조건:", "미수령 등록 후 모비스에 팩스를 보낼 수 있는 시점이 저장 전까지 유효한지 내부 협의가 선행되어야 함. 결론은 추가 검토로 유보함.");

sub("2) 분류 전 정렬입고저장 폐지 (재검토)");
item("현황:", "리스트를 만든 뒤 일괄 저장하는 정렬입고저장은 직불 대상이 섞이면 라인을 분할해야 하므로 기능 구현이 어려워 분류 후에만 사용 가능하도록 제한함.");
item("우려 사항:", "기존에 사용하던 기능을 없애는 것이므로 오픈 시 현장 불만이 클 것으로 예상됨. 변화 관리만으로는 현장까지 전달되기 어렵다는 지적이 있었음.");
item("검토 방향:", "직불 대상 여부 필드를 추가해 표시만 하고 대상 건은 별도 처리하도록 유지하는 방안을 함께 검토하기로 함.");

sub("3) 직불 처리 옵션 체계");
item("결정 사항:", "직불 처리는 ① 주문번호가 매칭된 건만 처리, ② 주문번호 유무와 무관하게 팝업으로 선택 처리, ③ 직불 미사용의 세 가지 옵션으로 개발하되, 오픈 시점에는 ①을 기본값으로 고정하고 나머지는 비활성화함.");
item("운영 방침:", "현장 불만이 커질 경우에 한해 옵션을 활성화하며, 직입고 화면에서는 상위 옵션 설정과 무관하게 직불 판매를 선택할 수 있도록 허용함.");

sub("4) 직입고 리스트 일괄 등록 폐지");
item("결정 사항:", "직입고는 부품을 스캔하고 단가·수량을 입력해 한 건씩 저장하는 방식으로 변경함. 리스트를 쌓아 일괄 저장하면 직불 처리가 꼬일 수 있어 기존 방식을 포기함.");
item("영향:", "대리점 간 재고 거래 등 대량 구매 시 불편이 예상되므로 사용 추이를 보며 보완하기로 함.");

sub("5) 판매 결과 메시지 표시");
item("논의 내용:", "직불 처리 후 어느 거래처에 판매되었는지 알 수 없으면 작업자가 증표를 재조회해 다시 분류하는 이중 작업이 발생함.");
item("결정 사항:", "일반 입고저장은 조치 증표를 한 건씩 매칭하므로 판매된 거래처를 메시지로 표시하고, 다수 주문이 엮이는 경우에는 주문 건수와 판매 수량 수준으로만 표시하기로 함.");

sub("6) PDA 화면 링크 방식");
item("논의 내용:", "그리드 라인에 링크를 거는 방식은 오터치가 잦고 버튼 크기 확대 요청이 많으므로, 하단 버튼 방식으로 통일하기로 함.");
item("주문내역 조회:", "부품 목적별 주문 내역을 조회하는 화면을 제공하되, 조치 증표 단위 매칭은 주문 테이블 구조상 어려우므로 수량 수준까지만 표시하기로 함.");

sub("7) 파일럿 범위 관련 이견");
item("논의 내용:", "공통 모듈이 구축되지 않은 상태에서 주문번호 기준 전 구간 흐름을 파일럿으로 검증하는 것이 가능한지에 대해 이견이 있었음. 통합 테스트 단계에서 확인할 항목이라는 의견과, 엔드유저 사용성 확인이 필요하다는 의견이 대립함.");
item("조치 사항:", "화면 중 사용 빈도가 높은 주요 화면에 주문번호를 태워 검증하는 범위로 조정이 필요하며, 별도 협의하기로 함.");

/* ==========================================================================
 *  Action Item 렌더링 (수정 불필요)
 * ========================================================================*/
const AI_HDR_H = 0.32, AI_ROW_HDR = 0.3;
const aiRowH = ACTIONS.map((a) => Math.max(0.3, estLines(a[1], 40) * 0.2 + 0.1));
const aiTotal = AI_HDR_H + AI_ROW_HDR + aiRowH.reduce((a, b) => a + b, 0);

if (cur.y + 0.2 + aiTotal > BOX_BOTTOM) {
  closeBox(BOX_BOTTOM);
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addPageNum(s);
  cur = { slide: s, boxTop: null, y: 0.42 };
} else {
  closeBox(cur.y + 0.06);
  cur.y = cur.y + 0.06;
  cur.boxTop = null;
}

const aiSlide = cur.slide;
let aY2 = cur.y;

cell(aiSlide, TX, aY2, TW, AI_HDR_H, { fill: GRAY_FILL });
aiSlide.addText("Action Item", {
  x: TX + 0.12, y: aY2, w: TW - 0.3, h: AI_HDR_H,
  align: "left", valign: "middle", fontFace: FONT, fontSize: 11, bold: true, color: BLACK,
});
aY2 += AI_HDR_H;

const CW_NO = 0.55, CW_OWNER = 1.5, CW_STAT = 0.8;
const CW_ITEM = TW - CW_NO - CW_OWNER - CW_STAT;
const XN = TX, XI = XN + CW_NO, XO = XI + CW_ITEM, XS = XO + CW_OWNER;

cell(aiSlide, XN, aY2, CW_NO, AI_ROW_HDR, { fill: GRAY_FILL });
cell(aiSlide, XI, aY2, CW_ITEM, AI_ROW_HDR, { fill: GRAY_FILL });
cell(aiSlide, XO, aY2, CW_OWNER, AI_ROW_HDR, { fill: GRAY_FILL });
cell(aiSlide, XS, aY2, CW_STAT, AI_ROW_HDR, { fill: GRAY_FILL });
cellText(aiSlide, XN, aY2, CW_NO, AI_ROW_HDR, "No", { bold: true, size: 9, color: GRAY_TXT });
cellText(aiSlide, XI + 0.12, aY2, CW_ITEM - 0.2, AI_ROW_HDR, "내용", { bold: true, size: 9, color: GRAY_TXT, align: "left" });
cellText(aiSlide, XO, aY2, CW_OWNER, AI_ROW_HDR, "담당", { bold: true, size: 9, color: GRAY_TXT });
cellText(aiSlide, XS, aY2, CW_STAT, AI_ROW_HDR, "상태", { bold: true, size: 9, color: GRAY_TXT });
aY2 += AI_ROW_HDR;

ACTIONS.forEach((a, i) => {
  const h = aiRowH[i];
  cell(aiSlide, XN, aY2, CW_NO, h);
  cell(aiSlide, XI, aY2, CW_ITEM, h);
  cell(aiSlide, XO, aY2, CW_OWNER, h);
  cell(aiSlide, XS, aY2, CW_STAT, h);
  cellText(aiSlide, XN, aY2, CW_NO, h, a[0], { size: 9.5 });
  aiSlide.addText(a[1], { x: XI + 0.12, y: aY2, w: CW_ITEM - 0.24, h, align: "left", valign: "middle", fontFace: FONT, fontSize: 9.5, color: BLACK });
  cellText(aiSlide, XO, aY2, CW_OWNER, h, a[2], { size: 9 });
  cellText(aiSlide, XS, aY2, CW_STAT, h, a[3], { size: 9.5, bold: true });
  aY2 += h;
});

pres.writeFile({ fileName: MEETING.출력파일 }).then(() => console.log("생성 완료 : " + MEETING.출력파일));
