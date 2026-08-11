#!/usr/bin/env node
/* ============================================================================
 * 회의록 PPT 생성 스크립트 (모비스 회의록 양식)
 *
 * 사용법
 *   1) 전사 파일 → 구조화(JSON) → PPTX
 *        ANTHROPIC_API_KEY=sk-... node 회의록_생성.js transcript.txt
 *
 *   2) 구조화만 (JSON 검토·수정용)
 *        ANTHROPIC_API_KEY=sk-... node 회의록_생성.js transcript.txt --json-only
 *
 *   3) 이미 만든 JSON으로 PPTX만 생성 (API 호출 없음)
 *        node 회의록_생성.js minutes.json --from-json
 *
 * 옵션
 *   --out <파일명>     출력 pptx 파일명 지정
 *   --model <모델>     기본 claude-sonnet-5
 * ==========================================================================*/

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

/* ==========================================================================
 * [0] CLI 파싱
 * ========================================================================*/
const argv = process.argv.slice(2);
if (argv.length === 0 || argv.includes("-h") || argv.includes("--help")) {
  console.log(fs.readFileSync(__filename, "utf8").split("* ====")[1]);
  process.exit(0);
}

const inputPath = argv.find((a) => !a.startsWith("--"));
const flag = (name, def) => {
  const i = argv.indexOf(`--${name}`);
  return i >= 0 && argv[i + 1] && !argv[i + 1].startsWith("--") ? argv[i + 1] : def;
};
const FROM_JSON = argv.includes("--from-json");
const JSON_ONLY = argv.includes("--json-only");
const MODEL = flag("model", "claude-sonnet-5");
const OUT_OVERRIDE = flag("out", null);

if (!inputPath || !fs.existsSync(inputPath)) {
  console.error(`입력 파일을 찾을 수 없습니다: ${inputPath}`);
  process.exit(1);
}

/* ==========================================================================
 * [1] 전사 → 구조화 (Anthropic API)
 * ========================================================================*/
const EXTRACT_PROMPT = `너는 회의 녹음 전사본을 사내 표준 회의록으로 정리하는 서기다.
아래 전사본을 읽고 JSON 하나만 출력한다. 설명, 머리말, 마크다운 코드펜스 없이 JSON만.

스키마:
{
  "meeting": {
    "주제": "회의 주제 한 줄",
    "주관팀": "", "작성자": "", "장소": "",
    "회의일": "YYYY.MM.DD (요일)",
    "시간": "HH:MM ~ HH:MM",
    "회의구분": "정보전달 | 이해조정 | 의견교환 | 문제해결 | 기타 중 하나",
    "참석자": ["[소속] 이름 직급", "..."]
  },
  "sections": [
    {
      "title": "1. 주요 설명 사항",
      "subs": [
        { "title": "1) 소주제", "items": ["항목 문장.", "항목 문장."] }
      ]
    }
  ],
  "actions": [
    { "no": "1", "내용": "", "담당": "", "상태": "결정 | 협의 | 검토 | 확인 | 미결 | 재검토" }
  ]
}

작성 규칙:
- sections 는 보통 "1. 주요 설명 사항", "2. 협의 및 결정 사항" 두 개로 나눈다.
  결정·합의가 오간 대목은 2번으로, 배경 설명·현황 공유는 1번으로 보낸다.
- 2번 섹션의 sub title 뒤에는 "(결정)" "(협의)" "(검토)" "(미결)" "(재검토)" "(확인)" 중
  실제 논의 결과에 맞는 것을 붙인다.
- items 의 각 문장은 한 항목당 1~2문장, 120자 이내. "~함", "~됨", "~임" 체로 끝맺는다.
- 전사본에 없는 내용을 지어내지 않는다. 확인되지 않은 값은 빈 문자열로 둔다.
- 잡담, 인사, 중복 발언은 버린다. 숫자·고유명사·시스템명은 전사본 표기를 그대로 쓴다.
- actions 는 후속 조치가 필요한 건만. 없으면 빈 배열.

전사본:
---
`;

async function extractFromTranscript(transcript) {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    console.error("환경변수 ANTHROPIC_API_KEY 가 필요합니다.");
    process.exit(1);
  }

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 16000,
      messages: [{ role: "user", content: EXTRACT_PROMPT + transcript + "\n---" }],
    }),
  });

  if (!res.ok) {
    console.error(`API 오류 ${res.status}: ${await res.text()}`);
    process.exit(1);
  }

  const data = await res.json();
  const text = data.content
    .filter((b) => b.type === "text")
    .map((b) => b.text)
    .join("\n")
    .trim();

  const cleaned = text.replace(/^```(?:json)?\s*/i, "").replace(/```\s*$/, "").trim();
  try {
    return JSON.parse(cleaned);
  } catch (e) {
    const s = cleaned.indexOf("{"), t = cleaned.lastIndexOf("}");
    if (s >= 0 && t > s) return JSON.parse(cleaned.slice(s, t + 1));
    console.error("JSON 파싱 실패. 원문:\n" + cleaned.slice(0, 2000));
    process.exit(1);
  }
}

/* ==========================================================================
 * [2] 렌더링 상수
 * ========================================================================*/
const FONT = "맑은 고딕";
const BLACK = "000000";
const LINE = "555555";
const GRAY_FILL = "EFEFEF";
const GRAY_TXT = "666666";
const WHITE = "FFFFFF";

const PAGE_W = 8.2699, PAGE_H = 11.6889;   // A4 세로
const TX = 0.42, TW = 7.43, TR = TX + TW;
const BOX_BOTTOM = 11.02;
const PAGE_NUM_Y = 11.13;
const PAD_TOP = 0.16;

const BODY_PT = 10;                         // 본문 글자 크기
const LINE_MULT = 1.16;                     // 줄 간격 배수
const GAP_PT = 6;                           // "-" 항목 사이 간격 (6pt)
const GAP_IN = GAP_PT / 72;                 // 인치 환산

// 한 줄의 실제 높이는 글자 크기(em)가 아니라 폰트의 ascent+descent 기준이다.
// 맑은 고딕 실측 결과 약 1.20~1.22em → 겹침 방지를 위해 1.22로 잡는다.
const LINE_BOX = 1.22;
const LINE_H = (BODY_PT * LINE_BOX * LINE_MULT) / 72;   // 한 줄 높이(인치)

const X_SEC = TX + 0.16;
const X_SUB = TX + 0.38;
const X_BUL = TX + 0.60;                    // 불릿 텍스트박스 좌측
const W_BUL = TR - X_BUL - 0.16;            // 불릿 텍스트박스 폭
const HANG_PT = 18;                         // 행잉 인덴트 (불릿 ↔ 본문 간격)
const HANG_IN = HANG_PT / 72;

/* --- 줄 수 추정 -----------------------------------------------------------
 * 맑은 고딕 기준 전각(한글) 1글자의 실제 진행폭은 글자 크기의 약 0.90~0.92배.
 * 넘치면 글자가 잘리므로 0.95로 보수적으로 잡는다(추정이 크면 여백만 조금 남음).
 * ------------------------------------------------------------------------*/
const ADV_RATIO = 0.95;      // 전각 1글자 진행폭 / 글자 크기
const TEXT_INSET = 0;        // margin:0 으로 텍스트박스 내부 여백 제거

function widthUnits(s) {
  let u = 0;
  for (const ch of s) u += /[\x00-\x7F]/.test(ch) ? 0.5 : 1.0;
  return u;
}
function estLines(text, boxWidthIn, fontPt) {
  const capacity = ((boxWidthIn - TEXT_INSET) * 72) / (fontPt * ADV_RATIO);
  return Math.max(1, Math.ceil(widthUnits(text) / capacity));
}
function itemLines(text) {
  return estLines(text, W_BUL - HANG_IN, BODY_PT);
}

/* ==========================================================================
 * [3] 렌더러
 * ========================================================================*/
function render(doc, outFile) {
  const M = doc.meeting || {};
  const SECTIONS = doc.sections || [];
  const ACTIONS = doc.actions || [];
  const 참석자 = M.참석자 || [];

  const pres = new pptxgen();
  pres.defineLayout({ name: "A4P", width: PAGE_W, height: PAGE_H });
  pres.layout = "A4P";

  let pageNum = 0;

  const cell = (slide, x, y, w, h, fill) =>
    slide.addShape("rect", {
      x, y, w, h,
      fill: { color: fill || WHITE },
      line: { color: LINE, width: 0.75 },
    });

  const cellText = (slide, x, y, w, h, text, o) =>
    slide.addText(text, {
      x, y, w, h,
      align: (o && o.align) || "center",
      valign: (o && o.valign) || "middle",
      fontFace: FONT,
      fontSize: (o && o.size) || 10,
      bold: !!(o && o.bold),
      color: (o && o.color) || BLACK,
    });

  const addPageNum = (slide) => {
    pageNum += 1;
    slide.addText(String(pageNum), {
      x: TX, y: PAGE_NUM_Y, w: TW, h: 0.22, align: "center",
      fontFace: FONT, fontSize: 8.5, color: GRAY_TXT,
    });
  };

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

  cell(s1, TX, R1, C_LBL, RH, GRAY_FILL);
  cellText(s1, TX, R1, C_LBL, RH, "주 제", { bold: true, size: 10.5 });
  cell(s1, TX + C_LBL, R1, C_DIV - TX - C_LBL, RH);
  cellText(s1, TX + C_LBL + 0.1, R1, C_DIV - TX - C_LBL - 0.2, RH, M.주제 || "",
    { align: "left", size: 8.5 });

  cell(s1, C_DIV, R1, C_DIV_W, RH * 3);
  cellText(s1, C_DIV, R1 + 0.02, C_DIV_W, 0.3, "회 의 구 분", { bold: true, size: 10.5 });
  const KINDS = ["정보전달", "이해조정", "의견교환", "문제해결", "기 타"];
  const 구분 = (M.회의구분 || "").replace(/\s/g, "");
  let ky = R1 + 0.36;
  KINDS.forEach((k) => {
    s1.addText(k, {
      x: C_DIV + 0.1, y: ky, w: 1.0, h: 0.175, align: "left", valign: "middle",
      fontFace: FONT, fontSize: 8, color: BLACK,
    });
    if (k.replace(/\s/g, "") === 구분) {
      s1.addText("V", {
        x: C_DIV + 1.05, y: ky, w: 0.45, h: 0.175, align: "center", valign: "middle",
        fontFace: FONT, fontSize: 9, bold: true, color: BLACK,
      });
    }
    ky += 0.165;
  });

  const R2 = R1 + RH;
  cell(s1, TX, R2, C_LBL, RH, GRAY_FILL);
  cellText(s1, TX, R2, C_LBL, RH, "주 관 팀", { bold: true, size: 10.5 });
  cell(s1, TX + C_LBL, R2, C_MID_L - TX - C_LBL, RH);
  cellText(s1, TX + C_LBL, R2, C_MID_L - TX - C_LBL, RH, M.주관팀 || "", { size: 9.5 });
  cell(s1, C_MID_L, R2, C_MID_LW, RH, GRAY_FILL);
  cellText(s1, C_MID_L, R2, C_MID_LW, RH, "작 성 자", { bold: true, size: 10 });
  cell(s1, C_MID_L + C_MID_LW, R2, C_DIV - C_MID_L - C_MID_LW, RH);
  cellText(s1, C_MID_L + C_MID_LW, R2, C_DIV - C_MID_L - C_MID_LW, RH, M.작성자 || "", { size: 9.5 });

  const R3 = R2 + RH;
  cell(s1, TX, R3, C_LBL, RH, GRAY_FILL);
  cellText(s1, TX, R3, C_LBL, RH, "장 소", { bold: true, size: 10.5 });
  cell(s1, TX + C_LBL, R3, C_MID_L - TX - C_LBL, RH);
  cellText(s1, TX + C_LBL, R3, C_MID_L - TX - C_LBL, RH, M.장소 || "", { size: 9.5 });
  cell(s1, C_MID_L, R3, C_MID_LW, RH, GRAY_FILL);
  cellText(s1, C_MID_L, R3, C_MID_LW, RH, "회 의 일", { bold: true, size: 10 });
  cell(s1, C_MID_L + C_MID_LW, R3, C_DIV - C_MID_L - C_MID_LW, RH);
  cellText(s1, C_MID_L + C_MID_LW, R3, C_DIV - C_MID_L - C_MID_LW, RH, M.회의일 || "", { size: 9.5 });

  const R4 = R3 + RH, R4H = 0.5;
  cell(s1, TX, R4, C_LBL, R4H, GRAY_FILL);
  cellText(s1, TX, R4, C_LBL, R4H, "시 간", { bold: true, size: 10.5 });
  cell(s1, TX + C_LBL, R4, TR - TX - C_LBL, R4H);
  cellText(s1, TX + C_LBL, R4, TR - TX - C_LBL, R4H, M.시간 || "", { size: 9.5 });

  const R5 = R4 + R4H, R5H = 0.32;
  cell(s1, TX, R5, TW, R5H, GRAY_FILL);
  cellText(s1, TX, R5, TW, R5H, "참 석 자", { bold: true, size: 11 });

  const ATT_LH = 0.245;
  const R6 = R5 + R5H;
  const R6H = Math.max(0.6, 참석자.length * ATT_LH + 0.16);
  cell(s1, TX, R6, TW, R6H);
  let ay = R6 + 0.08;
  참석자.forEach((t) => {
    s1.addText(t, {
      x: TX + 0.12, y: ay, w: TW - 0.3, h: ATT_LH, valign: "middle",
      fontFace: FONT, fontSize: 9.5, color: BLACK,
    });
    ay += ATT_LH;
  });

  const R7 = R6 + R6H, R7H = 0.32;
  cell(s1, TX, R7, TW, R7H, GRAY_FILL);
  cellText(s1, TX, R7, TW, R7H, "회 의 내 용", { bold: true, size: 11 });

  /* ---------- 본문 흐름 엔진 ---------- */
  let cur = { slide: s1, boxTop: R7 + R7H, y: R7 + R7H + PAD_TOP };

  function closeBox(bottomY) {
    if (cur.boxTop === null) return;
    cur.slide.addShape("rect", {
      x: TX, y: cur.boxTop, w: TW, h: bottomY - cur.boxTop,
      fill: { type: "none" },
      line: { color: LINE, width: 0.75 },
    });
  }
  function newContentPage() {
    closeBox(BOX_BOTTOM);
    const s = pres.addSlide();
    s.background = { color: WHITE };
    addPageNum(s);
    cell(s, TX, 0.42, TW, 0.32, GRAY_FILL);
    cellText(s, TX, 0.42, TW, 0.32, "회 의 내 용", { bold: true, size: 11 });
    cur = { slide: s, boxTop: 0.74, y: 0.74 + PAD_TOP };
  }
  const roomLeft = () => BOX_BOTTOM - 0.12 - cur.y;
  function ensure(h) {
    if (h > roomLeft()) newContentPage();
  }

  function section(t) {
    ensure(0.42 + 0.31 + LINE_H);   // 대제목 + 소제목 + 최소 한 줄
    cur.slide.addText(t, {
      x: X_SEC, y: cur.y, w: TW - 0.3, h: 0.26,
      fontFace: FONT, fontSize: 11.5, bold: true, color: BLACK,
    });
    cur.y += 0.38;
  }
  function sub(t) {
    ensure(0.31 + LINE_H);          // 소제목만 홀로 남지 않도록 한 줄분 확보
    cur.slide.addText(t, {
      x: X_SUB, y: cur.y, w: TW - 0.5, h: 0.24,
      fontFace: FONT, fontSize: 10.5, bold: true, color: BLACK,
    });
    cur.y += 0.31;
  }

  /* 항목 여러 개를 하나의 텍스트박스에 문단으로 넣는다.
   * → 문단 간격(6pt)과 줄바꿈을 PowerPoint가 처리하므로 항목 사이 간격이 어긋나지 않음. */
  function drawChunk(items) {
    if (!items.length) return;
    const lines = items.map(itemLines);
    const h = lines.reduce((a, b) => a + b, 0) * LINE_H
            + (items.length - 1) * GAP_IN + 0.04;

    cur.slide.addText(
      items.map((t, i) => ({
        text: t,
        options: {
          bullet: { characterCode: "002D", indent: HANG_PT },
          paraSpaceBefore: i === 0 ? 0 : GAP_PT,
          breakLine: i < items.length - 1,
        },
      })),
      {
        x: X_BUL, y: cur.y, w: W_BUL, h,
        margin: 0,
        fontFace: FONT, fontSize: BODY_PT, color: BLACK, valign: "top",
        lineSpacingMultiple: LINE_MULT,
      }
    );
    cur.y += h;
  }

  /* 남은 높이에 맞춰 항목을 잘라 담고, 넘치면 페이지를 넘긴다. */
  function items(list) {
    let buf = [], bufH = 0;
    const flush = () => { drawChunk(buf); buf = []; bufH = 0; };

    list.forEach((t) => {
      const h = itemLines(t) * LINE_H + (buf.length ? GAP_IN : 0);
      if (buf.length && bufH + h > roomLeft()) {
        flush();
        newContentPage();
      } else if (!buf.length && h > roomLeft()) {
        newContentPage();
      }
      buf.push(t);
      bufH += buf.length === 1 ? itemLines(t) * LINE_H : h;
    });
    flush();
  }

  /* ---------- 본문 출력 ---------- */
  SECTIONS.forEach((sec) => {
    section(sec.title);
    (sec.subs || []).forEach((sb) => {
      const list = (sb.items || []).map((t) => String(t || "").trim()).filter(Boolean);
      sub(sb.title);
      items(list);
      cur.y += 0.14;   // 소주제 블록 사이 여백
    });
  });

  /* ---------- Action Item ---------- */
  if (ACTIONS.length) {
    const AI_HDR_H = 0.32, AI_ROW_HDR = 0.3;
    const CW_NO = 0.55, CW_OWNER = 1.5, CW_STAT = 0.8;
    const CW_ITEM = TW - CW_NO - CW_OWNER - CW_STAT;
    const rowH = ACTIONS.map((a) =>
      Math.max(0.3, estLines(a.내용 || "", CW_ITEM - 0.24, 9.5) * 0.2 + 0.1));
    const aiTotal = AI_HDR_H + AI_ROW_HDR + rowH.reduce((a, b) => a + b, 0);

    if (cur.y + 0.2 + aiTotal > BOX_BOTTOM) {
      closeBox(BOX_BOTTOM);
      const s = pres.addSlide();
      s.background = { color: WHITE };
      addPageNum(s);
      cur = { slide: s, boxTop: null, y: 0.42 };
    } else {
      closeBox(cur.y + 0.06);
      cur.y += 0.06;
      cur.boxTop = null;
    }

    const ai = cur.slide;
    let aY = cur.y;

    cell(ai, TX, aY, TW, AI_HDR_H, GRAY_FILL);
    ai.addText("Action Item", {
      x: TX + 0.12, y: aY, w: TW - 0.3, h: AI_HDR_H,
      align: "left", valign: "middle",
      fontFace: FONT, fontSize: 11, bold: true, color: BLACK,
    });
    aY += AI_HDR_H;

    const XN = TX, XI = XN + CW_NO, XO = XI + CW_ITEM, XS = XO + CW_OWNER;
    [[XN, CW_NO, "No"], [XI, CW_ITEM, "내용"], [XO, CW_OWNER, "담당"], [XS, CW_STAT, "상태"]]
      .forEach(([x, w, label], i) => {
        cell(ai, x, aY, w, AI_ROW_HDR, GRAY_FILL);
        cellText(ai, i === 1 ? x + 0.12 : x, aY, i === 1 ? w - 0.2 : w, AI_ROW_HDR, label,
          { bold: true, size: 9, color: GRAY_TXT, align: i === 1 ? "left" : "center" });
      });
    aY += AI_ROW_HDR;

    ACTIONS.forEach((a, i) => {
      const h = rowH[i];
      cell(ai, XN, aY, CW_NO, h);
      cell(ai, XI, aY, CW_ITEM, h);
      cell(ai, XO, aY, CW_OWNER, h);
      cell(ai, XS, aY, CW_STAT, h);
      cellText(ai, XN, aY, CW_NO, h, String(a.no ?? i + 1), { size: 9.5 });
      ai.addText(a.내용 || "", {
        x: XI + 0.12, y: aY, w: CW_ITEM - 0.24, h,
        align: "left", valign: "middle",
        fontFace: FONT, fontSize: 9.5, color: BLACK,
      });
      cellText(ai, XO, aY, CW_OWNER, h, a.담당 || "", { size: 9 });
      cellText(ai, XS, aY, CW_STAT, h, a.상태 || "", { size: 9.5, bold: true });
      aY += h;
    });
  } else {
    closeBox(BOX_BOTTOM);
  }

  return pres.writeFile({ fileName: outFile });
}

/* ==========================================================================
 * [4] main
 * ========================================================================*/
function defaultOutName(doc) {
  const d = (doc.meeting && doc.meeting.회의일 || "").replace(/[^\d]/g, "").slice(2, 8);
  return `회의록_${d || "출력"}.pptx`;
}

(async () => {
  let doc;

  if (FROM_JSON) {
    doc = JSON.parse(fs.readFileSync(inputPath, "utf8"));
  } else {
    const transcript = fs.readFileSync(inputPath, "utf8");
    console.log(`전사본 ${transcript.length}자 → 구조화 중 (${MODEL}) ...`);
    doc = await extractFromTranscript(transcript);

    const jsonOut = path.join(
      path.dirname(inputPath),
      path.basename(inputPath).replace(/\.[^.]+$/, "") + "_회의록.json"
    );
    fs.writeFileSync(jsonOut, JSON.stringify(doc, null, 2), "utf8");
    console.log(`JSON 저장 : ${jsonOut}`);
    if (JSON_ONLY) return;
  }

  const outFile = OUT_OVERRIDE || defaultOutName(doc);
  await render(doc, outFile);
  console.log(`생성 완료 : ${outFile}`);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
