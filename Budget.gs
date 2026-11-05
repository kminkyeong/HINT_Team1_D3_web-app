/**************************************************
 * 팀 예산 및 영수증 모듈 - Budget.gs
 * Google Sheets + Gemini API (영수증 OCR, 무료 등급)
 *
 * 팀 목록과 월 예산은 '팀예산' 시트가 기준입니다.
 * 아래 TEAMS는 시트가 비어 있을 때 처음 한 번만 채우는 기본값입니다.
 **************************************************/

const CONFIG = {
  BUDGET_SHEET: '팀예산',
  EXPENSE_SHEET: '지출내역',
  TEAMS: [
    '개발팀', '기획팀', '영업팀',
    '마케팅팀', '인사팀', '경영지원팀'
  ],
  DEFAULT_BUDGET: 2000000
};

const EXPENSE_HEADERS = [
  'ID', '등록일시', '결제일', '팀명',
  '상호명', '사용용도', '결제금액',
  '비고', 'OCR원문', '분석JSON'
];

const PURPOSES = [
  '회의 및 공간 대관',
  '식비 및 회식',
  '카페 및 다과',
  '교통 및 출장',
  '사무용품',
  '교육 및 도서',
  '소프트웨어 및 서비스',
  '비품 및 장비',
  '기타'
];

/* ===== 스프레드시트 설정 ===== */

function sheet_(name, headers) {
  const ss = ss_();

  let sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
  }

  if (!sh.getLastRow()) {
    sh.appendRow(headers);
  }

  return sh;
}

/** 관리자용: 메뉴에서 실행. 시트가 없으면 만들고 서식을 다시 적용 (기존 기록 유지) */
function setupBudget() {
  assertOwner_();
  const s = ensureBudgetSheets_();
  formatBudgetSheets_(s.b, s.e);
  return '설정 완료: 기존 기록 유지';
}

/**
 * 시트가 이미 준비돼 있으면 아무것도 하지 않음 (페이지를 열 때마다 호출해도 가벼움).
 * 처음 한 번만 시트 생성 + 기본 팀 입력 + 서식 적용.
 * 팀 행은 '팀예산' 시트에 팀이 하나도 없을 때만 채우므로, 지운 팀이 다시 생기지 않습니다.
 */
function ensureBudgetSheets_() {
  const ss = ss_();
  let b = ss.getSheetByName(CONFIG.BUDGET_SHEET);
  let e = ss.getSheetByName(CONFIG.EXPENSE_SHEET);

  if (b && e && b.getLastRow() > 1 && e.getLastRow() > 0) {
    return { b: b, e: e };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    b = sheet_(CONFIG.BUDGET_SHEET, ['팀명', '월 예산']);
    e = sheet_(CONFIG.EXPENSE_SHEET, EXPENSE_HEADERS);

    if (b.getLastRow() < 2) {
      b.getRange(2, 1, CONFIG.TEAMS.length, 2).setValues(
        CONFIG.TEAMS.map(t => [t, CONFIG.DEFAULT_BUDGET])
      );
      formatBudgetSheets_(b, e);
    }

    return { b: b, e: e };
  } finally {
    lock.releaseLock();
  }
}

/** '팀예산' 시트 A열의 팀 목록 (비어 있으면 기본값) */
function teams_() {
  const b = ss_().getSheetByName(CONFIG.BUDGET_SHEET);

  if (!b || b.getLastRow() < 2) {
    return CONFIG.TEAMS.slice();
  }

  const list = b.getRange(2, 1, b.getLastRow() - 1, 1)
    .getDisplayValues().flat()
    .map(s => String(s).trim())
    .filter(Boolean);

  return list.length ? [...new Set(list)] : CONFIG.TEAMS.slice();
}

function formatBudgetSheets_(b, e) {
  b.getRange('B:B')
    .setNumberFormat('#,##0"원"');

  e.getRange('C:C')
    .setNumberFormat('yyyy-mm-dd');

  e.getRange('G:G')
    .setNumberFormat('#,##0"원"');

  [b, e].forEach(sh => {
    sh.setFrozenRows(1);

    sh.getRange(
      1, 1, 1, sh.getLastColumn()
    )
      .setBackground('#1f2937')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  });
}

/* ===== 공통 함수 ===== */

function dateText_(date) {
  return Utilities.formatDate(
    date, appTz_(), 'yyyy-MM-dd'
  );
}

function monthText_(date) {
  return Utilities.formatDate(
    date, appTz_(), 'yyyy-MM'
  );
}

function validDate_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s))) {
    return false;
  }

  const [y, m, d] = s.split('-').map(Number);
  const x = new Date(y, m - 1, d);

  return (
    x.getFullYear() === y &&
    x.getMonth() === m - 1 &&
    x.getDate() === d
  );
}

function cleanCell_(value, max) {
  let s = String(
    value == null ? '' : value
  )
    .replace(/\u0000/g, '')
    .trim()
    .slice(0, max || 500);

  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function getInitialData() {
  assertAccess_();
  ensureBudgetSheets_();

  const month = monthText_(new Date());
  const teams = teams_();

  return {
    teams: teams,
    currentMonth: month,
    dashboard: getDashboard(
      month, teams[0]
    )
  };
}

/* ===== Gemini OCR ===== */

/**
 * 영수증 이미지 → Gemini가 글자를 그대로 옮겨 적음(OCR) → 기존 규칙(parseReceipt_)으로 날짜·금액·상호 추출
 * Gemini에게는 '글자 옮겨 적기'만 시키고, 해석은 아래 규칙 코드가 담당합니다.
 */
function analyzeReceipt(base64Image, mimeType) {
  assertAccess_();

  if (
    !['image/jpeg', 'image/png', 'image/webp']
      .includes(mimeType)
  ) {
    throw new Error(
      'JPG, PNG, WEBP만 지원합니다.'
    );
  }

  if (
    typeof base64Image !== 'string' ||
    !/^[A-Za-z0-9+/]+={0,2}$/.test(base64Image) ||
    base64Image.length > 12500000
  ) {
    throw new Error(
      '이미지 데이터 또는 크기가 올바르지 않습니다.'
    );
  }

  // Main.gs 의 gemini_() 로 이미지 + 지시문 전송
  const rawText = gemini_([
    { inlineData: { mimeType: mimeType, data: base64Image } },
    {
      text:
        '이 영수증 사진에 적힌 글자를 위에서 아래로, 줄 단위 그대로 옮겨 적으세요.\n' +
        '- 설명, 요약, 표 서식(마크다운) 없이 영수증의 텍스트만 출력하세요.\n' +
        '- 숫자, 쉼표, 날짜, 금액은 보이는 그대로 적으세요.\n' +
        '- 영수증이 아니거나 글자가 없으면 아무것도 출력하지 마세요.'
    }
  ]).replace(/^```[a-z]*\n?|```$/g, '').trim();

  if (!rawText.trim()) {
    throw new Error(
      '영수증에서 문자를 인식하지 못했습니다.'
    );
  }

  const parsed = parseReceipt_(rawText);

  return {
    success: true,
    parsed: parsed,
    rawText: rawText,
    json: JSON.stringify(parsed, null, 2)
  };
}

/* ===== OCR 텍스트 전처리 ===== */

function receiptLines_(text) {
  return String(text || '')
    .replace(/\u00a0/g, ' ')
    .split(/\r?\n/)
    .map(s =>
      s.replace(/[\t ]+/g, ' ').trim()
    )
    .filter(Boolean);
}

function compact_(s) {
  return String(s || '')
    .replace(/\s+/g, '')
    .toLowerCase();
}

function receiptBodyLines_(text) {
  const lines = receiptLines_(text);

  // 영수증 제목 앞의 배경 글씨 제외
  const start = lines.findIndex(s =>
    /영\s*수\s*증|신용\s*카드\s*승인\s*전표|카드\s*매출\s*전표|현금\s*영수\s*증/i
      .test(s)
  );

  return start < 0
    ? lines
    : lines.slice(start);
}

/* ===== 결제일 추출 ===== */

function extractDate_(text) {
  const lines = receiptBodyLines_(text);

  const pattern =
    /(20\d{2}|\d{2})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})/;

  const sorted = lines
    .filter(s =>
      /결제일|거래일|승인일|판매일|일시|날짜|발행일/i
        .test(s)
    )
    .concat(lines);

  for (const line of sorted) {
    const m = line.match(pattern);

    if (!m) continue;

    const y = +m[1] < 100
      ? +m[1] + 2000
      : +m[1];

    const s =
      `${y}-` +
      `${String(+m[2]).padStart(2, '0')}-` +
      `${String(+m[3]).padStart(2, '0')}`;

    if (validDate_(s)) {
      return s;
    }
  }

  return '';
}

/* ===== 결제금액 분석 규칙 ===== */

const MONEY_RULES_ = [
  [
    /최종\s*결제|실\s*결제|총\s*결제|실\s*지불|받으실\s*금액|받을\s*금액/i,
    125
  ],
  [
    /신용\s*카드|체크\s*카드|카드\s*결제|카드\s*승인\s*금액|카드\s*금액|현금\s*결제/i,
    115
  ],
  [
    /승인\s*금액|결제\s*금액|청구\s*금액|지불\s*금액|결제\s*합계/i,
    110
  ],
  [
    /총\s*합계|합\s*계|총\s*금액|총\s*액|결제\s*총액/i,
    100
  ]
];

const MONEY_EXCLUDE_ =
  /공급\s*가액|공급\s*금액|부가\s*세|v\s*a\s*t|세액|면세|과세|할인|쿠폰|적립|포인트|거스름|잔돈|잔액|봉사료|수량|단가|할부|취소|환불|카드\s*번호|승인\s*번호|사업자\s*(?:등록)?\s*번호|전화\s*번호|주문\s*번호|거래\s*번호|주문\s*내역/i;

/* ===== 금액 후보 추출 ===== */

function amountTokens_(line) {
  const result = [];

  const re =
    /(?:^|[^\d.,:\/-])([₩￦]?\s*\d{1,3}(?:,\d{3})+|[₩￦]?\s*\d{1,8})(?:\s*원)?(?=$|[^\d.,:\/-])/g;

  let m;

  while ((m = re.exec(line)) !== null) {
    const raw = m[1];

    const value = Number(
      raw.replace(/[^\d]/g, '')
    );

    const start =
      m.index + m[0].indexOf(raw);

    const end = start + raw.length;

    const currency =
      /[₩￦]/.test(raw) ||
      /^\s*원/.test(line.slice(end));

    if (
      !Number.isSafeInteger(value) ||
      value <= 0 ||
      value > 100000000 ||
      (value < 100 && !currency)
    ) {
      continue;
    }

    result.push({
      value: value,
      start: start,
      end: end,
      currency: currency,
      comma: raw.includes(',')
    });
  }

  return result;
}

/* ===== 최종 결제금액 결정 ===== */

function extractAmountDetail_(text) {
  const lines = receiptBodyLines_(text);
  const found = [];

  lines.forEach((line, index) => {
    // 부가세, 할인금액 등 제외
    if (MONEY_EXCLUDE_.test(line)) {
      return;
    }

    const labels = [];

    MONEY_RULES_.forEach(([regex, score]) => {
      const r = new RegExp(
        regex.source, 'gi'
      );

      let m;

      while ((m = r.exec(line)) !== null) {
        labels.push({
          start: m.index,
          end: r.lastIndex,
          score: score
        });
      }
    });

    const tokens = amountTokens_(line);

    tokens.forEach(t => {
      const prior = labels
        .filter(l =>
          l.end <= t.start &&
          t.start - l.end <= 26 &&
          !/\d/.test(
            line.slice(l.end, t.start)
          )
        )
        .sort((a, b) =>
          b.end - a.end ||
          b.score - a.score
        );

      if (!prior.length) {
        return;
      }

      const score =
        prior[0].score +
        (t.currency ? 3 : 0) +
        (t.comma ? 2 : 0);

      found.push({
        amount: t.value,
        score: score,
        line: line,
        index: index
      });
    });

    // 결제 레이블과 금액이 다른 줄에 있는 경우
    if (
      !tokens.length &&
      labels.length &&
      index + 1 < lines.length
    ) {
      const next = lines[index + 1];

      if (
        /^(?:[₩￦]\s*)?\d[\d,]*\s*원?$/.test(next)
      ) {
        const t = amountTokens_(next);

        if (t.length === 1) {
          found.push({
            amount: t[0].value,
            score:
              Math.max(
                ...labels.map(l => l.score)
              ) - 5,
            line: line + ' / ' + next,
            index: index + 1
          });
        }
      }
    }
  });

  // 같은 금액이 여러 곳에서 등장하면 그룹화
  const groups = new Map();

  found.forEach(c => {
    if (!groups.has(c.amount)) {
      groups.set(c.amount, {
        amount: c.amount,
        score: 0,
        lines: [],
        index: c.index
      });
    }

    const g = groups.get(c.amount);

    g.score = Math.max(g.score, c.score);
    g.index = Math.max(g.index, c.index);

    if (!g.lines.includes(c.line)) {
      g.lines.push(c.line);
    }
  });

  const ranked = [...groups.values()]
    .map(g => ({
      ...g,
      score:
        g.score +
        Math.min(g.lines.length - 1, 2) * 4
    }))
    .sort((a, b) =>
      b.score - a.score ||
      b.index - a.index
    );

  if (!ranked.length) {
    return {
      amount: null,
      confidence: 'low',
      selectedLine: '',
      candidates: []
    };
  }

  const best = ranked[0];

  const margin = ranked[1]
    ? best.score - ranked[1].score
    : Infinity;

  const confidence =
    best.score >= 110 && margin >= 12
      ? 'high'
      : best.score >= 95 && margin >= 8
        ? 'medium'
        : 'low';

  return {
    amount:
      confidence === 'low'
        ? null
        : best.amount,

    confidence: confidence,
    selectedLine: best.lines[0],

    candidates: ranked.slice(0, 5).map(x => ({
      amount: x.amount,
      score: x.score,
      lines: x.lines
    }))
  };
}

function extractAmount_(text) {
  return extractAmountDetail_(text).amount;
}

/* ===== 상호명 전처리 ===== */

function cleanMerchant_(s) {
  return String(s || '')
    .replace(
      /^[\s:：\-]+|[\s:：\-]+$/g, ''
    )
    .replace(/^\[|\]$/g, '')
    .trim()
    .slice(0, 80);
}

function merchantInvalid_(name) {
  const s = compact_(name);

  if (
    name.length < 2 ||
    name.length > 75 ||
    !/[가-힣a-z]/i.test(name)
  ) {
    return true;
  }

  if (
    /영수증|승인전표|매출전표|신용카드|체크카드|결제금액|승인금액|합계|부가세|공급가액|사업자등록|대표자|^대표[:：]|전화|tel|fax|주소|카드번호|승인번호|주문번호|거래일시|결제일시|수량|단가|고객용|^상품$|^품목$|거래번호|장뭄/
      .test(s)
  ) {
    return true;
  }

  if (
    /^(visa|mastercard|bc카드|국민카드|신한카드|현대카드|삼성카드|롯데카드)$/i
      .test(s)
  ) {
    return true;
  }

  if (
    /\d{2,4}[./-]\d{1,2}|\d{2,3}-\d{3,4}-\d{4}/
      .test(name)
  ) {
    return true;
  }

  return false;
}

/* ===== 상호명 후보 추출 ===== */

function extractMerchantDetail_(text) {
  const lines = receiptBodyLines_(text);
  const candidates = [];

  // 정상적으로 인식된 상호명 레이블
  const exact =
    /^(?:상호명|상호|가맹점명|가맹점|매장명|업체명|사업장명|상점명)\s*[:：\-]?\s*(.*)$/i;

  // OCR 오인식이 발생한 상호명 레이블
  const fuzzy =
    /^(?:상차|상호자|사망차|사망자|상흐|상흐명)\s*[:：\-]\s*(.+)$/i;

  lines.forEach((line, index) => {
    let m = line.match(exact);
    let score = 140;
    let source = 'exact-label';

    if (!m) {
      m = line.match(fuzzy);
      score = 110;
      source = 'ocr-corrected-label';
    }

    if (!m) {
      return;
    }

    let name = cleanMerchant_(m[1]);

    if (!name && lines[index + 1]) {
      name = cleanMerchant_(
        lines[index + 1]
      );
    }

    if (merchantInvalid_(name)) {
      return;
    }

    candidates.push({
      name: name,
      score: score,
      line: line,
      index: index,
      source: source
    });
  });

  // 명시적 상호명 레이블이 없을 경우
  if (!candidates.length) {
    lines.forEach((line, index) => {
      if (index > 22) {
        return;
      }

      const name = cleanMerchant_(line);

      if (merchantInvalid_(name)) {
        return;
      }

      const around = lines.slice(
        Math.max(0, index - 2),
        Math.min(lines.length, index + 5)
      ).join(' ');

      let score = 30;

      // 사업자 정보 주변이면 가산점
      if (
        /사업자|주소|대표|전화|tel|fax/i
          .test(around)
      ) {
        score += 20;
      }

      // 상호명이나 업종으로 보이는 표현
      if (
        /식당|치킨|피자|카페|커피|스타벅스|마트|호텔|서점|문구|편의점|회의실|스터디룸|대관|센터|주식회사|\(주\)|㈜|점$/i
          .test(name)
      ) {
        score += 25;
      }

      if (!/\d/.test(name)) {
        score += 5;
      }

      candidates.push({
        name: name,
        score: score,
        line: line,
        index: index,
        source: 'inferred'
      });
    });
  }

  candidates.sort((a, b) =>
    b.score - a.score ||
    a.index - b.index
  );

  if (!candidates.length) {
    return {
      merchant: '',
      confidence: 'low',
      candidates: []
    };
  }

  const top = candidates[0];
  const second = candidates[1];

  const confidence =
    top.source === 'exact-label'
      ? 'high'
      : top.source === 'ocr-corrected-label'
        ? 'medium'
        : (
            top.score >= 75 &&
            (
              !second ||
              top.score - second.score >= 10
            )
              ? 'medium'
              : 'low'
          );

  return {
    merchant:
      confidence === 'low'
        ? ''
        : top.name,

    confidence: confidence,

    candidates: candidates.slice(0, 5).map(c => ({
      name: c.name,
      score: c.score,
      line: c.line,
      source: c.source
    }))
  };
}

function extractMerchant_(text) {
  return extractMerchantDetail_(text).merchant;
}

/* ===== 지출 분야 자동 분류 ===== */

const PURPOSE_RULES_ = [
  [
    '회의 및 공간 대관',
    /회의실|스터디룸|회의공간|공간\s*대관|대관료|세미나실|컨퍼런스룸|공유\s*오피스|meeting\s*room|coworking|위메트/i
  ],
  [
    '식비 및 회식',
    /식당|치킨|피자|맥주|고깃집|김밥|햄버거|국밥|회식|음식점|레스토랑|restaurant|분식|도시락/i
  ],
  [
    '카페 및 다과',
    /카페|커피|스타벅스|투썸|이디야|빽다방|베이커리|디저트|아메리카노|카페라떼|음료|coffee|cafe/i
  ],
  [
    '교통 및 출장',
    /택시|주차|주유|고속도로|철도|기차|ktx|버스|항공|숙박|호텔|교통비|출장비|toll/i
  ],
  [
    '사무용품',
    /문구|복사용지|프린터|잉크|토너|사무용품|모닝글로리|오피스디포|파일철|필기구/i
  ],
  [
    '교육 및 도서',
    /서점|도서|교재|교육|강의|세미나|컨퍼런스|워크숍|연수|book/i
  ],
  [
    '소프트웨어 및 서비스',
    /software|cloud|hosting|openai|google\s*workspace|구독|라이선스|서버|클라우드|소프트웨어/i
  ],
  [
    '비품 및 장비',
    /전자기기|모니터|키보드|마우스|컴퓨터|노트북|장비|비품|기자재/i
  ]
];

function inferPurpose_(text) {
  const input = String(text || '');

  for (const [name, regex] of PURPOSE_RULES_) {
    if (regex.test(input)) {
      return name;
    }
  }

  return '기타';
}

/* ===== 최종 JSON 생성 ===== */

function parseReceipt_(text) {
  const amount = extractAmountDetail_(text);
  const merchant = extractMerchantDetail_(text);

  // 상호명이 비어 있어도 전체 OCR 문맥으로 분류
  const purpose = inferPurpose_(
    merchant.merchant +
    '\n' +
    receiptBodyLines_(text).join('\n')
  );

  return {
    paymentDate: extractDate_(text),
    amount: amount.amount,
    merchant: merchant.merchant,
    purpose: purpose,
    note: '',
    needsReview: true,

    amountConfidence: amount.confidence,
    merchantConfidence: merchant.confidence,

    amountSelectedLine: amount.selectedLine,
    amountCandidates: amount.candidates,
    merchantCandidates: merchant.candidates
  };
}

/* ===== 지출 내역 저장 ===== */

function saveExpense(data) {
  assertAccess_();

  if (
    !data ||
    !teams_().includes(data.team)
  ) {
    throw new Error(
      '올바른 팀을 선택하세요.'
    );
  }

  if (
    !validDate_(
      String(data.paymentDate || '')
    )
  ) {
    throw new Error(
      '결제일을 확인하세요.'
    );
  }

  const amount = Number(data.amount);

  if (
    !Number.isSafeInteger(amount) ||
    amount <= 0 ||
    amount > 100000000
  ) {
    throw new Error(
      '결제금액을 확인하세요.'
    );
  }

  const purpose = cleanCell_(data.purpose, 80);

  if (!purpose) {
    throw new Error(
      '사용 용도를 입력하세요.'
    );
  }

  const sh = ss_()
    .getSheetByName(CONFIG.EXPENSE_SHEET);

  if (!sh) {
    throw new Error(
      'setup()을 먼저 실행하세요.'
    );
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  let id;

  try {
    id = Utilities.getUuid();

    const [y, m, d] =
      data.paymentDate.split('-').map(Number);

    const record = {
      id: id,
      team: data.team,
      paymentDate: data.paymentDate,
      merchant: cleanCell_(
        data.merchant, 80
      ),
      purpose: purpose,
      amount: amount,
      note: cleanCell_(
        data.note, 300
      )
    };

    sh.appendRow([
      id,
      new Date(),
      new Date(y, m - 1, d, 12),
      record.team,
      record.merchant,
      purpose,
      amount,
      record.note,
      cleanCell_(data.rawText, 30000),
      JSON.stringify(record)
    ]);

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  // 이번 지출로 예산을 처음 넘었으면 알림 메일 (메일 실패는 저장 결과에 영향 없음)
  let alerted = false;

  try {
    alerted = notifyBudgetExceeded_(
      data.team, data.paymentDate.slice(0, 7), amount
    );
  } catch (err) {
    Logger.log('예산 초과 알림 실패: ' + err.message);
  }

  return {
    success: true,
    id: id,
    alerted: alerted,
    message: alerted
      ? '지출 내역이 저장되었습니다. 예산을 초과해 담당자에게 알림 메일을 보냈습니다.'
      : '지출 내역이 저장되었습니다.'
  };
}

/**
 * 예산 초과 알림.
 * 받는 사람: 스크립트 속성 BUDGET_ALERT_EMAILS(쉼표 구분), 없으면 스크립트 소유자
 * 이번 지출 전에는 예산 이내였고, 이번 지출로 초과한 경우에만 1회 발송
 */
function notifyBudgetExceeded_(team, month, amount) {
  const s = getDashboard(month, team).selected;

  if (!(s.spent > s.budget && s.spent - amount <= s.budget)) {
    return false;
  }

  const to = String(
    PropertiesService.getScriptProperties().getProperty('BUDGET_ALERT_EMAILS') ||
    ownerEmail_()
  ).trim();

  if (!to) return false;

  const won = n => Number(n).toLocaleString('ko-KR') + '원';

  MailApp.sendEmail({
    to: to,
    subject: '[예산 초과] ' + team + ' ' + month + ' 예산을 초과했습니다',
    htmlBody:
      '<p><b>' + escHtml_(team) + '</b>의 ' + escHtml_(month) + ' 지출이 월 예산을 초과했습니다.</p>' +
      '<ul><li>월 예산: ' + won(s.budget) + '</li>' +
      '<li>누적 지출: ' + won(s.spent) + ' (' + s.rate + '%)</li>' +
      '<li>초과 금액: ' + won(s.spent - s.budget) + '</li></ul>' +
      '<p><a href="' + escHtml_(appBase_() + '?page=budget') + '">예산 관리 화면 열기</a></p>'
  });

  return true;
}

/* ===== 팀별 월간 대시보드 ===== */

function getDashboard(month, selectedTeam) {
  assertAccess_();

  const TEAMS = teams_();

  if (
    !/^\d{4}-(0[1-9]|1[0-2])$/.test(String(month))
  ) {
    throw new Error(
      '조회 월을 확인하세요.'
    );
  }

  if (!TEAMS.includes(selectedTeam)) {
    throw new Error(
      '올바른 팀을 선택하세요.'
    );
  }

  const ss = ss_();

  const b = ss.getSheetByName(
    CONFIG.BUDGET_SHEET
  );

  const e = ss.getSheetByName(
    CONFIG.EXPENSE_SHEET
  );

  if (!b || !e) {
    throw new Error(
      'setup()을 먼저 실행하세요.'
    );
  }

  const budgets = Object.fromEntries(
    TEAMS.map(t => [
      t, CONFIG.DEFAULT_BUDGET
    ])
  );

  if (b.getLastRow() > 1) {
    b.getRange(
      2, 1, b.getLastRow() - 1, 2
    ).getValues().forEach(r => {
      const name = String(r[0]).trim();
      if (TEAMS.includes(name)) {
        budgets[name] = Number(r[1]) || 0;
      }
    });
  }

  const summaries = {};

  TEAMS.forEach(t => {
    summaries[t] = {
      team: t,
      budget: budgets[t],
      spent: 0,
      remaining: budgets[t],
      count: 0,
      rate: 0,
      exceeded: false
    };
  });

  const records = [];

  if (e.getLastRow() > 1) {
    e.getRange(
      2, 1, e.getLastRow() - 1, 10
    ).getValues().forEach(r => {
      const date =
        r[2] instanceof Date &&
        !isNaN(r[2].getTime())
          ? dateText_(r[2])
          : String(r[2] || '').slice(0, 10);

      if (date.slice(0, 7) !== month) {
        return;
      }

      const team = String(r[3]);
      const amount = Number(r[6]) || 0;

      if (!summaries[team]) {
        return;
      }

      summaries[team].spent += amount;
      summaries[team].count++;

      if (team === selectedTeam) {
        records.push({
          id: String(r[0]),
          paymentDate: date,
          team: team,
          merchant: String(r[4] || ''),
          purpose: String(r[5] || ''),
          amount: amount,
          note: String(r[7] || '')
        });
      }
    });
  }

  const teams = TEAMS.map(t => {
    const s = summaries[t];

    s.remaining = s.budget - s.spent;

    s.rate = s.budget > 0
      ? Math.round(
          s.spent / s.budget * 1000
        ) / 10
      : (s.spent > 0 ? 100 : 0);

    s.exceeded = s.spent > s.budget;

    return s;
  });

  records.sort((a, b) =>
    b.paymentDate.localeCompare(a.paymentDate)
  );

  return {
    month: month,
    selectedTeam: selectedTeam,
    selected: summaries[selectedTeam],
    teams: teams,
    records: records.slice(0, 100),
    totalSpent: teams.reduce(
      (s, x) => s + x.spent, 0
    )
  };
}

/* ===== API 키 확인 ===== */

// Gemini API 키 점검은 Main.gs 의 testGeminiApiKey() 를 실행하세요.

/* ===== 데이터 초기화 ===== */

/**
 * 관리자 전용: 스프레드시트 메뉴 [⚠ 예산·지출 데이터 초기화]에서만 실행됩니다.
 * 'DELETE'를 직접 입력해야 진행되며, 지우기 전에 두 시트를 백업 시트로 복사합니다.
 */
function resetData() {
  assertOwner_();

  let ui;

  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    throw new Error(
      '스프레드시트 메뉴 [🗂 통합 업무 관리 > ⚠ 예산·지출 데이터 초기화]에서 실행하세요.'
    );
  }

  const answer = ui.prompt(
    '⚠ 예산·지출 데이터 초기화',
    '지출내역이 모두 삭제되고 팀예산이 기본값으로 돌아갑니다.\n' +
    '(삭제 전 백업 시트가 자동으로 만들어집니다)\n\n' +
    '계속하려면 DELETE 를 입력하세요.',
    ui.ButtonSet.OK_CANCEL
  );

  if (
    answer.getSelectedButton() !== ui.Button.OK ||
    answer.getResponseText().trim() !== 'DELETE'
  ) {
    ui.alert('초기화를 취소했습니다.');
    return;
  }

  const ss = ss_();
  const stamp = Utilities.formatDate(
    new Date(), appTz_(), 'yyyyMMdd_HHmmss'
  );

  [
    CONFIG.EXPENSE_SHEET,
    CONFIG.BUDGET_SHEET
  ].forEach(name => {
    const sh = ss.getSheetByName(name);

    if (sh && sh.getLastRow() > 1) {
      sh.copyTo(ss).setName(name + '_백업_' + stamp);

      sh.getRange(
        2, 1,
        sh.getLastRow() - 1,
        sh.getLastColumn()
      ).clearContent();
    }
  });

  const b = ss.getSheetByName(
    CONFIG.BUDGET_SHEET
  );

  if (b) {
    CONFIG.TEAMS.forEach((t, i) => {
      b.getRange(
        i + 2, 1, 1, 2
      ).setValues([
        [t, CONFIG.DEFAULT_BUDGET]
      ]);
    });
  }

  ui.alert('초기화 완료. 백업 시트: *_백업_' + stamp);
}

/* ===== 영수증 파싱 테스트 ===== */

function testReceiptParser() {
  assertOwner_();

  const cases = [
    [
      '상호명: 행복카페\n' +
      '상품금액 60,000\n' +
      '할인금액 5,000\n' +
      '신용카드 55,000',
      55000,
      '행복카페',
      '카페 및 다과'
    ],
    [
      '가맹점명: 서울문구\n' +
      '공급가액 30,000\n' +
      '부가세 3,000\n' +
      '승인금액 33,000',
      33000,
      '서울문구',
      '사무용품'
    ],
    [
      '상호명: OO식당\n' +
      '합계 45,000\n' +
      '신용카드\n' +
      '45,000',
      45000,
      'OO식당',
      '식비 및 회식'
    ],
    [
      '가지\n영 수 증\n' +
      '사망차 : 스타벅스 (역삼대로점)\n' +
      '주소: 서울\n' +
      '공급가액 39,273원\n' +
      'VAT 3,927원\n' +
      '합 계\n43,200원',
      43200,
      '스타벅스 (역삼대로점)',
      '카페 및 다과'
    ],
    [
      '영수증\n' +
      '상차 : 위메트 회의실 & 스터디룸\n' +
      '대표: 최성호\n장뭄\n' +
      '합 계\n98,000원',
      98000,
      '위메트 회의실 & 스터디룸',
      '회의 및 공간 대관'
    ]
  ];

  cases.forEach((c, i) => {
    const r = parseReceipt_(c[0]);

    Logger.log(JSON.stringify({
      test: i + 1,
      pass:
        r.amount === c[1] &&
        r.merchant === c[2] &&
        r.purpose === c[3],
      actual: {
        amount: r.amount,
        merchant: r.merchant,
        purpose: r.purpose
      }
    }));
  });
}
