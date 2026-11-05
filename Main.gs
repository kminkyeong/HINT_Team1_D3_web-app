
/**
 * 통합 업무 관리 웹앱
 * 공통 진입점 및 안전한 데이터 초기화
 *
 * 기존 Meeting.gs / Budget.gs / Minutes.gs는 수정하지 않습니다.
 */

function appBase_() {
  return ScriptApp.getService().getUrl() || '';
}

function renderBudget_() {
  const html = HtmlService
    .createHtmlOutputFromFile('IndexBudget')
    .getContent()
    .replace(/__APP_BASE__/g, appBase_());

  return HtmlService.createHtmlOutput(html)
    .setTitle('팀 예산 및 영수증')
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    );
}

function renderHome_() {
  const html = HtmlService
    .createHtmlOutputFromFile('IndexHome')
    .getContent()
    .replace(/__APP_BASE__/g, appBase_());

  return HtmlService.createHtmlOutput(html)
    .setTitle('Mobility Workspace | 통합 업무 관리')
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    );
}

function renderMinutes_() {
  const html = HtmlService
    .createHtmlOutputFromFile('IndexMinutes')
    .getContent()
    .replace(/__APP_BASE__/g, appBase_());

  return HtmlService.createHtmlOutput(html)
    .setTitle('스마트 회의록 · AI 요약')
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    );
}

/**
 * 웹앱 공통 진입점
 */

function doGet(e) {
  const p = (e && e.parameter) || {};

  if (p.m) {
    return withWorkspaceNav_(
      render_(
        p.admin ? 'admin' : 'vote',
        p.m,
        p.admin || ''
      ),
      p.admin ? 'admin' : 'vote'
    );
  }

  if (p.page === 'create') {
    return withWorkspaceNav_(
      render_('create', '', ''),
      'create'
    );
  }

  if (p.page === 'budget') {
    return withWorkspaceNav_(
      renderBudget_(),
      'budget'
    );
  }

  if (p.page === 'minutes') {
    return withWorkspaceNav_(
      renderMinutes_(),
      'minutes'
    );
  }

  return renderHome_();
}


/**
 * 최초 초기 설정
 * 기존 데이터는 유지
 */
function setup() {
  setupMeeting();
  setupBudget();
  setupMinutes();

  return '초기 설정이 완료되었습니다. 기존 데이터는 유지되었습니다.';
}

/**
 * Google Sheets 상단 메뉴
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚘 통합 업무 관리')
    .addItem(
      '웹앱 초기 설정 (기존 데이터 유지)',
      'setup'
    )
    .addSeparator()
    .addItem(
      '새 회의 만들기',
      'showCreateDialog'
    )
    .addSeparator()
    .addItem(
      '회의 시트 설정',
      'setupMeeting'
    )
    .addItem(
      '예산 시트 설정',
      'setupBudget'
    )
    .addItem(
      '회의록 시트 설정',
      'setupMinutes'
    )
    .addSeparator()
    .addItem(
      '🧹 업무 기록 초기화 (설정 유지)',
      'resetWorkRecords'
    )
    .addItem(
      '⚠️ 전체 시트 초기화 (설정 포함)',
      'resetAllWorkspaceSheets'
    )
    .addToUi();
}

/**
 * 업무 이력만 삭제
 *
 * Rooms, 팀예산의 설정은 유지
 */
function resetWorkRecords() {
  resetWorkspaceFromSheetMenu_('records');
}

/**
 * 업무 기록과 설정 모두 초기화
 */
function resetAllWorkspaceSheets() {
  resetWorkspaceFromSheetMenu_('all');
}

/**
 * Google Sheets 메뉴에서만 초기화 실행
 */

/**
 * 업무 데이터 초기화
 * Google Sheets 메뉴에서 실행
 */
function resetWorkspaceFromSheetMenu_(mode) {

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('스프레드시트를 찾을 수 없습니다.');
  }

  const props = PropertiesService.getScriptProperties();

  const recordSheets = [
    'Meetings',
    'Members',
    'Availability',
    'Reservations',
    '지출내역',
    props.getProperty('SHEET_NAME') || '회의록'
  ];

  const configSheets = [
    'Rooms',
    '팀예산'
  ];

  const targets = mode === 'all'
    ? [...recordSheets, ...configSheets]
    : recordSheets;

  // 1차 확인
  const answer = ui.alert(
    '데이터 초기화',
    mode === 'all'
      ? '업무 기록과 설정 데이터를 모두 삭제하시겠습니까?'
      : '업무 기록을 모두 삭제하시겠습니까?\n회의실 및 예산 설정은 유지됩니다.',
    ui.ButtonSet.YES_NO
  );

  if (answer !== ui.Button.YES) return;

  // 2차 확인
  const confirm = ui.prompt(
    '최종 확인',
    '초기화를 실행하려면 "초기화"를 입력하세요.',
    ui.ButtonSet.OK_CANCEL
  );

  if (
    confirm.getSelectedButton() !== ui.Button.OK ||
    confirm.getResponseText().trim() !== '초기화'
  ) {
    ui.alert('초기화가 취소되었습니다.');
    return;
  }

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    ui.alert('다른 작업이 진행 중입니다. 다시 시도해 주세요.');
    return;
  }

  const results = [];
  const errors = [];

  try {

    const uniqueTargets = [...new Set(targets)];

    uniqueTargets.forEach(name => {

      try {

        const sheet = ss.getSheetByName(name);

        if (!sheet) {
          results.push(name + ': 시트 없음');
          return;
        }

        const lastRow = sheet.getLastRow();
        const lastColumn = sheet.getLastColumn();

        if (lastRow <= 1 || lastColumn === 0) {
          results.push(name + ': 초기화할 데이터 없음');
          return;
        }

        // 헤더는 보존하고 데이터만 삭제
        const batchSize = 1000;

        for (
          let startRow = 2;
          startRow <= lastRow;
          startRow += batchSize
        ) {

          const rows = Math.min(
            batchSize,
            lastRow - startRow + 1
          );

          sheet.getRange(
            startRow,
            1,
            rows,
            lastColumn
          ).clearContent();

        }

        results.push(name + ': 초기화 완료');

      } catch (error) {

        errors.push(name + ': ' + error.message);

      }

    });

    SpreadsheetApp.flush();

  } finally {
    lock.releaseLock();
  }

  const message =
    results.join('\n') +
    (errors.length
      ? '\n\n오류:\n' + errors.join('\n')
      : '');

  ui.alert(
    errors.length ? '일부 초기화 실패' : '초기화 완료',
    message,
    ui.ButtonSet.OK
  );
}


/**
 * 통합 웹앱 공통 상단 메뉴
 *
 * 회의 일정, 팀 예산, AI 회의록 화면에
 * 동일한 메뉴를 추가합니다.
 */
function withWorkspaceNav_(output, currentPage) {

  const base = appBase_();

  const pages = [
    {
      id: 'home',
      label: '홈',
      url: base
    },
    {
      id: 'create',
      label: '회의 일정 조율',
      url: base + '?page=create'
    },
    {
      id: 'budget',
      label: '팀 예산 및 영수증',
      url: base + '?page=budget'
    },
    {
      id: 'minutes',
      label: 'AI 회의록 작성',
      url: base + '?page=minutes'
    }
  ];

  // 참석자 투표 및 관리자 화면도
  // 회의 일정 메뉴가 선택된 상태로 표시
  const activePage =
    ['vote', 'admin'].includes(currentPage)
      ? 'create'
      : currentPage;

  const escapeHtml = value =>
    String(value).replace(
      /[&<>"']/g,
      char => ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
      }[char])
    );

  const styles = `
    <style>
      .workspace-nav {
        background: #002c5f;
        color: #ffffff;
        padding: 12px 20px;
        font-family:
          Arial,
          'Malgun Gothic',
          sans-serif;
        position: relative;
        z-index: 1000;
      }

      .workspace-nav-inner {
        max-width: 1160px;
        margin: 0 auto;
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 15px;
      }

      .workspace-nav-brand {
        color: white;
        text-decoration: none;
        font-size: 15px;
        font-weight: 800;
        letter-spacing: .04em;
      }

      .workspace-nav-links {
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: 7px;
      }

      .workspace-nav-link {
        display: inline-block;
        color: #d9e8f6;
        text-decoration: none;
        font-size: 12px;
        font-weight: 600;
        padding: 10px 13px;
        border-radius: 5px;
        border: 1px solid transparent;
        background: transparent;
        transition: background .2s;
      }

      .workspace-nav-link:hover {
        background: #145184;
        color: white;
      }

      .workspace-nav-link.active {
        background: #007fa8;
        color: white;
        border-color: #00aad2;
      }

      @media (max-width: 700px) {
        .workspace-nav {
          padding: 12px;
        }

        .workspace-nav-inner {
          flex-direction: column;
          align-items: stretch;
        }

        .workspace-nav-brand {
          text-align: center;
        }

        .workspace-nav-links {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 7px;
        }

        .workspace-nav-link {
          text-align: center;
          font-size: 11px;
          padding: 11px 5px;
        }
      }
    </style>
  `;

  const links = pages.map(page => {

    const active = page.id === activePage;

    return `
      <a
        class="workspace-nav-link ${active ? 'active' : ''}"
        href="${escapeHtml(page.url)}"
        target="_top"
        ${active ? 'aria-current="page"' : ''}
      >
        ${escapeHtml(page.label)}
      </a>
    `;

  }).join('');

  const navigation = `
    <nav
      class="workspace-nav"
      aria-label="통합 업무 메뉴"
    >
      <div class="workspace-nav-inner">

        <a
          class="workspace-nav-brand"
          href="${escapeHtml(base)}"
          target="_top"
        >
          MOBILITY WORKSPACE
        </a>

        <div class="workspace-nav-links">
          ${links}
        </div>

      </div>
    </nav>
  `;

  let html = output.getContent();

  // HTML head에 공통 메뉴 스타일 삽입
  if (/<\/head\s*>/i.test(html)) {
    html = html.replace(
      /<\/head\s*>/i,
      styles + '</head>'
    );
  }

  // body 시작 지점에 공통 메뉴 삽입
  // 기존 body의 data 속성은 유지
  if (/<body\b[^>]*>/i.test(html)) {
    html = html.replace(
      /<body\b[^>]*>/i,
      match => match + navigation
    );
  } else {
    throw new Error(
      '화면에 body 태그가 없어 메뉴를 추가할 수 없습니다.'
    );
  }

  return HtmlService
    .createHtmlOutput(html)
    .setTitle('Mobility Workspace')
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    );
}
