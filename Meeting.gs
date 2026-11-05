/**
 * 회의 일정 관리 - Meeting.gs
 *
 * 업무 흐름
 *  ① 회의 등록       createMeeting       : 회의명·날짜·시간·회의실·참석자·안건·사전 할 일 입력
 *  ② 초대 발송       (createMeeting 안)   : 구글 캘린더 초대 + 참석자별 '개인 응답 링크' 메일
 *  ③ 참석 응답       submitResponse      : 참석/불참/미정 선택 + 맡을 사전 할 일 체크
 *  ④ 상태 자동 변경  updateMeetingStatuses: 예정 → 진행중 → 완료 (10분마다 트리거로 실행)
 *                                         완료되면 주최자에게 '회의록 작성 요청' 메일
 *  ⑤⑥ AI 회의록·할 일 배포는 Minutes.gs 에서 처리
 *
 * 시트 (메뉴 ① 초기 설정으로 자동 생성)
 *  - Meetings  : 회의 기본 정보와 상태
 *  - Attendees : 회의별 참석자, 개인 응답 토큰, 참석 여부
 *  - Tasks     : 할 일 (phase: 사전 = 회의 전 등록 / 회의후 = 회의록에서 등록)
 *  - Rooms     : 회의실 목록 (정원, 회의실 캘린더 ID)
 */

const HEADERS = {
  Meetings: ['meetingId', 'title', 'date', 'startTime', 'endTime', 'roomId', 'status', 'eventId',
    'organizerEmail', 'adminKey', 'agenda', 'minutesRequested', 'minutesSavedAt', 'createdAt'],
  Attendees: ['meetingId', 'name', 'email', 'token', 'response', 'respondedAt'],
  Tasks: ['taskId', 'meetingId', 'title', 'assignee', 'assigneeEmail', 'dueDate', 'phase', 'done', 'createdAt'],
  Rooms: ['roomId', 'roomName', 'capacity', 'calendarId', 'location']
};
const STATUS = { PLANNED: '예정', ONGOING: '진행중', DONE: '완료', CANCELED: '취소' };
const RESPONSES = ['참석', '불참', '미정'];
const FMT = 'yyyy-MM-dd HH:mm';

/* ===================== 설정 ===================== */

function setupMeeting() {
  assertOwner_();
  const ss = ss_();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  const stamp = Utilities.formatDate(new Date(), appTz_(), 'yyyyMMdd_HHmm');

  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 0) {
      const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
      const same = HEADERS[name].every((h, i) => head[i] === h);
      if (!same && sh.getLastRow() > 1) {
        // 예전 형식(시간 투표 버전) 데이터는 지우지 않고 이름을 바꿔 보관
        sh.setName(name + '_구버전_' + stamp);
        sh = null;
      } else if (!same) {
        sh.getRange(1, 1, 1, sh.getMaxColumns()).clearContent();
      }
    }
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange('A:Z').setNumberFormat('@'); // 날짜·시간 문자열이 자동 변환되지 않게
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });

  const rooms = ss.getSheetByName('Rooms');
  if (rooms.getLastRow() < 2) {
    rooms.getRange(2, 1, 3, 5).setValues([
      ['R1', '소회의실 A', '4', '', '3층'],
      ['R2', '중회의실 B', '8', '', '3층'],
      ['R3', '대회의실', '16', '', '5층']
    ]);
  }

  installStatusTrigger_();
  try { SpreadsheetApp.getUi().alert('회의 시트 설정 완료!\nRooms 시트에 실제 회의실 정보를 입력하세요.\n회의 상태 자동 변경 트리거(10분 간격)도 설치했습니다.'); } catch (e) {}
}

/** 10분마다 updateMeetingStatuses 실행 (중복 설치 방지) */
function installStatusTrigger_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'updateMeetingStatuses')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('updateMeetingStatuses').timeBased().everyMinutes(10).create();
}

/* ===================== 화면 진입 ===================== */

/**
 * IndexMeeting.html 생성. page: create / respond / manage / view
 * withNav=false 이면 공통 메뉴 없이 생성 (스프레드시트 대화상자용)
 */
function render_(page, p, withNav) {
  p = p || {};
  return page_('IndexMeeting', withNav ? 'meeting' : '', {
    '__PAGE__': safeToken_(page),
    '__MID__': safeToken_(p.mid),
    '__KEY__': safeToken_(p.key),
    '__TOKEN__': safeToken_(p.token)
  });
}

function showCreateDialog() {
  assertOwner_();
  SpreadsheetApp.getUi().showModalDialog(render_('create', {}, false).setWidth(640).setHeight(720), '새 회의 등록');
}

/* ===================== 시트 헬퍼 ===================== */

function sh_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error(`'${name}' 시트가 없습니다. 메뉴 > ① 초기 설정을 먼저 실행하세요.`);
  return s;
}
function rows_(name) {
  const v = sh_(name).getDataRange().getDisplayValues();
  const h = v.shift() || [];
  return v.map((r, i) => {
    const o = { _row: i + 2 };
    h.forEach((k, j) => (o[k] = r[j]));
    return o;
  }).filter(o => o[h[0]] !== '');
}
/** '='로 시작하는 값이 수식으로 실행되지 않게 막음 */
const cell_ = v => { const s = v == null ? '' : String(v); return s.charAt(0) === '=' ? "'" + s : s; };
function append_(name, obj) {
  sh_(name).appendRow(HEADERS[name].map(k => cell_(obj[k])));
}
function patch_(name, row, obj) {
  const s = sh_(name);
  HEADERS[name].forEach((k, j) => {
    if (obj[k] !== undefined) s.getRange(row, j + 1).setValue(cell_(obj[k]));
  });
}
const tz_ = () => appTz_();
const fmt_ = d => Utilities.formatDate(d, tz_(), FMT);
const parse_ = s => Utilities.parseDate(s, tz_(), FMT);
const startOf_ = m => parse_(m.date + ' ' + m.startTime);
const endOf_ = m => parse_(m.date + ' ' + m.endTime);
const newId_ = (prefix, n) => prefix + Utilities.getUuid().replace(/-/g, '').slice(0, n);
const clean_ = (v, max) => String(v == null ? '' : v).replace(/\u0000/g, '').trim().slice(0, max || 500);

// 웹앱 주소는 Main.gs의 appBase_()로 통일 (필요하면 스크립트 속성 WEB_APP_URL에 …/exec 주소 지정)
const webUrl_ = () => {
  const url = appBase_();
  if (!url) throw new Error('먼저 [배포 > 새 배포 > 웹 앱]으로 배포해 주세요.');
  return url;
};
const respondUrl_ = token => webUrl_() + '?r=' + token;
const manageUrl_ = m => webUrl_() + '?m=' + m.meetingId + '&admin=' + m.adminKey;
const minutesUrl_ = id => webUrl_() + '?page=minutes&meeting=' + encodeURIComponent(id);

function meeting_(id) {
  const m = rows_('Meetings').find(r => r.meetingId === id);
  if (!m) throw new Error('회의를 찾을 수 없습니다. 링크를 확인해 주세요.');
  return m;
}
function assertAdmin_(m, key) {
  if (!key || key !== m.adminKey) throw new Error('관리자 링크가 올바르지 않습니다.');
}
const attendees_ = id => rows_('Attendees').filter(r => r.meetingId === id);
const tasks_ = id => rows_('Tasks').filter(r => r.meetingId === id);
const room_ = roomId => rows_('Rooms').find(r => r.roomId === roomId) || null;

function info_(m) {
  const room = room_(m.roomId);
  return {
    meetingId: m.meetingId, title: m.title, date: m.date, startTime: m.startTime, endTime: m.endTime,
    roomName: room ? room.roomName : '', location: room ? room.location : '',
    status: m.status, agenda: m.agenda
  };
}

/* ===================== 캘린더 ===================== */

/** 캘린더 일정 설명: 안건, 할 일, 참석 응답, (회의록 저장 후) 요약 */
function describe_(m, people, tasks, minutes) {
  const pre = tasks.filter(t => t.phase === '사전');
  const post = tasks.filter(t => t.phase === '회의후');
  const lines = ['[통합 업무 관리] 회의 일정'];
  if (m.agenda) lines.push('', '■ 안건', m.agenda);
  if (pre.length) lines.push('', '■ 사전 할 일', ...pre.map(t => `- ${t.title} : ${t.assignee || '담당자 미정'}`));
  lines.push('', '■ 참석 응답', ...people.map(p => `- ${p.name} : ${p.response}`));
  if (minutes) lines.push('', '■ 회의록 요약', minutes);
  if (post.length) lines.push('', '■ 회의 후 할 일', ...post.map(t => `- ${t.title} : ${t.assignee || '미정'}${t.dueDate ? ' (~' + t.dueDate + ')' : ''}`));
  return lines.join('\n').slice(0, 8000);
}

/** 응답·할 일이 바뀔 때마다 캘린더 일정 설명을 최신 내용으로 갱신 */
function refreshCalendar_(m, minutes) {
  if (!m.eventId) return false;
  try {
    const ev = CalendarApp.getDefaultCalendar().getEventById(m.eventId);
    if (!ev) return false;
    ev.setDescription(describe_(m, attendees_(m.meetingId), tasks_(m.meetingId), minutes));
    return true;
  } catch (e) {
    Logger.log('캘린더 설명 갱신 실패: ' + e.message);
    return false;
  }
}

/** 같은 회의실·같은 시간대에 다른 회의가 있으면 이유 문자열, 없으면 '' */
function roomBusy_(room, date, startTime, endTime) {
  const s = parse_(date + ' ' + startTime), e = parse_(date + ' ' + endTime);
  const clash = rows_('Meetings').find(m => m.roomId === room.roomId && m.status !== STATUS.CANCELED &&
    m.date === date && startOf_(m) < e && endOf_(m) > s);
  if (clash) return `'${clash.title}' 회의(${clash.startTime}~${clash.endTime})와 시간이 겹칩니다.`;
  if (room.calendarId) {
    try {
      const cal = CalendarApp.getCalendarById(room.calendarId);
      if (cal && cal.getEvents(s, e).length) return '회의실 캘린더에 이미 일정이 있습니다.';
    } catch (err) { /* 캘린더 접근 불가 시 시트 기준으로만 확인 */ }
  }
  return '';
}

/* ===================== ① 회의 등록 + ② 초대 ===================== */

function getRoomList() {
  assertAccess_();
  return rows_('Rooms').map(r => ({
    roomId: r.roomId, roomName: r.roomName, capacity: Number(r.capacity) || 0, location: r.location
  }));
}

function createMeeting(f) {
  assertAccess_();
  f = f || {};

  // --- 입력값 검사 ---
  const title = clean_(f.title, 200);
  if (!title) throw new Error('회의명을 입력하세요.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(f.date)) throw new Error('회의 날짜를 확인하세요.');
  if (!/^\d{2}:\d{2}$/.test(f.startTime) || !/^\d{2}:\d{2}$/.test(f.endTime)) throw new Error('시작·종료 시간을 확인하세요.');
  if (f.startTime >= f.endTime) throw new Error('종료 시간은 시작 시간보다 늦어야 합니다.');
  const start = parse_(f.date + ' ' + f.startTime), end = parse_(f.date + ' ' + f.endTime);
  if (end <= new Date()) throw new Error('이미 지난 시간에는 회의를 등록할 수 없습니다.');

  const room = room_(String(f.roomId || ''));
  if (!room) throw new Error('회의실을 선택하세요.');

  const seen = new Set();
  const people = String(f.attendees || '').split('\n').map(l => l.trim()).filter(Boolean).map(l => {
    const [name, email] = l.split(/[,\t]/).map(s => clean_(s, 100));
    return { name, email: (email || '').toLowerCase() };
  });
  people.forEach(p => {
    if (!p.name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(p.email)) throw new Error(`참석자 형식을 확인하세요: "${p.name}, ${p.email}" (이름, 이메일)`);
    if (seen.has(p.email)) throw new Error('같은 이메일이 두 번 입력되었습니다: ' + p.email);
    seen.add(p.email);
  });
  if (!people.length) throw new Error('참석자를 1명 이상 입력하세요.');
  if (new Set(people.map(p => p.name)).size !== people.length) throw new Error('참석자 이름이 중복됩니다. (예: 김민수A, 김민수B)');
  const cap = Number(room.capacity) || 0;
  if (cap && people.length > cap) throw new Error(`회의실 정원(${cap}명)보다 참석자(${people.length}명)가 많습니다.`);

  const agenda = clean_(f.agenda, 3000);
  const taskTitles = String(f.tasks || '').split('\n').map(s => clean_(s, 200)).filter(Boolean).slice(0, 30);

  // --- 저장 (회의실 중복 확인 → 캘린더 → 시트) ---
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let m, attendees;
  try {
    const busy = roomBusy_(room, f.date, f.startTime, f.endTime);
    if (busy) throw new Error('회의실을 사용할 수 없습니다. ' + busy);

    const id = newId_('m', 8); // 영문으로 시작해야 시트가 숫자로 바꾸지 않음
    const now = fmt_(new Date());
    m = {
      meetingId: id, title, date: f.date, startTime: f.startTime, endTime: f.endTime, roomId: room.roomId,
      status: STATUS.PLANNED, eventId: '', organizerEmail: currentEmail_() || ownerEmail_(),
      adminKey: newId_('k', 16), agenda, minutesRequested: '', minutesSavedAt: '', createdAt: now
    };
    attendees = people.map(p => ({ meetingId: id, name: p.name, email: p.email, token: newId_('t', 16), response: '미응답', respondedAt: '' }));
    const tasks = taskTitles.map(t => ({ taskId: newId_('a', 8), meetingId: id, title: t, assignee: '', assigneeEmail: '', dueDate: f.date, phase: '사전', done: 'FALSE', createdAt: now }));

    // 캘린더 일정 생성 → 실패하면 등록하지 않음
    let ev;
    try {
      const guests = attendees.map(p => p.email);
      if (room.calendarId) guests.push(room.calendarId); // Workspace 회의실 리소스면 실제 예약됨
      ev = CalendarApp.getDefaultCalendar().createEvent(title, start, end, {
        location: [room.roomName, room.location].filter(Boolean).join(' '),
        description: describe_(m, attendees, tasks),
        guests: guests.join(','),
        sendInvites: true
      });
    } catch (err) {
      throw new Error('캘린더 일정을 만들지 못해 등록하지 않았습니다. (' + err.message + ')');
    }
    m.eventId = ev.getId();

    // 시트 기록 → 실패하면 방금 만든 캘린더 일정도 삭제
    try {
      append_('Meetings', m);
      attendees.forEach(p => append_('Attendees', p));
      tasks.forEach(t => append_('Tasks', t));
    } catch (err) {
      try { ev.deleteEvent(); } catch (ignore) {}
      throw new Error('시트 저장에 실패해 등록을 취소했습니다. (' + err.message + ')');
    }
  } finally {
    lock.releaseLock();
  }

  // --- 메일 발송 (실패해도 등록은 유지) ---
  let mailed = 0;
  if (f.sendMail) {
    const when = `${m.date} ${m.startTime} ~ ${m.endTime}`;
    attendees.forEach(p => {
      try {
        MailApp.sendEmail({
          to: p.email,
          subject: `[회의 초대] ${title} – 참석 여부와 할 일을 선택해 주세요`,
          htmlBody: `<p>${escHtml_(p.name)}님, 안녕하세요.</p>
            <p><b>${escHtml_(title)}</b> 회의에 초대되었습니다.</p>
            <table style="border-collapse:collapse;font-size:14px">
              <tr><td style="padding:4px 12px 4px 0;color:#666">일시</td><td>${escHtml_(when)}</td></tr>
              <tr><td style="padding:4px 12px 4px 0;color:#666">회의실</td><td>${escHtml_(room.roomName)} ${escHtml_(room.location || '')}</td></tr>
              ${agenda ? `<tr><td style="padding:4px 12px 4px 0;color:#666;vertical-align:top">안건</td><td style="white-space:pre-wrap">${escHtml_(agenda)}</td></tr>` : ''}
            </table>
            <p><a href="${escHtml_(respondUrl_(p.token))}" style="display:inline-block;background:#002c5f;color:#fff;padding:10px 16px;text-decoration:none">참석 여부 · 할 일 선택하기</a></p>
            <p style="color:#888;font-size:12px">이 링크는 ${escHtml_(p.name)}님 전용입니다. 다른 사람에게 전달하지 마세요.</p>`
        });
        mailed++;
      } catch (err) { Logger.log('초대 메일 실패(' + p.email + '): ' + err.message); }
    });
    if (m.organizerEmail) {
      try {
        MailApp.sendEmail({
          to: m.organizerEmail,
          subject: `[회의 관리] ${title} – 관리자 링크`,
          htmlBody: `<p>참석 현황 확인과 회의 상태 변경은 아래 관리자 링크에서 할 수 있습니다. (외부 공유 금지)</p>
            <p><a href="${escHtml_(manageUrl_(m))}">${escHtml_(manageUrl_(m))}</a></p>`
        });
      } catch (err) { Logger.log('관리자 메일 실패: ' + err.message); }
    }
  }

  return {
    manageUrl: manageUrl_(m),
    mailed,
    links: attendees.map(p => ({ name: p.name, url: respondUrl_(p.token) }))
  };
}

/* ===================== ③ 참석자 응답 ===================== */

function attendeeByToken_(token) {
  const t = safeToken_(token);
  const a = t ? rows_('Attendees').find(r => r.token === t) : null;
  if (!a) throw new Error('응답 링크가 올바르지 않습니다. 메일로 받은 링크를 다시 확인해 주세요.');
  return a;
}

const canRespond_ = m => m.status === STATUS.PLANNED || m.status === STATUS.ONGOING;

function getResponseView(token) {
  assertAccess_();
  syncStatuses_();
  const a = attendeeByToken_(token);
  const m = meeting_(a.meetingId);
  const myEmail = a.email.toLowerCase();
  return {
    meeting: info_(m),
    me: { name: a.name, response: a.response },
    canEdit: canRespond_(m),
    attendees: attendees_(m.meetingId).map(p => ({ name: p.name, response: p.response })),
    tasks: tasks_(m.meetingId).filter(t => t.phase === '사전').map(t => ({
      taskId: t.taskId, title: t.title, assignee: t.assignee,
      mine: !!t.assigneeEmail && t.assigneeEmail.toLowerCase() === myEmail
    }))
  };
}

/**
 * @param {string} token    개인 응답 토큰
 * @param {string} response 참석 / 불참 / 미정
 * @param {string[]} taskIds 내가 맡을 사전 할 일 ID 목록 (체크 해제하면 담당에서 빠짐)
 */
function submitResponse(token, response, taskIds) {
  assertAccess_();
  if (!RESPONSES.includes(response)) throw new Error('참석 여부를 선택하세요.');
  const want = new Set((Array.isArray(taskIds) ? taskIds : []).map(String));
  const skipped = [];
  let m;

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const a = attendeeByToken_(token);
    m = meeting_(a.meetingId);
    if (!canRespond_(m)) throw new Error('완료되었거나 취소된 회의는 응답을 바꿀 수 없습니다.');
    patch_('Attendees', a._row, { response, respondedAt: fmt_(new Date()) });

    const myEmail = a.email.toLowerCase();
    tasks_(m.meetingId).filter(t => t.phase === '사전').forEach(t => {
      const mine = !!t.assigneeEmail && t.assigneeEmail.toLowerCase() === myEmail;
      if (want.has(t.taskId)) {
        if (!t.assigneeEmail) patch_('Tasks', t._row, { assignee: a.name, assigneeEmail: a.email });
        else if (!mine) skipped.push(t.title); // 그 사이 다른 사람이 먼저 맡음
      } else if (mine) {
        patch_('Tasks', t._row, { assignee: '', assigneeEmail: '' });
      }
    });
  } finally {
    lock.releaseLock();
  }

  refreshCalendar_(m);
  return { view: getResponseView(token), skipped };
}

/* ===================== 관리자 화면 ===================== */

function getManageView(id, key) {
  assertAccess_();
  syncStatuses_();
  const m = meeting_(safeToken_(id));
  assertAdmin_(m, key);
  return {
    meeting: info_(m),
    attendees: attendees_(m.meetingId).map(p => ({
      name: p.name, email: p.email, response: p.response, respondedAt: p.respondedAt, url: respondUrl_(p.token)
    })),
    tasks: tasks_(m.meetingId).map(t => ({ title: t.title, assignee: t.assignee, dueDate: t.dueDate, phase: t.phase })),
    minutesUrl: m.status === STATUS.DONE ? minutesUrl_(m.meetingId) : '',
    minutesSaved: !!m.minutesSavedAt
  };
}

/** 관리자가 상태를 직접 변경 (자동 변경과 별개로 언제든 가능, 취소는 되돌릴 수 없음) */
function setMeetingStatus(id, key, status) {
  assertAccess_();
  if (!Object.keys(STATUS).some(k => STATUS[k] === status)) throw new Error('상태 값이 올바르지 않습니다.');
  let m, changed = false;

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    m = meeting_(safeToken_(id));
    assertAdmin_(m, key);
    if (m.status === STATUS.CANCELED) throw new Error('취소된 회의는 상태를 바꿀 수 없습니다. 새로 등록해 주세요.');
    if (m.status !== status) {
      patch_('Meetings', m._row, { status });
      m.status = status;
      changed = true;
    }
  } finally {
    lock.releaseLock();
  }

  if (changed && status === STATUS.CANCELED) cancelEffects_(m);
  if (changed && status === STATUS.DONE) requestMinutes_(m);
  return getManageView(id, key);
}

/** 취소: 캘린더 일정 삭제 + 참석자에게 취소 메일 */
function cancelEffects_(m) {
  if (m.eventId) {
    try {
      const ev = CalendarApp.getDefaultCalendar().getEventById(m.eventId);
      if (ev) ev.deleteEvent();
    } catch (e) { Logger.log('캘린더 삭제 실패: ' + e.message); }
  }
  attendees_(m.meetingId).forEach(p => {
    try {
      MailApp.sendEmail({
        to: p.email,
        subject: `[회의 취소] ${m.title} (${m.date} ${m.startTime})`,
        htmlBody: `<p>${escHtml_(p.name)}님, <b>${escHtml_(m.title)}</b> 회의(${escHtml_(m.date)} ${escHtml_(m.startTime)}~${escHtml_(m.endTime)})가 취소되었습니다.</p>`
      });
    } catch (e) { Logger.log('취소 메일 실패: ' + e.message); }
  });
}

/* ===================== ④ 상태 자동 변경 ===================== */

/**
 * 시간 기준으로 상태를 앞으로만 진행: 예정 → 진행중(시작 시각) → 완료(종료 시각)
 * 트리거(10분 간격)와 화면을 열 때 모두 호출되므로 트리거가 없어도 화면 기준으로 맞춰짐
 */
function syncStatuses_() {
  const now = new Date();
  const done = [];
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return 0;
  try {
    rows_('Meetings').forEach(m => {
      if (m.status !== STATUS.PLANNED && m.status !== STATUS.ONGOING) return;
      let next = m.status;
      try {
        if (now >= endOf_(m)) next = STATUS.DONE;
        else if (now >= startOf_(m)) next = STATUS.ONGOING;
      } catch (e) { return; } // 날짜 형식이 깨진 행은 건너뜀
      if (next !== m.status) {
        patch_('Meetings', m._row, { status: next });
        m.status = next;
        if (next === STATUS.DONE) done.push(m);
      }
    });
  } finally {
    lock.releaseLock();
  }
  done.forEach(requestMinutes_);
  return done.length;
}

/** 트리거 전용 (시간이 지난 회의의 상태만 바꾸므로 누가 호출해도 안전) */
function updateMeetingStatuses() {
  return syncStatuses_();
}

/** 완료된 회의의 주최자에게 '회의록 작성 요청' 메일 (회의당 1회) */
function requestMinutes_(m) {
  if (m.minutesRequested) return;
  try {
    patch_('Meetings', m._row, { minutesRequested: fmt_(new Date()) });
    m.minutesRequested = 'Y';
    const to = m.organizerEmail || ownerEmail_();
    if (!to) return;
    const url = minutesUrl_(m.meetingId);
    MailApp.sendEmail({
      to,
      subject: `[회의록 작성 요청] ${m.title} (${m.date})`,
      htmlBody: `<p><b>${escHtml_(m.title)}</b> 회의가 완료되었습니다.</p>
        <p>아래 링크에서 녹음 파일을 올리면 AI가 회의록을 요약하고, 저장하면 할 일이 담당자에게 자동 배포됩니다.</p>
        <p><a href="${escHtml_(url)}" style="display:inline-block;background:#002c5f;color:#fff;padding:10px 16px;text-decoration:none">AI 회의록 작성하기</a></p>`
    });
  } catch (e) {
    Logger.log('회의록 요청 메일 실패: ' + e.message);
  }
}

/* ===================== ⑤ 회의록 연동 (Minutes.gs / IndexMinutes.html 에서 사용) ===================== */

/** 회의록 화면의 '회의 선택' 목록: 완료된 회의만 */
function getCompletedMeetings() {
  assertAccess_();
  syncStatuses_();
  return rows_('Meetings')
    .filter(m => m.status === STATUS.DONE)
    .sort((a, b) => (b.date + b.startTime).localeCompare(a.date + a.startTime))
    .slice(0, 50)
    .map(m => ({
      meetingId: m.meetingId,
      label: `${m.date} ${m.startTime} · ${m.title}`,
      minutesSaved: !!m.minutesSavedAt
    }));
}

function getMeetingForMinutes(id) {
  assertAccess_();
  const m = meeting_(safeToken_(id));
  if (m.status !== STATUS.DONE) {
    throw new Error(`회의가 '완료' 상태가 된 뒤에 회의록을 작성할 수 있습니다. (현재: ${m.status})`);
  }
  const info = info_(m);
  info.minutesSaved = !!m.minutesSavedAt;
  info.attendees = attendees_(m.meetingId).map(p => ({ name: p.name, response: p.response }));
  info.preTasks = tasks_(m.meetingId).filter(t => t.phase === '사전').map(t => ({ title: t.title, assignee: t.assignee }));
  return info;
}
