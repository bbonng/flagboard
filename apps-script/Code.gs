/**
 * 플래그풋볼 작전 저장소 (Google Apps Script)
 *   - 공간작전판(flagboard)  : 조별 작전 보관함 1·2·3
 *   - 작전 실행 계획서(flagplan): 조별 계획서(작전 3개 + 조원) 제출  ← 수행평가
 *   - FLAG GAME(flaggame)      : 경기 점수 기록 + 교실 TV 실시간 순위표
 *   세 페이지가 같은 시트, 같은 조 비밀번호, 같은 교사 열쇠를 씁니다.
 *
 * 쓰는 법
 *  1. 구글 시트를 새로 만들고 [확장 프로그램 > Apps Script]에 이 코드를 통째로 붙여 넣고 저장
 *  2. 시트로 돌아와 새로고침 → 위 메뉴 [작전판 > 처음 설정] 실행 (권한 허용)
 *  3. [작전판 > 조 비밀번호 만들기]로 반·조 비밀번호 생성
 *  4. Apps Script 화면에서 [배포 > 새 배포 > 유형: 웹 앱]
 *       - 다음 사용자로 실행: 나
 *       - 액세스 권한이 있는 사용자: 모든 사용자
 *     → 나온 웹 앱 주소(…/exec) 하나를 두 페이지에 모두 넣으면 끝
 *
 *  코드를 고친 뒤에는 [배포 > 배포 관리 > 연필(수정) > 버전: 새 버전 > 배포]로 올려야
 *  주소가 바뀌지 않습니다. ([새 배포]를 누르면 주소가 바뀝니다)
 *
 * 시트
 *  - 조     : 반 | 조 | 비밀번호            (선생님이 직접 고쳐도 됩니다. 조를 지우면 그 조는 로그인 못 함)
 *  - 작전   : 공간작전판 보관함. 시간 | 반 | 조 | 칸 | 작전 이름 | 설명 | 작전 코드 | 판 버전 | 숨김 | 저장 번호
 *  - 계획서 : 작전 실행 계획서 제출. 시간 | 반 | 조 | 조원 | 작전명 | 계획서 코드 | 판 버전 | 숨김 | 저장 번호
 *             → 둘 다 덮어쓰지 않고 한 줄씩 쌓입니다. '숨김'에 체크하면 그 줄은 없는 것으로 칩니다.
 *  - 초안   : 작전 실행 계획서 임시저장. 반 | 조 | 시간 | 초안 코드 | 판 버전
 *             → 조마다 한 줄, 학생이 [저장]을 누를 때마다 덮어씁니다. 평가는 [계획서] 제출본으로만 합니다.
 *             (이 탭은 처음 저장할 때 저절로 생깁니다. 메뉴를 누를 필요 없음)
 *  - 게임기록 : FLAG GAME 경기 결과. 시간 | 문제 | 반 | 번호 | 이름 | 난이도 | 점수 | 랭킹 점수 | 터치다운 | 숨김 | 저장 번호
 *             → 경기를 끝까지 하면 한 줄씩 쌓입니다. 순위표는 한 사람(반·번호·이름)의 최고 랭킹 점수만 셉니다.
 *             (이 탭은 처음 기록이 들어올 때 저절로 생깁니다)
 *  - 설정   : 교사열쇠 | 제출마감(공간작전판 보관함) | 계획서마감(수행평가 제출)
 *             예: 8,9 → 8반·9반 마감 / all → 전체 마감 / 비우면 마감 없음
 *
 * 공간작전판 보관함에는 학생 이름을 저장하지 않습니다.
 * 게임기록에는 순위표를 위해 반·번호·이름이 들어갑니다. 순위표는 교사 열쇠가 있어야 볼 수 있어요.
 * 계획서에는 평가를 위해 조원 이름과 포지션 이름이 들어갑니다 (선생님 시트에만 저장).
 */

const SHEET_TEAMS = '조', SHEET_PLAYS = '작전', SHEET_PLANS = '계획서', SHEET_CONF = '설정', SHEET_DRAFTS = '초안', SHEET_SCORES = '게임기록';
const SCORE_HEAD = ['시간', '문제', '반', '번호', '이름', '난이도', '점수', '랭킹 점수', '터치다운', '숨김', '저장 번호'];
const LV_MULT = [1, 1.5, 2, 3]; // FLAG GAME 난이도별 랭킹 배율 (루키·주전·프로·올스타)
const MAX_SCORE = 45;           // 드라이브 5번 × (터치다운 6 + 첫 공격 보너스 1 + 추가 2점)
const DRAFT_HEAD = ['반', '조', '시간', '초안 코드', '판 버전'];
const SLOTS = 3;              // 공간작전판 조마다 작전 칸 수
const MAX_FAIL = 5;           // 비밀번호를 이만큼 틀리면
const LOCK_SEC = 300;         // 이 시간(초) 동안 잠김
const MAX_CODE = 30000;       // 공간작전판 작전 코드 최대 길이
const MAX_PLAN = 48000;       // 계획서 코드 최대 길이 (시트 한 칸 한도 5만 자)

/* ---------------- 시트 메뉴 ---------------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('작전판')
    .addItem('처음 설정', 'setup')
    .addItem('조 비밀번호 만들기', 'makeTeams')
    .addItem('교사 열쇠 보기', 'showKey')
    .addToUi();
}

function setup() {
  setup_();
  SpreadsheetApp.getUi().alert('준비됐어요.\n\n교사 열쇠: ' + getConf_('교사열쇠') +
    '\n\n다음은 [작전판 > 조 비밀번호 만들기]예요.');
}

function setup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_TEAMS, ['반', '조', '비밀번호']);
  ensureSheet_(ss, SHEET_PLAYS, ['시간', '반', '조', '칸', '작전 이름', '설명', '작전 코드', '판 버전', '숨김', '저장 번호']);
  ensureSheet_(ss, SHEET_PLANS, ['시간', '반', '조', '조원', '작전명', '계획서 코드', '판 버전', '숨김', '저장 번호']);
  ensureSheet_(ss, SHEET_CONF, ['항목', '값']);
  ensureSheet_(ss, SHEET_DRAFTS, DRAFT_HEAD);
  if (!getConf_('교사열쇠')) setConf_('교사열쇠', randKey_());
  if (getConf_('제출마감') === null) setConf_('제출마감', '');
  if (getConf_('계획서마감') === null) setConf_('계획서마감', '');
}

function makeTeams() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('조 비밀번호 만들기', '반 수와 조 수를 쉼표로 적어 주세요 (예: 11,6)\n이미 있는 조의 비밀번호는 바꾸지 않아요.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const m = String(r.getResponseText()).split(',').map(s => parseInt(s, 10));
  if (!(m[0] > 0 && m[1] > 0)) { ui.alert('예: 11,6 처럼 적어 주세요'); return; }
  const added = makeTeams_(m[0], m[1]);
  ui.alert(added + '개 조를 새로 만들었어요. [조] 시트에서 비밀번호를 확인하세요.\n(조가 5개인 반은 6조 줄을 지우면 됩니다)');
}

function makeTeams_(nClass, nTeam) {
  setup_();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TEAMS);
  const have = {}; teams_().forEach(t => have[t.c + '-' + t.t] = 1);
  const used = {}; teams_().forEach(t => used[t.pin] = 1);
  const rows = [];
  for (let c = 1; c <= nClass; c++) for (let t = 1; t <= nTeam; t++) {
    if (have[c + '-' + t]) continue;
    let pin; do { pin = String(1000 + Math.floor(Math.random() * 9000)); } while (used[pin]); used[pin] = 1;
    rows.push([c, t, pin]);
  }
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
  return rows.length;
}

function showKey() { SpreadsheetApp.getUi().alert('교사 열쇠: ' + getConf_('교사열쇠') + '\n\n선생님 태블릿에서 한 번만 넣으면 됩니다.\n(공간작전판 [⚙ > 교사 화면], 계획서 평가판 모두 같은 열쇠)'); }

/* ---------------- 웹 앱 ---------------- */
function doGet() { return json_({ ok: true, msg: '플래그풋볼 작전 저장소가 켜져 있어요' }); }

function doPost(e) {
  let q;
  try { q = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, err: 'bad', msg: '요청을 읽지 못했어요' }); }
  try {
    switch (q.action) {
      case 'ping': return json_(ping_());
      // 공간작전판 보관함
      case 'login': return json_(login_(q));
      case 'save': return json_(save_(q));
      case 'teacher': return json_(teacher_(q));
      case 'history': return json_(history_(q));
      // 작전 실행 계획서
      case 'plan_login': return json_(planLogin_(q));
      case 'plan_save': return json_(planSave_(q));
      case 'plan_teacher': return json_(planTeacher_(q));
      case 'plan_history': return json_(planHistory_(q));
      case 'draft_save': return json_(draftSave_(q));
      // FLAG GAME
      case 'score_submit': return json_(scoreSubmit_(q));
      case 'score_board': return json_(scoreBoard_(q));
      default: return json_({ ok: false, err: 'action', msg: '알 수 없는 요청이에요' });
    }
  } catch (err) {
    return json_({ ok: false, err: 'server', msg: '저장소에서 문제가 생겼어요: ' + err });
  }
}

function ping_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ok = !!(ss.getSheetByName(SHEET_TEAMS) && ss.getSheetByName(SHEET_PLAYS) && ss.getSheetByName(SHEET_PLANS) && ss.getSheetByName(SHEET_CONF));
  return { ok: ok, msg: ok ? '연결됐어요' : '시트가 준비되지 않았어요. [작전판 > 처음 설정]을 실행하세요', teams: ok ? teams_().length : 0, now: new Date().toISOString() };
}

/* ---------------- 공간작전판 보관함 ---------------- */
function login_(q) {
  const a = auth_(q.c, q.t, q.pin); if (!a.ok) return a;
  return { ok: true, c: +q.c, t: +q.t, slots: latestOf_(+q.c, +q.t), closed: isClosed_(+q.c, '제출마감') };
}

function save_(q) {
  const a = auth_(q.c, q.t, q.pin); if (!a.ok) return a;
  const c = +q.c, t = +q.t, slot = +q.slot;
  if (!(slot >= 1 && slot <= SLOTS)) return { ok: false, err: 'slot', msg: '작전 칸은 1~' + SLOTS + '이에요' };
  if (isClosed_(c, '제출마감')) return { ok: false, err: 'closed', msg: c + '반은 저장이 마감됐어요' };
  const code = String(q.code || '');
  if (!code || code.length > MAX_CODE) return { ok: false, err: 'code', msg: '작전 내용이 비었거나 너무 길어요' };
  const id = String(q.id || '').slice(0, 40);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLAYS);
    if (id && seenId_(sh, 10, id)) return { ok: true, dup: true, slot: slot, n: countOf_(c, t, slot) }; // 응답이 끊겨 다시 보낸 경우 한 번만 저장
    sh.appendRow([new Date(), c, t, slot, safe_(q.name, 30) || ('작전 ' + slot), safe_(q.desc, 200), code, safe_(q.ver, 10), false, id]);
  } finally { lock.releaseLock(); }
  return { ok: true, slot: slot, n: countOf_(c, t, slot) };
}

function teacher_(q) {
  const k = keyCheck_(q.key); if (!k.ok) return k;
  const teams = teams_().map(t => ({ c: t.c, t: t.t }));
  const latest = {}, count = {};
  plays_().forEach(p => { if (p.hidden) return; const key = p.c + '-' + p.t + '-' + p.slot; latest[key] = p; count[key] = (count[key] || 0) + 1; });
  const plays = Object.keys(latest).map(key => { const p = latest[key]; return { c: p.c, t: p.t, slot: p.slot, name: p.name, desc: p.desc, code: p.code, ver: p.ver, time: p.time, n: count[key] }; });
  return { ok: true, teams: teams, plays: plays, closed: String(getConf_('제출마감') || ''), now: new Date().toISOString() };
}

function history_(q) {
  const k = keyCheck_(q.key); if (!k.ok) return k;
  const c = +q.c, t = +q.t, slot = +q.slot;
  const list = plays_().filter(p => p.c === c && p.t === t && p.slot === slot)
    .map(p => ({ name: p.name, desc: p.desc, code: p.code, ver: p.ver, time: p.time, hidden: p.hidden }));
  return { ok: true, c: c, t: t, slot: slot, list: list.reverse() };
}

/* ---------------- FLAG GAME 기록 ---------------- */
function scoreSubmit_(q) {
  const quiz = +q.q, c = +q.c, n = +q.n, lv = +q.lv, score = +q.score, td = +q.td, name = safe_(q.name, 10);
  if (!(quiz >= 1 && quiz <= 999999 && quiz === Math.floor(quiz))) return { ok: false, err: 'q', msg: '문제 번호가 이상해요' };
  if (!(c >= 1 && c <= 30 && n >= 1 && n <= 60) || !name) return { ok: false, err: 'who', msg: '반·번호·이름을 확인해 주세요' };
  if (!(lv >= 1 && lv <= LV_MULT.length)) return { ok: false, err: 'lv', msg: '난이도가 이상해요' };
  if (!(score >= 0 && score <= MAX_SCORE && score === Math.floor(score)) || !(td >= 0 && td <= 5)) return { ok: false, err: 'score', msg: '점수가 이상해요' };
  const rank = Math.round(score * LV_MULT[lv - 1]), id = String(q.id || '').slice(0, 40);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_SCORES, SCORE_HEAD);
    if (!(id && seenId_(sh, 11, id))) sh.appendRow([new Date(), quiz, c, n, name, lv, score, rank, td, false, id]); // 응답이 끊겨 다시 보낸 경우 한 번만
  } finally { lock.releaseLock(); }
  const board = boardOf_(quiz), me = c + '-' + n + '-' + name;
  const place = board.findIndex(r => r.c + '-' + r.n + '-' + r.name === me) + 1;
  return { ok: true, rank: rank, place: place, players: board.length };
}

function scoreBoard_(q) {
  const k = keyCheck_(q.key); if (!k.ok) return k;
  const quiz = +q.q; if (!(quiz >= 1)) return { ok: false, err: 'q', msg: '문제 번호를 넣어 주세요' };
  const board = boardOf_(quiz);
  return { ok: true, q: quiz, rows: board.slice(0, 300), players: board.length, now: new Date().toISOString() };
}

/* 한 사람(반·번호·이름)의 최고 기록만: 랭킹 점수 → 점수 → 먼저 낸 사람 순 */
function boardOf_(quiz) {
  const best = {};
  scores_().forEach(r => { if (r.q !== quiz || r.hidden) return; const key = r.c + '-' + r.n + '-' + r.name; const b = best[key];
    if (!b || r.rank > b.rank || (r.rank === b.rank && r.score > b.score)) best[key] = Object.assign({}, r, { tries: (b ? b.tries : 0) + 1 }); else b.tries++; });
  return Object.keys(best).map(k => best[k]).sort((a, b) => b.rank - a.rank || b.score - a.score || (a.time < b.time ? -1 : 1));
}

/* ---------------- 작전 실행 계획서 ---------------- */
function planLogin_(q) {
  const a = auth_(q.c, q.t, q.pin); if (!a.ok) return a;
  const c = +q.c, t = +q.t;
  let last = null, n = 0;
  plans_().forEach(p => { if (p.c === c && p.t === t && !p.hidden) { last = p; n++; } });
  const d = draftOf_(c, t);
  return { ok: true, c: c, t: t, closed: isClosed_(c, '계획서마감'), drafts: true,
    plan: last ? { code: last.code, ver: last.ver, time: last.time, n: n } : null,
    draft: d ? { code: d.code, ver: d.ver, time: d.time } : null };
}

/* 초안 저장: 조마다 한 줄을 덮어씀 (제출과 별개, 평가에 쓰지 않음)
   base = 이 기기가 마지막으로 불러오거나 저장한 초안 시각. 다른 기기가 그 뒤에 저장했으면 force 없이는 덮어쓰지 않음 */
function draftSave_(q) {
  const a = auth_(q.c, q.t, q.pin); if (!a.ok) return a;
  const c = +q.c, t = +q.t;
  const code = String(q.code || '');
  if (!code) return { ok: false, err: 'code', msg: '저장할 내용이 비었어요' };
  if (code.length > MAX_PLAN) return { ok: false, err: 'code', msg: '계획서가 너무 길어요. 선을 조금 줄이거나 선생님께 알려 주세요' };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ensureSheet_(ss, SHEET_DRAFTS, DRAFT_HEAD);
    const v = sh.getDataRange().getValues();
    let row = 0;
    for (let i = 1; i < v.length; i++) if (+v[i][0] === c && +v[i][1] === t) { row = i + 1; break; }
    if (row && !q.force) {
      const had = v[row - 1][2] instanceof Date ? v[row - 1][2].getTime() : 0;
      const base = Date.parse(q.base || '') || 0;
      if (had && had > base + 1000) return { ok: false, err: 'newer', time: iso_(v[row - 1][2]), msg: '다른 기기에서 더 최근에 저장했어요' };
    }
    const now = new Date();
    const vals = [[c, t, now, code, safe_(q.ver, 10)]];
    if (row) sh.getRange(row, 1, 1, 5).setValues(vals); else sh.appendRow(vals[0]);
    return { ok: true, time: now.toISOString() };
  } finally { lock.releaseLock(); }
}

function planSave_(q) {
  const a = auth_(q.c, q.t, q.pin); if (!a.ok) return a;
  const c = +q.c, t = +q.t;
  if (isClosed_(c, '계획서마감')) return { ok: false, err: 'closed', msg: c + '반은 계획서 제출이 마감됐어요' };
  const code = String(q.code || '');
  if (!code) return { ok: false, err: 'code', msg: '계획서 내용이 비었어요' };
  if (code.length > MAX_PLAN) return { ok: false, err: 'code', msg: '계획서가 너무 길어요. 선을 조금 줄이거나 선생님께 알려 주세요' };
  const id = String(q.id || '').slice(0, 40);
  const members = (Array.isArray(q.members) ? q.members : []).map(s => safe_(s, 12)).filter(Boolean).join(', ');
  const names = (Array.isArray(q.names) ? q.names : []).map(s => safe_(s, 30)).join(' / ');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLANS);
    if (id && seenId_(sh, 9, id)) return { ok: true, dup: true, n: planCount_(c, t) };
    sh.appendRow([new Date(), c, t, safe_(members, 120), safe_(names, 100), code, safe_(q.ver, 10), false, id]);
  } finally { lock.releaseLock(); }
  return { ok: true, n: planCount_(c, t) };
}

function planTeacher_(q) {
  const k = keyCheck_(q.key); if (!k.ok) return k;
  const teams = teams_().map(t => ({ c: t.c, t: t.t }));
  const latest = {}, count = {};
  plans_().forEach(p => { if (p.hidden) return; const key = p.c + '-' + p.t; latest[key] = p; count[key] = (count[key] || 0) + 1; });
  const plans = Object.keys(latest).map(key => { const p = latest[key]; return { c: p.c, t: p.t, code: p.code, ver: p.ver, time: p.time, n: count[key] }; });
  return { ok: true, teams: teams, plans: plans, closed: String(getConf_('계획서마감') || ''), now: new Date().toISOString() };
}

function planHistory_(q) {
  const k = keyCheck_(q.key); if (!k.ok) return k;
  const c = +q.c, t = +q.t;
  const list = plans_().filter(p => p.c === c && p.t === t)
    .map(p => ({ code: p.code, ver: p.ver, time: p.time, hidden: p.hidden, names: p.names }));
  return { ok: true, c: c, t: t, list: list.reverse() };
}

/* ---------------- 확인 ---------------- */
function auth_(c, t, pin) {
  c = +c; t = +t;
  if (!(c > 0 && t > 0)) return { ok: false, err: 'team', msg: '반과 조를 골라 주세요' };
  const cache = CacheService.getScriptCache(), ck = 'fail_' + c + '_' + t, n = +(cache.get(ck) || 0);
  if (n >= MAX_FAIL) return { ok: false, err: 'locked', msg: '비밀번호를 ' + MAX_FAIL + '번 틀려서 5분 동안 잠겼어요' };
  const row = teams_().filter(r => r.c === c && r.t === t)[0];
  if (!row) return { ok: false, err: 'team', msg: c + '반 ' + t + '조가 없어요. 선생님께 확인하세요' };
  if (String(row.pin).trim() !== String(pin == null ? '' : pin).trim()) {
    cache.put(ck, String(n + 1), LOCK_SEC);
    const left = MAX_FAIL - n - 1;
    return { ok: false, err: 'pin', msg: left > 0 ? '비밀번호가 틀렸어요 (' + left + '번 남음)' : '비밀번호를 ' + MAX_FAIL + '번 틀려서 5분 동안 잠겼어요' };
  }
  cache.remove(ck); return { ok: true };
}

function keyCheck_(key) {
  const cache = CacheService.getScriptCache(), n = +(cache.get('fail_teacher') || 0);
  if (n >= MAX_FAIL) return { ok: false, err: 'locked', msg: '교사 열쇠를 여러 번 틀려서 5분 동안 잠겼어요' };
  if (!key || String(key).trim() !== String(getConf_('교사열쇠')).trim()) { cache.put('fail_teacher', String(n + 1), LOCK_SEC); return { ok: false, err: 'key', msg: '교사 열쇠가 맞지 않아요' }; }
  cache.remove('fail_teacher'); return { ok: true };
}

function isClosed_(c, confKey) {
  const v = String(getConf_(confKey) || '').trim().toLowerCase(); if (!v) return false;
  if (v === 'all' || v === '전체') return true;
  return v.split(/[,\s]+/).map(Number).indexOf(+c) >= 0;
}

/* ---------------- 시트 읽기 ---------------- */
function teams_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TEAMS); if (!sh) return [];
  const v = sh.getDataRange().getValues(); const out = [];
  for (let i = 1; i < v.length; i++) { const c = +v[i][0], t = +v[i][1]; if (c > 0 && t > 0) out.push({ c: c, t: t, pin: String(v[i][2]).trim() }); }
  return out;
}

function plays_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLAYS); if (!sh) return [];
  const v = sh.getDataRange().getValues(); const out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i]; const c = +r[1], t = +r[2], slot = +r[3]; if (!(c > 0 && t > 0 && slot > 0)) continue;
    out.push({ time: iso_(r[0]), c: c, t: t, slot: slot, name: String(r[4]), desc: String(r[5]), code: String(r[6]), ver: String(r[7]), hidden: isTrue_(r[8]) });
  }
  return out;
}

function scores_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SCORES); if (!sh) return [];
  const v = sh.getDataRange().getValues(); const out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i]; const quiz = +r[1], c = +r[2], n = +r[3]; if (!(quiz > 0 && c > 0 && n > 0)) continue;
    out.push({ time: iso_(r[0]), q: quiz, c: c, n: n, name: String(r[4]), lv: +r[5], score: +r[6], rank: +r[7], td: +r[8], hidden: isTrue_(r[9]) });
  }
  return out;
}

function draftOf_(c, t) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DRAFTS); if (!sh) return null;
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (+v[i][0] === c && +v[i][1] === t && String(v[i][3])) return { time: iso_(v[i][2]), code: String(v[i][3]), ver: String(v[i][4]) };
  return null;
}

function plans_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PLANS); if (!sh) return [];
  const v = sh.getDataRange().getValues(); const out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i]; const c = +r[1], t = +r[2]; if (!(c > 0 && t > 0)) continue;
    out.push({ time: iso_(r[0]), c: c, t: t, members: String(r[3]), names: String(r[4]), code: String(r[5]), ver: String(r[6]), hidden: isTrue_(r[7]) });
  }
  return out;
}

function latestOf_(c, t) {
  const best = {}, count = {};
  plays_().forEach(p => { if (p.c !== c || p.t !== t || p.hidden) return; best[p.slot] = p; count[p.slot] = (count[p.slot] || 0) + 1; });
  const out = []; for (let s = 1; s <= SLOTS; s++) { const p = best[s]; out.push(p ? { slot: s, name: p.name, desc: p.desc, code: p.code, ver: p.ver, time: p.time, n: count[s] } : { slot: s }); }
  return out;
}

function countOf_(c, t, slot) { return plays_().filter(p => p.c === c && p.t === t && p.slot === slot && !p.hidden).length; }
function planCount_(c, t) { return plans_().filter(p => p.c === c && p.t === t && !p.hidden).length; }

/* ---------------- 도우미 ---------------- */
function seenId_(sh, col, id) {
  const last = sh.getLastRow(), from = Math.max(2, last - 300); if (last < 2) return false;
  const ids = sh.getRange(from, col, last - from + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return true;
  return false;
}
function ensureSheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, head.length).setValues([head]); sh.setFrozenRows(1); }
  return sh;
}
function getConf_(k) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONF); if (!sh) return null;
  const v = sh.getDataRange().getValues(); for (let i = 1; i < v.length; i++) if (String(v[i][0]) === k) return String(v[i][1]);
  return null;
}
function setConf_(k, val) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONF); const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === k) { sh.getRange(i + 1, 2).setValue(val); return; }
  sh.appendRow([k, val]);
}
function iso_(d) { return d instanceof Date ? d.toISOString() : String(d); }
function isTrue_(x) { return x === true || String(x).toUpperCase() === 'TRUE'; }
function randKey_() { const a = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; let s = ''; for (let i = 0; i < 10; i++) s += a[Math.floor(Math.random() * a.length)]; return s; }
function safe_(s, n) { s = String(s == null ? '' : s).replace(/[\r\n\t]+/g, ' ').trim().slice(0, n); return /^[=+\-@]/.test(s) ? "'" + s : s; } // 시트 수식으로 바뀌지 않게
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
