/**
 * 홍보팀 업무 대시보드 데이터 API (Google Apps Script)
 *
 * 시트 "언론유닛 업무 공유"의 여러 탭을 하나의 JSON 으로 묶어 내보냅니다.
 * 기존 언론유닛 대시보드(amc-pr-unit)가 쓰는 웹 앱과는 **별개의 배포**입니다.
 * 그쪽은 건드리지 마세요. 이건 새 프로젝트로 만들어 따로 배포합니다.
 *
 * 만들기:
 *   1. https://script.google.com → 새 프로젝트
 *   2. 이 파일 내용을 Code.gs 에 붙여넣기
 *   3. testFeed 실행 → 권한 승인 → 로그로 탭별 건수 확인
 *   4. 배포 > 새 배포 > 웹 앱 (실행: 나 / 액세스: 링크가 있는 모든 사용자)
 *   5. /exec URL 을 index.html 의 FEED_URL 에 넣기
 *
 * 내보내는 모양:
 *   { meta: {generatedAt, tabs:[{name,kind,rows}]},
 *     work: [ {source, year, month, ...}, ... ],
 *     refs: { "탭이름": {columns:[...], rows:[[...]]} } }
 */

/** 고칠 때마다 올리세요. 대시보드 '전체 현황' 맨 아래에 찍혀서, 재배포가 먹었는지 바로 보입니다. */
const SCRIPT_VERSION = '2026-11-25c';

const SPREADSHEET_ID = '15mRKud-otGhbvY5efenonaZxjlHGKRMhDU9LmoAA2WA';

/** 연락처 열은 공개 페이지에 올리지 않습니다. 헤더 이름이 걸리면 통째로 버립니다. */
const CONTACT_HEADERS = ['휴대폰', '전화', '연락처', '이메일', 'E-mail', 'EMAIL', 'e-mail'];

/**
 * 업무 기록 탭. 열 이름 → 내보낼 필드.
 * 시트에서 열 이름이 바뀌면 여기만 고치면 됩니다(인덱스가 아니라 이름으로 찾습니다).
 */
const WORK_TABS = [
  {
    sheet: '언론 업무', source: '언론',
    map: {year:'년도', month:'월', day:'게재일', type:'구분', staff:'담당',
          title:'제목/내용', dept:'진료과', doctor:'의료진', media:'배포범위', note:'비고'}
  },
  {
    sheet: '광고/협찬', source: '광고',
    map: {status:'진행 상황', year:'연도', media:'언론사', day:'일자', amount:'금액(천원)',
          kind:'품의 구분', title:'내용', budget:'예산'}
  },
  // 도움말자료는 '준비' 칸이 비고 기념일만 있는 줄이 훨씬 많습니다(1월 4주 두통의 날 등).
  // keep 에 적은 칸 중 하나라도 차 있으면 남깁니다. 아니면 달력의 대부분이 사라집니다.
  // 월은 5주를 세로로 병합해 둬서 2주부터 빈 값으로 읽힙니다 → carry 로 위 줄 값을 잇습니다.
  {
    sheet: '2026 도움말자료', source: '도움말', fixedYear: 2026,
    map: {month:'월', week:'주', title:'도움말자료 준비', anniversary:'질병 기념일', ref:'참고 자료'},
    keep: ['title', 'anniversary', 'ref'], carry: ['month']
  },
  {
    sheet: '2023 도움말자료', source: '도움말', fixedYear: 2023,
    map: {month:'월', week:'주', title:'도움말자료 준비', anniversary:'질병 기념일', ref:'참고 자료'},
    keep: ['title', 'anniversary', 'ref'], carry: ['month']
  }
];

/**
 * 참조 리스트 탭. 구조가 제각각이고 병합 셀이 많아 열을 하나하나 매핑하지 않고
 * 헤더를 찾아 그대로 표로 내보냅니다. 연락처 열은 빠집니다.
 */
const REF_TABS = [
  // headerRow 를 주면 그 줄(0부터)을, headerRows 를 주면 여러 줄을 합쳐 헤더로 씁니다.
  // 암병원 교수는 2행이 암종(폐암/위암...), 3행이 소속/교수라 두 줄을 합쳐야 '폐암 교수' 가 됩니다.
  // 자동 탐지에 맡기면 사람 이름이 적힌 4행을 머리글로 집습니다.
  {sheet: '암병원 교수', headerRows: [1, 2]},
  {sheet: '언론사 인터뷰 후보 리스트'},
  {sheet: '명의 출연진'},
  {sheet: '의료진 홍보 관련 특이사항'}
];

/** 시트 탭 이름 끝에 공백이 붙어 있는 경우가 있어(예: 'SNS ') 양끝 공백을 무시하고 찾습니다. */
function sheetByName_(ss, name) {
  const want = String(name).trim();
  const found = ss.getSheets().filter(function (sh) { return sh.getName().trim() === want; });
  return found.length ? found[0] : null;
}

/**
 * 광고·협찬은 금액이 들어 있어 공개 피드에 싣지 않습니다.
 * doGet 에서 빼고, VIEW_PW 가 맞는 POST 에만 따로 내줍니다.
 * 그래서 비번을 모르면 이 주소를 직접 열어도 광고 행이 아예 오지 않습니다.
 */
const LOCKED_SOURCES = ['광고'];

function doGet() {
  const feed = buildFeed_();
  let hidden = 0;
  feed.work = feed.work.filter(function (r) {
    if (LOCKED_SOURCES.indexOf(r.source) < 0) return true;
    hidden++;
    return false;
  });
  feed.meta.locked = {sources: LOCKED_SOURCES, rows: hidden};
  feed.meta.tabs.forEach(function (t) { if (t.name === '광고/협찬') t.locked = true; });
  return jsonOut_(feed);
}

function buildFeed_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const meta = {generatedAt: new Date().toISOString(), script: SCRIPT_VERSION, tabs: []};
  const work = [];
  const refs = {};

  WORK_TABS.forEach(function (def) {
    const sheet = sheetByName_(ss, def.sheet);
    if (!sheet) { meta.tabs.push({name: def.sheet, kind: 'work', rows: 0, error: '탭 없음'}); return; }

    const values = sheet.getDataRange().getDisplayValues();
    const headerRow = findHeaderRow_(values, objValues_(def.map));
    if (headerRow < 0) { meta.tabs.push({name: def.sheet, kind: 'work', rows: 0, error: '헤더 못 찾음'}); return; }

    const idx = columnIndex_(values[headerRow]);
    let count = 0, lastYear = '', carried = {};

    for (let i = headerRow + 1; i < values.length; i++) {
      const row = values[i];
      const rec = {source: def.source, sheet: def.sheet};
      Object.keys(def.map).forEach(function (field) {
        const col = idx[norm_(def.map[field])];
        rec[field] = col === undefined ? '' : String(row[col] || '').trim();
      });

      if (def.fixedYear) {
        rec.year = def.fixedYear;
      } else {
        // 년도 칸이 세로 병합이면 두 번째 줄부터 빈 값으로 읽힙니다
        if (/^\d{4}$/.test(rec.year)) lastYear = rec.year;
        else if (lastYear && rec.title) rec.year = lastYear;
        if (!/^\d{4}$/.test(String(rec.year))) continue;
        rec.year = Number(rec.year);
      }

      (def.carry || []).forEach(function (f) {        // 세로 병합 칸 잇기
        if (rec[f]) carried[f] = rec[f];
        else rec[f] = carried[f] || '';
      });

      const keep = def.keep || ['title'];
      if (!keep.some(function (f) { return rec[f]; })) continue;   // 빈 줄·소계 줄 제외
      rec.month = /^\d+$/.test(rec.month) ? Number(rec.month) : rec.month;
      rec._row = i + 1;                               // 시트 행 번호(대시보드에서 날짜를 고칠 때 씁니다)
      work.push(rec);
      count++;
    }
    meta.tabs.push({name: def.sheet, kind: 'work', rows: count});
  });

  REF_TABS.forEach(function (def) {
    const name = def.sheet;
    const sheet = sheetByName_(ss, name);
    if (!sheet) { meta.tabs.push({name: name, kind: 'ref', rows: 0, error: '탭 없음'}); return; }

    const values = sheet.getDataRange().getDisplayValues();
    const headerRows = def.headerRows || [def.headerRow !== undefined ? def.headerRow : findRefHeaderRow_(values)];
    const headerRow = headerRows[headerRows.length - 1];
    if (headerRow < 0 || headerRow >= values.length) { meta.tabs.push({name: name, kind: 'ref', rows: 0, error: '헤더 못 찾음'}); return; }

    const header = combineHeader_(values, headerRows);
    const keep = [];
    header.forEach(function (h, i) {
      const t = String(h || '').trim();
      if (!t) return;
      if (CONTACT_HEADERS.some(function (c) { return t.indexOf(c) >= 0; })) return;   // 연락처 제외
      keep.push({i: i, name: t});
    });
    if (!keep.length) { meta.tabs.push({name: name, kind: 'ref', rows: 0, error: '쓸 열 없음'}); return; }

    const rows = [];
    for (let i = headerRow + 1; i < values.length; i++) {
      const row = keep.map(function (k) { return String(values[i][k.i] || '').trim(); });
      if (row.every(function (v) { return !v; })) continue;      // 빈 줄
      rows.push(row);
    }
    refs[name] = {columns: keep.map(function (k) { return k.name; }), rows: rows};
    meta.tabs.push({name: name, kind: 'ref', rows: rows.length});
  });

  meta.holidays = holidays_();
  meta.newsroom = newsroom_();
  return {meta: meta, work: work, refs: refs};
}

/* ── 공휴일 ── */

/**
 * 구글이 공개해 둔 '대한민국의 휴일' 달력을 그대로 받아옵니다.
 * 네이버는 공개 API 가 없어 긁어오면 깨지고, 공공데이터포털은 인증키가 필요합니다.
 * 이 ICS 는 키 없이 열리고 음력 명절·대체공휴일이 해마다 알아서 맞습니다.
 * 6시간 캐시하므로 피드가 느려지지 않고, 못 받아오면 공휴일만 비고 나머지는 그대로 나갑니다.
 */
const HOLIDAY_ICS = 'https://calendar.google.com/calendar/ical/ko.south_korea%23holiday%40group.v.calendar.google.com/public/basic.ics';

function holidays_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('holidays');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }

  const out = {};
  try {
    const text = UrlFetchApp.fetch(HOLIDAY_ICS, {muteHttpExceptions: true}).getContentText();
    // ICS 는 긴 줄을 다음 줄 앞 공백으로 이어 붙입니다. 먼저 펴야 제목이 잘리지 않습니다.
    const lines = text.replace(/\r\n[ \t]/g, '').split(/\r?\n/);
    let start = null, end = null, name = null, kind = null;
    lines.forEach(function (ln) {
      if (ln.indexOf('BEGIN:VEVENT') === 0) { start = end = name = kind = null; return; }
      if (ln.indexOf('DTSTART') === 0) { start = (ln.split(':')[1] || '').trim(); return; }
      if (ln.indexOf('DTEND') === 0)   { end   = (ln.split(':')[1] || '').trim(); return; }
      if (ln.indexOf('SUMMARY:') === 0) { name = ln.slice(8).trim(); return; }
      // 이 달력에는 쉬는 날(DESCRIPTION 이 '공휴일')과 그냥 기념일(식목일·어버이날 등)이 같이 들어 있습니다
      if (ln.indexOf('DESCRIPTION:') === 0) { kind = ln.slice(12).trim(); return; }
      if (ln.indexOf('END:VEVENT') === 0 && start && name && kind && kind.indexOf('공휴일') === 0) {
        const y = Number(start.slice(0, 4));
        if (y >= 2018 && y <= 2032) {
          // 연휴는 DTEND 가 '끝난 다음 날'이라 하루 빼고 채웁니다
          const from = new Date(y, Number(start.slice(4, 6)) - 1, Number(start.slice(6, 8)));
          const last = end
            ? new Date(Number(end.slice(0, 4)), Number(end.slice(4, 6)) - 1, Number(end.slice(6, 8)) - 1)
            : from;
          for (var d = new Date(from); d <= last; d.setDate(d.getDate() + 1)) {
            out[Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd')] = name;
          }
        }
      }
    });
    cache.put('holidays', JSON.stringify(out), 21600);
  } catch (err) {
    // 공휴일을 못 받아도 나머지는 그대로 내보냅니다
  }
  return out;
}

/* ── 뉴스룸 보도자료 목록 ── */

/**
 * 뉴스룸(news.amc.seoul.kr)의 보도자료 목록을 받아 날짜별로 모아 둡니다.
 * 대시보드가 우리 배포 기록의 날짜로 뉴스룸 글을 찾아 제목에 링크를 겁니다.
 *
 * 제목으로 검색하면 35%밖에 못 찾습니다(시트 제목이 내부 약칭이라서).
 * 날짜로 맞추면 88%가 붙고, 같은 날 글이 둘 이상인 경우는 드뭅니다(145일 중 5일).
 * 남의 사이트 HTML 을 읽는 것이라 뉴스룸이 화면을 바꾸면 링크만 사라집니다. 나머지는 그대로입니다.
 */
const NEWSROOM_LIST = 'https://news.amc.seoul.kr/news/con/list.do?pressChk=Y&pageIndex=';
const NEWSROOM_PAGES = 40;   // 한 쪽 10건. 40쪽이면 4년치쯤 됩니다

function newsroom_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('newsroom');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }

  const out = {};
  try {
    const reqs = [];
    for (var p = 1; p <= NEWSROOM_PAGES; p++) reqs.push({url: NEWSROOM_LIST + p, muteHttpExceptions: true});
    // 한 쪽씩 돌면 40번이라 느립니다. fetchAll 로 한꺼번에 받습니다.
    const res = UrlFetchApp.fetchAll(reqs);
    const re = /detail\.do\?cntId=(\d+)[^"]*"[^>]*>([\s\S]*?)<\/a>/g;
    res.forEach(function (r) {
      if (r.getResponseCode() !== 200) return;
      const html = r.getContentText();
      var m;
      while ((m = re.exec(html)) !== null) {
        var txt = m[2].replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
        var d = txt.match(/^(\d{4})\.(\d{2})\.(\d{2})/);
        if (!d) continue;
        var title = txt.slice(d[0].length).trim();
        var first = title.split(' ')[0];
        if (first && first.length <= 3) title = title.slice(first.length).trim();   // '의료'·'행사' 같은 분류 떼기
        var key = d[1] + '-' + d[2] + '-' + d[3];
        if (!out[key]) out[key] = [];
        if (!out[key].some(function (x) { return x.i === m[1]; })) {
          out[key].push({i: m[1], t: title.slice(0, 90)});
        }
      }
    });
    cache.put('newsroom', JSON.stringify(out), 21600);
  } catch (err) {
    // 못 받아오면 링크만 안 붙습니다
  }
  return out;
}

/* ── 대시보드에서 날짜 고치기 ── */

/**
 * 공개 페이지에서 오는 요청이라 다음을 지킵니다.
 *   - 비밀번호는 스크립트 속성(EDIT_PW)에만 둡니다. 이 파일에도, 저장소에도 적지 않습니다.
 *   - 고칠 수 있는 칸은 '언론 업무' 탭의 년도·월·게재일뿐입니다.
 *   - 저장 전에 그 행의 제목이 브라우저가 알고 있는 제목과 같은지 봅니다.
 *     시트에서 행이 밀렸으면 엉뚱한 건을 고치게 되므로 중단합니다.
 *   - 무엇을 바꿨는지 '대시보드 수정 로그' 탭에 남깁니다.
 * 비밀번호를 아는 사람은 누구나 고칠 수 있습니다. 신원 확인이 아니라 자물쇠입니다.
 */
const EDITABLE_SHEET = '언론 업무';
const LOG_SHEET = '대시보드 수정 로그';

function doPost(e) {
  var out = {ok: false};
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var props = PropertiesService.getScriptProperties();

    // 광고 행 내주기 — 보기 전용 비번(VIEW_PW)
    if (body.action === 'ads') {
      var vpw = props.getProperty('VIEW_PW');
      if (!vpw) { out.error = '스크립트에 VIEW_PW 가 없습니다'; return jsonOut_(out); }
      if (String(body.pw || '') !== vpw) { out.error = '비밀번호가 다릅니다'; return jsonOut_(out); }
      out.ok = true;
      out.rows = buildFeed_().work.filter(function (r) { return LOCKED_SOURCES.indexOf(r.source) >= 0; });
      return jsonOut_(out);
    }

    // 아래는 모두 쓰기 비번(EDIT_PW)
    var pw = props.getProperty('EDIT_PW');
    if (!pw) { out.error = '스크립트에 EDIT_PW 가 없습니다'; return jsonOut_(out); }
    if (String(body.pw || '') !== pw) { out.error = '비밀번호가 다릅니다'; return jsonOut_(out); }

    // 참조 리스트에 한 줄 더하기
    if (body.action === 'addRef') return jsonOut_(addRefRow_(body));


    if (body.action !== 'setDate') { out.error = '허용되지 않은 요청입니다'; return jsonOut_(out); }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = sheetByName_(ss, EDITABLE_SHEET);
    if (!sheet) { out.error = '탭을 찾지 못했습니다'; return jsonOut_(out); }

    var values = sheet.getDataRange().getDisplayValues();
    var def = null;
    WORK_TABS.forEach(function (d) { if (d.sheet === EDITABLE_SHEET) def = d; });
    var headerRow = findHeaderRow_(values, objValues_(def.map));
    var idx = columnIndex_(values[headerRow]);

    var row = Number(body.row);
    if (!(row > headerRow + 1 && row <= values.length)) { out.error = '행 번호가 잘못됐습니다'; return jsonOut_(out); }

    var cTitle = idx[norm_(def.map.title)], cYear = idx[norm_(def.map.year)],
        cMonth = idx[norm_(def.map.month)], cDay = idx[norm_(def.map.day)];
    if (cTitle === undefined || cYear === undefined || cMonth === undefined || cDay === undefined) {
      out.error = '열을 찾지 못했습니다'; return jsonOut_(out);
    }

    var cur = values[row - 1];
    if (norm_(cur[cTitle]) !== norm_(body.title)) {
      out.error = '시트의 그 줄이 바뀌었습니다. 새로고침한 뒤 다시 해주세요';
      return jsonOut_(out);
    }

    var before = [cur[cYear], cur[cMonth], cur[cDay]].join('-');
    sheet.getRange(row, cYear + 1).setValue(Number(body.year));
    sheet.getRange(row, cMonth + 1).setValue(Number(body.month));
    sheet.getRange(row, cDay + 1).setValue(body.day === '' || body.day === null ? '' : Number(body.day));
    var after = [body.year, body.month, body.day].join('-');

    logEdit_(ss, row, cur[cTitle], before, after);
    out.ok = true;
    out.before = before;
    out.after = after;
  } catch (err) {
    out.error = String(err);
  }
  return jsonOut_(out);
}

/**
 * 참조 리스트 탭 맨 아래에 한 줄 더합니다.
 * 열은 **이름으로** 찾습니다. 대시보드가 보낸 열 이름이 그 탭에 없으면 그 값은 버립니다.
 * REF_TABS 에 있는 탭만 허용합니다 — 아무 탭에나 쓰지 못하게.
 */
function addRefRow_(body) {
  var out = {ok: false};
  var name = String(body.sheet || '');
  var allowed = REF_TABS.some(function (d) { return d.sheet === name; });
  if (!allowed) { out.error = '허용되지 않은 탭입니다'; return out; }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = sheetByName_(ss, name);
  if (!sheet) { out.error = '탭을 찾지 못했습니다'; return out; }

  var def = null;
  REF_TABS.forEach(function (d) { if (d.sheet === name) def = d; });
  var values = sheet.getDataRange().getDisplayValues();
  var headerRows = def.headerRows || [def.headerRow !== undefined ? def.headerRow : findRefHeaderRow_(values)];
  var header = combineHeader_(values, headerRows);
  var idx = columnIndex_(header);

  var vals = body.values || {};
  var row = [];
  for (var i = 0; i < header.length; i++) row.push('');
  var filled = 0;
  Object.keys(vals).forEach(function (k) {
    var c = idx[norm_(k)];
    if (c === undefined) return;
    row[c] = String(vals[k] == null ? '' : vals[k]);
    if (row[c]) filled++;
  });
  if (!filled) { out.error = '넣을 내용이 없습니다'; return out; }

  sheet.appendRow(row);
  logEdit_(ss, sheet.getLastRow(), name + ' 새 줄', '', row.filter(String).join(' / ').slice(0, 200));
  out.ok = true;
  out.row = sheet.getLastRow();
  return out;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function logEdit_(ss, row, title, before, after) {
  var sh = sheetByName_(ss, LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET);
    sh.appendRow(['시각', '탭', '행', '제목', '이전 날짜', '바꾼 날짜']);
  }
  sh.appendRow([new Date(), EDITABLE_SHEET, row, title, before, after]);
}

/* ── 헬퍼 ── */

/**
 * 머리글 줄 여러 개를 한 줄로 합칩니다(예: '폐암' + '교수' → '폐암 교수').
 * 윗줄은 가로 병합이라 첫 칸에만 값이 있고 나머지는 빈칸이므로 앞 값을 이어 씁니다.
 * 맨 아랫줄은 실제 열 이름이라 이어 쓰지 않습니다 — 빈 칸은 빈 채로 둬야 버릴 수 있습니다.
 */
function combineHeader_(values, idxs) {
  const last = idxs[idxs.length - 1];
  const width = (values[last] || []).length;
  const out = [];
  for (var c = 0; c < width; c++) out.push([]);
  idxs.forEach(function (r) {
    var carry = '';
    for (var c = 0; c < width; c++) {
      var v = String(((values[r] || [])[c]) || '').trim();
      if (v) carry = v;
      out[c].push(r === last ? v : carry);
    }
  });
  return out.map(function (parts) {
    const seen = [];
    parts.forEach(function (v) { if (v && seen.indexOf(v) < 0) seen.push(v); });
    return seen.join(' ');
  });
}


function norm_(v) { return String(v || '').replace(/\s+/g, ''); }
function objValues_(o) { return Object.keys(o).map(function (k) { return o[k]; }); }

function columnIndex_(headerCells) {
  const idx = {};
  headerCells.forEach(function (cell, i) {
    const n = norm_(cell);
    if (n && idx[n] === undefined) idx[n] = i;       // 같은 이름이 여럿이면 먼저 나온 열
  });
  return idx;
}

/** 찾는 열 이름이 절반 이상 보이는 줄을 헤더로 봅니다. 배너·병합 줄이 위에 몇 개씩 있어서요. */
function findHeaderRow_(values, wanted) {
  const want = wanted.map(norm_);
  const limit = Math.min(values.length, 12);
  let best = -1, bestHit = 0;
  for (let i = 0; i < limit; i++) {
    const have = values[i].map(norm_);
    const hit = want.filter(function (w) { return have.indexOf(w) >= 0; }).length;
    if (hit > bestHit) { bestHit = hit; best = i; }
  }
  return bestHit >= Math.ceil(want.length / 2) ? best : -1;
}

/** 참조 탭은 찾는 이름이 없으므로, 서로 다른 값이 가장 많은 줄을 헤더로 봅니다. */
function findRefHeaderRow_(values) {
  const limit = Math.min(values.length, 12);
  let best = -1, bestCount = 0;
  for (let i = 0; i < limit; i++) {
    const cells = values[i].map(function (c) { return String(c || '').trim(); }).filter(Boolean);
    const uniq = {};
    cells.forEach(function (c) { uniq[c] = 1; });
    const n = Object.keys(uniq).length;
    if (n > bestCount) { bestCount = n; best = i; }
  }
  return bestCount >= 2 ? best : -1;
}

/** 배포 전 점검용. 편집기에서 실행하고 실행 로그를 보세요. */
function testFeed() {
  const feed = buildFeed_();
  Logger.log('업무 기록 총 %s건', feed.work.length);
  feed.meta.tabs.forEach(function (t) {
    Logger.log('  [%s] %s : %s건%s', t.kind, t.name, t.rows, t.error ? ' ← ' + t.error : '');
  });
  const bySource = {};
  feed.work.forEach(function (r) { bySource[r.source] = (bySource[r.source] || 0) + 1; });
  Logger.log('source 별: %s', JSON.stringify(bySource));
  Object.keys(feed.refs).forEach(function (k) {
    Logger.log('참조 [%s] 열: %s', k, JSON.stringify(feed.refs[k].columns));
  });
  Logger.log('업무 첫 행: %s', JSON.stringify(feed.work[0]));
  Logger.log('공휴일 %s일 (못 받아오면 0)', Object.keys(feed.meta.holidays || {}).length);
  var nr = feed.meta.newsroom || {};
  var nrN = 0; Object.keys(nr).forEach(function (k) { nrN += nr[k].length; });
  Logger.log('뉴스룸 보도자료 %s건 / %s일', nrN, Object.keys(nr).length);
  var locked = feed.work.filter(function (r) { return LOCKED_SOURCES.indexOf(r.source) >= 0; }).length;
  Logger.log('공개 피드에서 빼는 광고 행 %s건 (VIEW_PW 로만 내줍니다)', locked);
}
