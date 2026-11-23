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
    sheet: 'SNS', source: 'SNS',
    map: {year:'년도', month:'제작월', format:'형식', category:'분류', views:'조회수',
          title:'주제', planner:'기획', editor:'편집', dept:'진료과(팀)',
          doctor:'교수(출연자)', done:'완료'}
  },
  {
    sheet: '광고/협찬', source: '광고',
    map: {status:'진행 상황', year:'연도', media:'언론사', day:'일자', amount:'금액(천원)',
          kind:'품의 구분', title:'내용', budget:'예산'}
  },
  {
    sheet: '2026 도움말자료', source: '도움말', fixedYear: 2026,
    map: {month:'월', week:'주', title:'도움말자료 준비', anniversary:'질병 기념일'}
  },
  {
    sheet: '2023 도움말자료', source: '도움말', fixedYear: 2023,
    map: {month:'월', week:'주', title:'도움말자료 준비', anniversary:'질병 기념일'}
  }
];

/**
 * 참조 리스트 탭. 구조가 제각각이고 병합 셀이 많아 열을 하나하나 매핑하지 않고
 * 헤더를 찾아 그대로 표로 내보냅니다. 연락처 열은 빠집니다.
 */
const REF_TABS = [
  '암병원 교수',
  '언론사 인터뷰 후보 리스트',
  '명의 출연진',
  '의료진 홍보 관련 특이사항'
];

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify(buildFeed_()))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildFeed_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const meta = {generatedAt: new Date().toISOString(), tabs: []};
  const work = [];
  const refs = {};

  WORK_TABS.forEach(function (def) {
    const sheet = ss.getSheetByName(def.sheet);
    if (!sheet) { meta.tabs.push({name: def.sheet, kind: 'work', rows: 0, error: '탭 없음'}); return; }

    const values = sheet.getDataRange().getDisplayValues();
    const headerRow = findHeaderRow_(values, objValues_(def.map));
    if (headerRow < 0) { meta.tabs.push({name: def.sheet, kind: 'work', rows: 0, error: '헤더 못 찾음'}); return; }

    const idx = columnIndex_(values[headerRow]);
    let count = 0, lastYear = '';

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

      if (!rec.title) continue;                       // 빈 줄·소계 줄 제외
      rec.month = /^\d+$/.test(rec.month) ? Number(rec.month) : rec.month;
      work.push(rec);
      count++;
    }
    meta.tabs.push({name: def.sheet, kind: 'work', rows: count});
  });

  REF_TABS.forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { meta.tabs.push({name: name, kind: 'ref', rows: 0, error: '탭 없음'}); return; }

    const values = sheet.getDataRange().getDisplayValues();
    const headerRow = findRefHeaderRow_(values);
    if (headerRow < 0) { meta.tabs.push({name: name, kind: 'ref', rows: 0, error: '헤더 못 찾음'}); return; }

    const header = values[headerRow];
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

  return {meta: meta, work: work, refs: refs};
}

/* ── 헬퍼 ── */

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
}
