// ============================================================
// 연구실 좌석 관리 - Google Apps Script API
// ============================================================
// [배포 방법]
// 1. sheets.new 에서 새 스프레드시트 생성
// 2. URL에서 SPREADSHEET_ID 복사 후 아래에 붙여넣기
// 3. 확장 프로그램 → Apps Script → 이 코드 붙여넣기 후 저장
// 4. initSheet() 실행 (최초 1회)
// 5. 배포 → 새 배포 → 웹 앱
//    - 실행 계정: 나
//    - 액세스: 모든 사용자
// 6. 배포 URL을 index.html의 API_URL에 붙여넣기
// ============================================================

const SPREADSHEET_ID = '1MmMYiPfmAHWmyiHlBr7d7s4gO-5LclylHnSW8uZzGdo'; // ← 여기에 새 시트 ID 입력
const SEATS_SHEET = 'seats';
const LOGS_SHEET  = 'logs';

// 연구실별 초기 자리 수 (원격PC실 제외)
const ROOM_CONFIG = {
  '프라임관 101':  6,
  '프라임관 102':  6,
  '프라임관 206': 10,
  '프라임관 207': 14,
  '프라임관 602': 10,
  '프라임관 603': 12,
  '프라임관 702': 15,
};

const REMOTE_ROOM = '프라임관 703'; // 로그 전용 공간

const SEAT_W = 124, SEAT_H = 86, GAP = 14, COLS = 4;

const SEATS_HEADERS = ['room', 'seat_no', 'label', 'user', 'pc', 'monitor', 'x', 'y'];
const LOGS_HEADERS  = ['id', 'name', 'lab', 'contact', 'pc_asset', 'start_date', 'end_date', 'memo', 'created_at'];

// ─────────────────────────────────────────
// GET
// ?action=getAll          → seats 전체 + logs 전체
// ?action=getRoom&room=.. → 특정 연구실 seats
// ?action=getLogs         → logs 전체
// ─────────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action || 'getAll';

    if (action === 'getAll') {
      return respond({ seats: readAllSeats(), logs: readAllLogs() });
    }
    if (action === 'getRoom') {
      const room = e.parameter.room;
      return respond({ seats: readAllSeats().filter(s => s.room === room) });
    }
    if (action === 'getLogs') {
      return respond({ logs: readAllLogs() });
    }

    return respond({ error: 'Unknown action' });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// ─────────────────────────────────────────
// POST
// seats 관련:
//   { action: 'updateSeat',   room, seat_no, user, pc, monitor }
//   { action: 'updateLayout', room, seats: [{seat_no, x, y}] }
//   { action: 'addPublic',    room, label, pc, monitor, x, y }
//   { action: 'deletePublic', room, seat_no }
// logs 관련:
//   { action: 'addLog',    name, lab, contact, pc_asset, start_date, end_date, memo }
//   { action: 'updateLog', id, end_date, memo }  ← 종료일/메모만 수정 가능
// ─────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    // ── seats ──
    if (action === 'updateSeat') {
      if (!body.user || !body.pc || !body.monitor)
        return respond({ error: '모든 항목을 입력해주세요.' });
      updateSeatRow(body.room, body.seat_no, { user: body.user, pc: body.pc, monitor: body.monitor });
      return respond({ success: true });
    }

    if (action === 'updateLayout') {
      body.seats.forEach(s => updateSeatRow(body.room, s.seat_no, { x: s.x, y: s.y }));
      return respond({ success: true });
    }

    if (action === 'addPublic') {
      if (!body.label || !body.pc || !body.monitor)
        return respond({ error: '모든 항목을 입력해주세요.' });
      addPublicRow(body.room, body);
      return respond({ success: true });
    }

    if (action === 'deletePublic') {
      deleteSeatRow(body.room, body.seat_no);
      return respond({ success: true });
    }

    // ── logs ──
    if (action === 'addLog') {
      if (!body.name || !body.lab || !body.contact || !body.pc_asset || !body.start_date || !body.end_date)
        return respond({ error: '모든 필수 항목을 입력해주세요.' });
      const id = addLogRow(body);
      return respond({ success: true, id });
    }

    if (action === 'updateLog') {
      if (!body.id) return respond({ error: 'id가 필요합니다.' });
      updateLogRow(body.id, { end_date: body.end_date, memo: body.memo });
      return respond({ success: true });
    }

    return respond({ error: 'Unknown action' });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// ─────────────────────────────────────────
// seats 내부 함수
// ─────────────────────────────────────────
function readAllSeats() {
  const sheet = getSheet(SEATS_SHEET);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  return values.slice(1).map(row => ({
    room:    row[0],
    seat_no: row[1],
    label:   row[2],
    user:    row[3],
    pc:      row[4],
    monitor: row[5],
    x:       row[6],
    y:       row[7],
  }));
}

function updateSeatRow(room, seatNo, fields) {
  const sheet = getSheet(SEATS_SHEET);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === room && String(values[i][1]) === String(seatNo)) {
      if (fields.label   !== undefined) sheet.getRange(i+1, 3).setValue(fields.label);
      if (fields.user    !== undefined) sheet.getRange(i+1, 4).setValue(fields.user);
      if (fields.pc      !== undefined) sheet.getRange(i+1, 5).setValue(fields.pc);
      if (fields.monitor !== undefined) sheet.getRange(i+1, 6).setValue(fields.monitor);
      if (fields.x       !== undefined) sheet.getRange(i+1, 7).setValue(fields.x);
      if (fields.y       !== undefined) sheet.getRange(i+1, 8).setValue(fields.y);
      return;
    }
  }
}

function addPublicRow(room, body) {
  const sheet = getSheet(SEATS_SHEET);
  const seatNo = 'pub_' + new Date().getTime();
  sheet.appendRow([room, seatNo, body.label, '', body.pc, body.monitor, body.x||0, body.y||0]);
}

function deleteSeatRow(room, seatNo) {
  const sheet = getSheet(SEATS_SHEET);
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][0] === room && String(values[i][1]) === String(seatNo)) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

// ─────────────────────────────────────────
// logs 내부 함수
// ─────────────────────────────────────────
function readAllLogs() {
  const sheet = getSheet(LOGS_SHEET);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  return values.slice(1).map(row => ({
    id:         row[0],
    name:       row[1],
    lab:        row[2],
    contact:    row[3],
    pc_asset:   row[4],
    start_date: formatDate(row[5]),
    end_date:   formatDate(row[6]),
    memo:       row[7],
    created_at: formatDate(row[8]),
  }));
}

function addLogRow(body) {
  const sheet = getSheet(LOGS_SHEET);
  const id = 'log_' + new Date().getTime();
  sheet.appendRow([
    id,
    body.name,
    body.lab,
    body.contact,
    body.pc_asset,
    body.start_date,
    body.end_date,
    body.memo || '',
    new Date(),
  ]);
  return id;
}

function updateLogRow(id, fields) {
  const sheet = getSheet(LOGS_SHEET);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      if (fields.end_date !== undefined) sheet.getRange(i+1, 7).setValue(fields.end_date);
      if (fields.memo     !== undefined) sheet.getRange(i+1, 8).setValue(fields.memo);
      return;
    }
  }
}

// ─────────────────────────────────────────
// 공통 유틸
// ─────────────────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`"${name}" 시트가 없습니다. initSheet()를 먼저 실행하세요.`);
  return sheet;
}

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Seoul', 'yyyy-MM-dd');
  return String(val);
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────
// [최초 1회] 초기 시트 세팅
// ─────────────────────────────────────────
function initSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // seats 시트
  let seatsSheet = ss.getSheetByName(SEATS_SHEET);
  if (seatsSheet) ss.deleteSheet(seatsSheet);
  seatsSheet = ss.insertSheet(SEATS_SHEET);
  seatsSheet.appendRow(SEATS_HEADERS);
  Object.entries(ROOM_CONFIG).forEach(([room, count]) => {
    for (let i = 1; i <= count; i++) {
      const col = (i-1) % COLS;
      const row = Math.floor((i-1) / COLS);
      seatsSheet.appendRow([room, i, '', '', '', '', col*(SEAT_W+GAP), row*(SEAT_H+GAP)]);
    }
  });
  seatsSheet.getRange(1, 1, 1, SEATS_HEADERS.length).setFontWeight('bold').setBackground('#e8eaf6');
  seatsSheet.setFrozenRows(1);

  // logs 시트
  let logsSheet = ss.getSheetByName(LOGS_SHEET);
  if (logsSheet) ss.deleteSheet(logsSheet);
  logsSheet = ss.insertSheet(LOGS_SHEET);
  logsSheet.appendRow(LOGS_HEADERS);
  logsSheet.getRange(1, 1, 1, LOGS_HEADERS.length).setFontWeight('bold').setBackground('#fce8b2');
  logsSheet.setFrozenRows(1);

  Logger.log(`초기화 완료!`);
  Logger.log(`seats: ${seatsSheet.getLastRow()-1}개 자리`);
  Logger.log(`logs: 준비됨`);
}
