const BASE_HEADERS = [
  '제출일시',
  '이름',
  '성별',
  '연락처',
  '이메일',
  '학교',
  '학년',
  '전공계열',
  '학과명',
  '자차여부',
  '교구수령주소',
  '교육대상경험',
  '메이커교구경험',
  '지원동기'
];

const WEEKDAYS = ['일', '월', '화', '수', '목', '금', '토'];
const HOLIDAYS_2026 = {
  '2026-07-17': '제헌절',
  '2026-08-15': '광복절',
  '2026-08-17': '광복절 대체공휴일',
  '2026-09-24': '추석 연휴',
  '2026-09-25': '추석',
  '2026-09-26': '추석 연휴',
  '2026-10-03': '개천절',
  '2026-10-05': '개천절 대체공휴일',
  '2026-10-09': '한글날'
};

function setupSheetHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = buildHeaders_();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet, headers);
}

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = JSON.parse(e.parameter.payload || '{}');
  const headers = buildHeaders_();

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sheet, headers);

  const row = headers.map(header => data[header] || '');
  sheet.appendRow(row);
  sheet.getRange(sheet.getLastRow(), 1, 1, headers.length).setWrap(true);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildHeaders_() {
  const headers = [...BASE_HEADERS];
  for (let month = 7; month <= 11; month++) {
    const lastDay = new Date(2026, month, 0).getDate();
    for (let day = 1; day <= lastDay; day++) {
      const date = new Date(2026, month - 1, day);
      const iso = Utilities.formatDate(date, 'Asia/Seoul', 'yyyy-MM-dd');
      const holidayName = HOLIDAYS_2026[iso];
      headers.push(`${month}. ${day}. (${WEEKDAYS[date.getDay()]})${holidayName ? `\n${holidayName}` : ''}`);
    }
  }
  return headers;
}

function styleHeader_(sheet, headers) {
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBackground('#f8fafc');

  headers.forEach((header, index) => {
    const date = parseDateHeader_(header);
    if (!date) return;

    const cell = sheet.getRange(1, index + 1);
    const day = date.getDay();
    const iso = Utilities.formatDate(date, 'Asia/Seoul', 'yyyy-MM-dd');

    if (day === 6) cell.setFontColor('#2563eb');
    if (day === 0 || HOLIDAYS_2026[iso]) cell.setFontColor('#dc2626');
  });

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, Math.min(headers.length, sheet.getMaxColumns()));
}

function parseDateHeader_(header) {
  const match = String(header).match(/^(\d{1,2})\.\s*(\d{1,2})\./);
  if (!match) return null;
  return new Date(2026, Number(match[1]) - 1, Number(match[2]));
}
