/**
 * 바이럴 GFS - 댓글 3세트 양산 시스템
 */

const CONFIG = {
  TARGET_SHEET: 'cursor',
  SET_COUNT: 3
};

const COMMENT_TYPES = ['댓글1', '댓글1-1', '댓글2-1', '댓글3-1', '댓글4-1', '댓글2-2', '댓글2-3'];

const TEST_KEYWORDS = [
  '스위치온 1일차 식단',
  '위고비',
  '다이어트 저녁식단',
  '린다이어트',
  '다이어트 보조제순위',
  '떡볶이 다이어트',
  '다이어트유산균',
  '파비플로라 효과',
  '서플리케이 틴시아',
  '나이트번 프로 효과'
];

/** 메뉴 추가 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 댓글 양산')
    .addItem('★ cursor 시트 3세트 생성', 'buildCursorSheet')
    .addToUi();
}

/** ★ 메인: cursor 시트에 3세트 구조 생성 */
function buildCursorSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.TARGET_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.TARGET_SHEET);
  }

  // 헤더
  const headers = ['No', '키워드', '세트', ...COMMENT_TYPES, 'Type'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#FFD700')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 데이터 생성
  const data = [];
  let no = 1;

  TEST_KEYWORDS.forEach(keyword => {
    for (let set = 1; set <= CONFIG.SET_COUNT; set++) {
      const row = [no++, keyword, set, ...Array(COMMENT_TYPES.length).fill(''), ''];
      data.push(row);
    }
  });

  // 기존 데이터 클리어 후 쓰기
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clear();
  }
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

  // 세트별 색상
  const colors = { 1: '#FFFFFF', 2: '#E8F4FD', 3: '#FFF3E0' };
  data.forEach((row, i) => {
    sheet.getRange(i + 2, 1, 1, row.length).setBackground(colors[row[2]]);
  });

  // 열 너비 & 틀 고정
  sheet.setColumnWidths(1, 1, 50);
  sheet.setColumnWidths(2, 1, 180);
  sheet.setColumnWidths(3, 1, 50);
  sheet.setColumnWidths(4, COMMENT_TYPES.length + 1, 250);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);

  SpreadsheetApp.getUi().alert(`✅ 완료!\n${TEST_KEYWORDS.length}개 키워드 × ${CONFIG.SET_COUNT}세트 = ${data.length}행`);
}
