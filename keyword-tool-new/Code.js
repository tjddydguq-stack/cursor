/**
 * 키워드 자동 분류 시스템 (배포용)
 * A~D열(빅데이터)을 D열(분류) 기준으로 우측 각 섹션에 자동 미러링
 * M(모바일) 검색량 높은순 정렬
 *
 * [자동 감지 방식]
 * 1행 헤더에서 "분류" 열을 찾아 해당 영역의 분류명을 자동 인식
 * 예: I1셀에 "질환"이라고 쓰면 F-I열이 "질환" 분류 영역이 됨
 */

// ===== 설정 =====
const CONFIG = {
  // 소스 데이터 범위 (A~D열 = 빅데이터)
  SOURCE: {
    HEADER_ROW: 2,       // 헤더 행 (키워드, PC, M, 분류)
    START_ROW: 3,        // 데이터 시작 행 (1행=날짜, 2행=헤더, 3행~=데이터)
    KEYWORD_COL: 1,      // A열: 키워드
    PC_COL: 2,           // B열: PC 검색량
    M_COL: 3,            // C열: 모바일 검색량
    CATEGORY_COL: 4      // D열: 분류
  },

  // 각 분류 영역은 5열 단위 (4열 데이터 + 1열 구분자)
  COLS_PER_SECTION: 5,
  DATA_COLS: 4,

  // 분류 영역 시작 열 (F열 = 6부터 시작, E열은 구분자)
  DEST_START_COL: 6,

  // 각 섹션의 데이터 시작 행 (3행부터 데이터)
  DEST_START_ROW: 3,

  // 초기화할 최대 행 수
  MAX_ROWS: 500,

  // 스캔할 최대 열 수
  MAX_SCAN_COLS: 50,

  // 헤더 스캔 행 (분류명이 있는 행)
  HEADER_SCAN_ROW: 2
};

/**
 * 메시지 표시
 */
function showMessage(message) {
  SpreadsheetApp.getUi().alert(message);
}

/**
 * 메뉴 추가
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('키워드 도구')
    .addItem('키워드 분류 실행 (현재 시트)', '키워드분류')
    .addItem('키워드 분류 실행 (전체 시트)', '전체시트분류')
    .addSeparator()
    .addItem('중복 제거 (A~D열)', '중복제거')
    .addSeparator()
    .addItem('분류 목록 확인', '분류목록확인')
    .addItem('설정 확인', '설정확인')
    .addSeparator()
    .addItem('도움말', '도움말')
    .addToUi();
}

// ========================================
// 키워드 분류 기능 (A~D열 → 우측 분류 영역)
// ========================================

/**
 * 2행 헤더를 스캔하여 분류 영역을 자동 감지
 */
function detectDestinations(sheet) {
  const headerRow = sheet.getRange(CONFIG.HEADER_SCAN_ROW, 1, 1, CONFIG.MAX_SCAN_COLS).getValues()[0];
  const destinations = {};

  for (let col = CONFIG.DEST_START_COL; col < CONFIG.MAX_SCAN_COLS; col += CONFIG.COLS_PER_SECTION) {
    const categoryCol = col + CONFIG.DATA_COLS - 1;
    const categoryName = String(headerRow[categoryCol - 1] || '').trim();

    if (categoryName && categoryName !== '분류') {
      destinations[categoryName] = {
        startCol: col,
        categoryCol: categoryCol
      };
    }
  }

  return destinations;
}

/**
 * 메인 함수: 키워드 분류 실행
 */
function 키워드분류() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  try {
    const destinations = detectDestinations(sheet);

    if (Object.keys(destinations).length === 0) {
      showMessage('분류 영역을 찾을 수 없습니다.\n\n' +
        '2행의 I, N, S, X, AC열 등에 분류명을 입력해주세요.\n' +
        '(4열 단위의 마지막 열에 분류명 입력)\n\n' +
        '예: I2="질환", N2="병원&시술"');
      return;
    }

    const sourceData = getSourceData(sheet);

    if (sourceData.length === 0) {
      showMessage('분류할 데이터가 없습니다.');
      return;
    }

    const groupedData = groupByCategory(sourceData);
    distributeData(sheet, groupedData, destinations);

    const report = generateReport(groupedData, destinations);
    showMessage(report);

  } catch (error) {
    showMessage('오류 발생: ' + error.message);
    console.error(error);
  }
}

/**
 * 소스 데이터(A~D열) 가져오기
 */
function getSourceData(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.SOURCE.START_ROW) return [];

  const range = sheet.getRange(
    CONFIG.SOURCE.START_ROW,
    1,
    lastRow - CONFIG.SOURCE.START_ROW + 1,
    4
  );

  const values = range.getValues();
  return values.filter(row => row[0] && row[3]);
}

/**
 * 분류별로 데이터 그룹화 + M검색량 높은순 정렬
 */
function groupByCategory(data) {
  const grouped = {};

  data.forEach(row => {
    const keyword = row[0];
    const pc = row[1] || '';
    const m = row[2] || '';
    const category = String(row[3]).trim();

    if (!grouped[category]) {
      grouped[category] = [];
    }

    grouped[category].push([keyword, pc, m, category]);
  });

  // M검색량 높은순 정렬
  Object.keys(grouped).forEach(category => {
    grouped[category].sort((a, b) => {
      const mA = parseFloat(String(a[2]).replace(/,/g, '')) || 0;
      const mB = parseFloat(String(b[2]).replace(/,/g, '')) || 0;
      return mB - mA;
    });
  });

  return grouped;
}

/**
 * 각 분류 영역에 데이터 배치
 */
function distributeData(sheet, groupedData, destinations) {
  // 원본 D열의 드롭다운 가져오기
  const sourceValidationCell = sheet.getRange(CONFIG.SOURCE.START_ROW, CONFIG.SOURCE.CATEGORY_COL);
  const sourceValidation = sourceValidationCell.getDataValidation();

  // D열의 조건부 서식 가져오기
  const conditionalFormatRules = sheet.getConditionalFormatRules();
  const dColRules = conditionalFormatRules.filter(rule => {
    const ranges = rule.getRanges();
    return ranges.some(range => range.getColumn() === CONFIG.SOURCE.CATEGORY_COL);
  });

  Object.keys(destinations).forEach(category => {
    const dest = destinations[category];
    const startCol = dest.startCol;
    const categoryCol = startCol + CONFIG.DATA_COLS - 1; // 분류 열 (I, N, S 등)

    // 기존 데이터 및 유효성 검사 모두 초기화
    const clearRange = sheet.getRange(
      CONFIG.DEST_START_ROW,
      startCol,
      CONFIG.MAX_ROWS,
      CONFIG.DATA_COLS
    );
    clearRange.clearContent();
    clearRange.clearDataValidations();

    if (groupedData[category] && groupedData[category].length > 0) {
      const dataRange = sheet.getRange(
        CONFIG.DEST_START_ROW,
        startCol,
        groupedData[category].length,
        CONFIG.DATA_COLS
      );
      dataRange.setValues(groupedData[category]);

      // PC, M 열에 숫자 콤마 형식 적용
      sheet.getRange(
        CONFIG.DEST_START_ROW,
        startCol + 1,
        groupedData[category].length,
        2
      ).setNumberFormat('#,##0');

      // 분류 열(4번째 열)에 가운데 정렬 + 드롭다운 적용
      const categoryColRange = sheet.getRange(
        CONFIG.DEST_START_ROW,
        categoryCol,
        groupedData[category].length,
        1
      );
      categoryColRange.setHorizontalAlignment('center');

      if (sourceValidation) {
        categoryColRange.setDataValidation(sourceValidation);
      }
    }
  });

  // 조건부 서식을 각 분류 열에 복사
  copyConditionalFormatting(sheet, dColRules, destinations);
}

/**
 * D열의 조건부 서식을 각 분류 열에 복사
 */
function copyConditionalFormatting(sheet, dColRules, destinations) {
  if (dColRules.length === 0) return;

  const newRules = sheet.getConditionalFormatRules();

  Object.keys(destinations).forEach(category => {
    const dest = destinations[category];
    const categoryCol = dest.startCol + CONFIG.DATA_COLS - 1;

    dColRules.forEach(rule => {
      // 새 범위 생성 (해당 분류 열 전체)
      const newRange = sheet.getRange(
        CONFIG.DEST_START_ROW,
        categoryCol,
        CONFIG.MAX_ROWS,
        1
      );

      // 기존 규칙 복사하여 새 범위에 적용
      const newRule = rule.copy().setRanges([newRange]).build();
      newRules.push(newRule);
    });
  });

  sheet.setConditionalFormatRules(newRules);
}

/**
 * 결과 리포트 생성
 */
function generateReport(groupedData, destinations) {
  let report = '분류 완료!\n\n';

  let total = 0;
  const allCategories = Object.keys(groupedData);

  allCategories.forEach(category => {
    const count = groupedData[category].length;
    total += count;
    const hasDestination = destinations[category] ? '' : ' (헤더 없음)';
    report += '- ' + category + ': ' + count + '개' + hasDestination + '\n';
  });

  report += '\n총 ' + total + '개 키워드 분류됨';

  const unmapped = allCategories.filter(cat => !destinations[cat]);
  if (unmapped.length > 0) {
    report += '\n\n[주의] 다음 분류는 헤더에 없어서 배치되지 않았습니다:\n';
    report += unmapped.join(', ');
  }

  return report;
}

/**
 * 전체 시트 키워드 분류 실행
 */
function 전체시트분류() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const ui = SpreadsheetApp.getUi();

  const sheetNames = sheets.map(s => s.getName()).join('\n- ');
  const confirm = ui.alert(
    '전체 시트 분류',
    '다음 시트들에서 키워드 분류를 실행합니다:\n\n- ' + sheetNames + '\n\n계속하시겠습니까?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  let totalReport = '=== 전체 시트 분류 결과 ===\n\n';
  let processedCount = 0;

  sheets.forEach(sheet => {
    try {
      const destinations = detectDestinations(sheet);

      if (Object.keys(destinations).length === 0) {
        totalReport += '[' + sheet.getName() + '] 분류 영역 없음 (스킵)\n';
        return;
      }

      const sourceData = getSourceData(sheet);

      if (sourceData.length === 0) {
        totalReport += '[' + sheet.getName() + '] 데이터 없음 (스킵)\n';
        return;
      }

      const groupedData = groupByCategory(sourceData);
      distributeData(sheet, groupedData, destinations);

      const count = Object.values(groupedData).reduce((sum, arr) => sum + arr.length, 0);
      totalReport += '[' + sheet.getName() + '] ' + count + '개 키워드 분류 완료\n';
      processedCount++;

    } catch (e) {
      totalReport += '[' + sheet.getName() + '] 오류: ' + e.message + '\n';
    }
  });

  totalReport += '\n총 ' + processedCount + '개 시트 처리됨';
  showMessage(totalReport);
}

// ========================================
// 중복 제거 기능
// ========================================

/**
 * A~D열 중복 제거
 */
function 중복제거() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    '중복 제거',
    'A열 키워드 기준으로 중복된 행을 삭제합니다.\n' +
    '가장 위에 있는 행을 남기고 나머지는 삭제됩니다.\n\n' +
    '계속하시겠습니까?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.SOURCE.START_ROW) {
      showMessage('데이터가 없습니다.');
      return;
    }

    const keywordRange = sheet.getRange(
      CONFIG.SOURCE.START_ROW,
      CONFIG.SOURCE.KEYWORD_COL,
      lastRow - CONFIG.SOURCE.START_ROW + 1,
      1
    );
    const keywordValues = keywordRange.getValues();

    const seen = new Set();
    const duplicateRows = [];

    for (let i = 0; i < keywordValues.length; i++) {
      const keyword = String(keywordValues[i][0]);

      if (keyword === '') continue;

      if (seen.has(keyword)) {
        duplicateRows.push(CONFIG.SOURCE.START_ROW + i);
      } else {
        seen.add(keyword);
      }
    }

    if (duplicateRows.length === 0) {
      showMessage('중복된 키워드가 없습니다.');
      return;
    }

    duplicateRows.reverse();
    for (const row of duplicateRows) {
      sheet.deleteRow(row);
    }

    showMessage('중복 제거 완료!\n\n' +
      '- 삭제된 행: ' + duplicateRows.length + '개\n' +
      '- 남은 키워드: ' + seen.size + '개');

  } catch (error) {
    showMessage('오류 발생: ' + error.message);
    console.error(error);
  }
}

// ========================================
// 유틸리티 함수
// ========================================

/**
 * 열 번호를 알파벳으로 변환
 */
function columnToLetter(column) {
  let letter = '';
  while (column > 0) {
    const temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = Math.floor((column - temp - 1) / 26);
  }
  return letter;
}

/**
 * 현재 A~D열의 분류 목록 확인
 */
function 분류목록확인() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const destinations = detectDestinations(sheet);
  const sourceData = getSourceData(sheet);

  const categories = {};
  sourceData.forEach(row => {
    const cat = String(row[3]).trim();
    categories[cat] = (categories[cat] || 0) + 1;
  });

  let info = '=== 현재 분류 목록 ===\n\n';
  info += '[데이터 분류]\n';
  Object.keys(categories).forEach(cat => {
    const hasHeader = destinations[cat] ? 'O' : 'X';
    info += '[' + hasHeader + '] ' + cat + ': ' + categories[cat] + '개\n';
  });
  info += '\n(O=헤더있음, X=헤더없음)\n\n';

  info += '[헤더에서 감지된 분류]\n';
  if (Object.keys(destinations).length === 0) {
    info += '(감지된 분류 없음)\n';
  } else {
    Object.keys(destinations).forEach(cat => {
      const dest = destinations[cat];
      const startLetter = columnToLetter(dest.startCol);
      const endLetter = columnToLetter(dest.categoryCol);
      info += '- ' + cat + ': ' + startLetter + '-' + endLetter + '열\n';
    });
  }

  showMessage(info);
}

/**
 * 현재 설정 확인
 */
function 설정확인() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const destinations = detectDestinations(sheet);

  let info = '=== 현재 설정 ===\n\n';
  info += '소스 데이터: A~D열 (빅데이터)\n';
  info += '데이터 시작 행: ' + CONFIG.SOURCE.START_ROW + '행\n';
  info += '헤더 스캔 행: ' + CONFIG.HEADER_SCAN_ROW + '행\n';
  info += '정렬: M(모바일) 검색량 높은순\n\n';
  info += '[헤더에서 자동 감지된 분류 영역]\n';

  if (Object.keys(destinations).length === 0) {
    info += '(감지된 분류 없음)\n';
    info += '\n2행의 I, N, S, X, AC열 등에 분류명을 입력하세요.';
  } else {
    Object.keys(destinations).forEach(cat => {
      const dest = destinations[cat];
      const startLetter = columnToLetter(dest.startCol);
      const endLetter = columnToLetter(dest.categoryCol);
      info += '  - ' + cat + ' -> ' + startLetter + '-' + endLetter + '열\n';
    });
  }

  showMessage(info);
}

/**
 * 도움말
 */
function 도움말() {
  const help =
'=== 키워드 분류 도구 ===\n\n' +
'[키워드 분류]\n' +
'A~D열 데이터를 우측 분류 영역에 배치합니다.\n' +
'M(모바일) 검색량 높은순으로 정렬됩니다.\n\n' +
'[구조]\n' +
'1행: 날짜/제목\n' +
'2행: 헤더 (키워드, PC, M, 분류)\n' +
'3행~: 데이터\n\n' +
'A~D열: 빅데이터 (모든 키워드)\n' +
'우측 영역: 분류별 자동 미러링\n\n' +
'[자동 감지 방식]\n' +
'2행의 I, N, S, X, AC열(4열 단위 마지막)에\n' +
'분류명을 입력하면 자동으로 인식됩니다.\n\n' +
'예: I2="질환" -> F-I열이 질환 영역';

  showMessage(help);
}
