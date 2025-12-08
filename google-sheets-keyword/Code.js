/**
 * 키워드 자동 분류 시스템
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

  // 각 분류 영역은 4열 단위 (키워드, PC, M, 분류)
  COLS_PER_SECTION: 4,

  // 분류 영역 시작 열 (E열 = 5, 하지만 보통 F열 = 6부터 시작)
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
 * 2행 헤더를 스캔하여 분류 영역을 자동 감지
 * 4열 단위로 마지막 열(분류열)에 있는 값을 분류명으로 인식
 */
function detectDestinations(sheet) {
  const headerRow = sheet.getRange(CONFIG.HEADER_SCAN_ROW, 1, 1, CONFIG.MAX_SCAN_COLS).getValues()[0];
  const destinations = {};

  // F열(6)부터 4열 단위로 스캔
  for (let col = CONFIG.DEST_START_COL; col < CONFIG.MAX_SCAN_COLS; col += CONFIG.COLS_PER_SECTION) {
    const categoryCol = col + CONFIG.COLS_PER_SECTION - 1; // 4번째 열 (분류열)
    const categoryName = String(headerRow[categoryCol - 1] || '').trim();

    // 분류명이 있고, "분류"라는 일반 텍스트가 아닌 경우
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
    // 1. 헤더에서 분류 영역 자동 감지
    const destinations = detectDestinations(sheet);

    if (Object.keys(destinations).length === 0) {
      showMessage('분류 영역을 찾을 수 없습니다.\n\n' +
        '2행의 I, N, S, X, AC열 등에 분류명을 입력해주세요.\n' +
        '(4열 단위의 마지막 열에 분류명 입력)\n\n' +
        '예: I2="질환", N2="병원&시술"');
      return;
    }

    // 2. 소스 데이터 가져오기
    const sourceData = getSourceData(sheet);

    if (sourceData.length === 0) {
      showMessage('분류할 데이터가 없습니다.');
      return;
    }

    // 3. 분류별로 그룹화 + M검색량 높은순 정렬
    const groupedData = groupByCategory(sourceData);

    // 4. 각 분류 영역에 데이터 배치
    distributeData(sheet, groupedData, destinations);

    // 5. 결과 리포트
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

  // 빈 행 필터링 (키워드와 분류가 있는 행만)
  return values.filter(row => row[0] && row[3]);
}

/**
 * 분류별로 데이터 그룹화 + M검색량 높은순 정렬
 */
function groupByCategory(data) {
  const grouped = {};

  data.forEach(row => {
    const keyword = row[0];
    const pc = row[1] || 0;
    const m = row[2] || 0;
    const category = String(row[3]).trim();

    if (!grouped[category]) {
      grouped[category] = [];
    }

    // 키워드, PC, M, 분류 모두 저장
    grouped[category].push([keyword, pc, m, category]);
  });

  // 각 분류별로 M검색량(인덱스 2) 높은순 정렬
  Object.keys(grouped).forEach(category => {
    grouped[category].sort((a, b) => {
      const mA = typeof a[2] === 'number' ? a[2] : 0;
      const mB = typeof b[2] === 'number' ? b[2] : 0;
      return mB - mA; // 내림차순
    });
  });

  return grouped;
}

/**
 * 각 분류 영역에 데이터 배치
 */
function distributeData(sheet, groupedData, destinations) {
  Object.keys(destinations).forEach(category => {
    const dest = destinations[category];
    const startCol = dest.startCol;

    // 해당 영역 초기화 (4열: 키워드, PC, M, 분류)
    sheet.getRange(
      CONFIG.DEST_START_ROW,
      startCol,
      CONFIG.MAX_ROWS,
      4
    ).clearContent();

    // 해당 분류 데이터가 있으면 입력
    if (groupedData[category] && groupedData[category].length > 0) {
      sheet.getRange(
        CONFIG.DEST_START_ROW,
        startCol,
        groupedData[category].length,
        4
      ).setValues(groupedData[category]);
    }
  });

  // 헤더에 없는 분류 확인
  const unknownCategories = Object.keys(groupedData).filter(
    cat => !destinations[cat]
  );

  if (unknownCategories.length > 0) {
    console.log('헤더에 없는 분류:', unknownCategories);
  }
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

  // 헤더에 없는 분류 경고
  const unmapped = allCategories.filter(cat => !destinations[cat]);
  if (unmapped.length > 0) {
    report += '\n\n[주의] 다음 분류는 헤더에 없어서 배치되지 않았습니다:\n';
    report += unmapped.join(', ');
    report += '\n\n해당 분류명을 2행 헤더에 추가해주세요.';
  }

  return report;
}

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
    .addItem('키워드 분류 실행', '키워드분류')
    .addItem('분류 목록 확인', '분류목록확인')
    .addItem('설정 확인', '설정확인')
    .addSeparator()
    .addItem('도움말', '도움말')
    .addToUi();
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
'[구조]\n' +
'1행: 날짜/제목\n' +
'2행: 헤더 (키워드, PC, M, 분류)\n' +
'3행~: 데이터\n\n' +
'A~D열: 빅데이터 (모든 키워드)\n' +
'우측 영역: 분류별 자동 미러링\n\n' +
'[자동 감지 방식]\n' +
'2행의 I, N, S, X, AC열(4열 단위 마지막)에\n' +
'분류명을 입력하면 자동으로 인식됩니다.\n\n' +
'예: I2="질환" -> F-I열이 질환 영역\n' +
'    N2="병원&시술" -> K-N열이 병원&시술 영역\n\n' +
'[사용법]\n' +
'1. 2행 헤더에 분류명 설정\n' +
'2. A~D열에 키워드 데이터 입력 (3행부터)\n' +
'3. 메뉴 -> 키워드 도구 -> 키워드 분류 실행\n' +
'4. M검색량 높은순으로 자동 정렬되어 배치';

  showMessage(help);
}

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
