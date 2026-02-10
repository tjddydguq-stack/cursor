/**
 * 키워드 자동 분류 시스템 (배포용)
 * A~E열(빅데이터)을 D열(분류) 기준으로 우측 각 섹션에 자동 미러링
 * M(모바일) 검색량 높은순 정렬
 *
 * [자동 감지 방식]
 * 2행 헤더에서 분류명 열을 찾아 해당 영역의 분류명을 자동 인식
 * 예: J2셀에 "메인"이라고 쓰면 G-K열이 "메인" 분류 영역이 됨
 */

// ===== 설정 =====
const CONFIG = {
  // 소스 데이터 범위 (A~E열 = 빅데이터)
  SOURCE: {
    HEADER_ROW: 2,       // 헤더 행 (키워드, PC, M, 카페탭, 분류)
    START_ROW: 3,        // 데이터 시작 행 (1행=날짜, 2행=헤더, 3행~=데이터)
    KEYWORD_COL: 1,      // A열: 키워드
    PC_COL: 2,           // B열: PC 검색량
    M_COL: 3,            // C열: 모바일 검색량
    CAFE_COL: 4,         // D열: 카페탭 O/X
    CATEGORY_COL: 5      // E열: 분류
  },

  // 각 분류 영역은 6열 단위 (5열 데이터 + 1열 구분자)
  COLS_PER_SECTION: 6,
  DATA_COLS: 5,

  // 분류 영역 시작 열 (G열 = 7부터 시작, F열은 구분자)
  DEST_START_COL: 7,

  // 각 섹션의 데이터 시작 행 (3행부터 데이터)
  DEST_START_ROW: 3,

  // 초기화할 최대 행 수
  MAX_ROWS: 500,

  // 스캔할 최대 열 수
  MAX_SCAN_COLS: 60,

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
    .addItem('중복 제거 (A~E열)', '중복제거')
    .addSeparator()
    .addItem('분류 목록 확인', '분류목록확인')
    .addItem('설정 확인', '설정확인')
    .addSeparator()
    .addItem('도움말', '도움말')
    .addToUi();

  ui.createMenu('댓글 도구')
    .addItem('댓글 변환 실행', '댓글변환')
    .addItem('도움말', '댓글도움말')
    .addToUi();
}

// ========================================
// 키워드 분류 기능 (A~D열 → 우측 분류 영역)
// ========================================

/**
 * 2행 헤더를 스캔하여 분류 영역을 자동 감지
 * 6열 단위로 스캔, 5번째 열(분류열)에서 분류명 인식
 * G열(7)부터 시작: G-K + L구분자, M-Q + R구분자, ...
 * 카페열: J(10), P(16), V(22), ...
 * 분류열: K(11), Q(17), W(23), ...
 */
function detectDestinations(sheet) {
  const headerRow = sheet.getRange(CONFIG.HEADER_SCAN_ROW, 1, 1, CONFIG.MAX_SCAN_COLS).getValues()[0];
  const destinations = {};

  // G열(7)부터 6열 단위로 스캔
  for (let col = CONFIG.DEST_START_COL; col < CONFIG.MAX_SCAN_COLS; col += CONFIG.COLS_PER_SECTION) {
    // 분류열은 시작열 + 4 (5번째 열): G+4=K, M+4=Q, ...
    const categoryCol = col + 4;
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
        '2행의 J, P, V열 등에 분류명을 입력해주세요.\n' +
        '(6열 단위의 4번째 열에 분류명 입력)\n\n' +
        '예: J2="메인", P2="기내용"');
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
 * 소스 데이터(A~E열) 가져오기 - 원본 행 번호 포함
 */
function getSourceData(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.SOURCE.START_ROW) return [];

  const range = sheet.getRange(
    CONFIG.SOURCE.START_ROW,
    1,
    lastRow - CONFIG.SOURCE.START_ROW + 1,
    CONFIG.DATA_COLS  // 5열 (A~E)
  );

  const values = range.getValues();
  const result = [];

  // 키워드(A)와 분류(E)가 있는 행만 필터, 원본 행 번호 저장
  values.forEach((row, index) => {
    if (row[0] && row[4]) {
      result.push({
        data: row,
        sourceRow: CONFIG.SOURCE.START_ROW + index  // 원본 행 번호
      });
    }
  });

  return result;
}

/**
 * 분류별로 데이터 그룹화 + M검색량 높은순 정렬
 */
function groupByCategory(data) {
  const grouped = {};

  data.forEach(item => {
    const row = item.data;
    const sourceRow = item.sourceRow;
    const keyword = row[0];
    const pc = row[1] || '';
    const m = row[2] || '';
    const cafe = row[3] || '';  // 카페탭 O/X (D열)
    const category = String(row[4]).trim();  // 분류 (E열)

    if (!grouped[category]) {
      grouped[category] = [];
    }

    // 출력 순서: 키워드, PC, M, 카페탭, 분류 + 원본 행 번호
    grouped[category].push({
      values: [keyword, pc, m, cafe, category],
      sourceRow: sourceRow
    });
  });

  // 카페 O 우선 + M검색량 높은순 정렬
  Object.keys(grouped).forEach(category => {
    grouped[category].sort((a, b) => {
      const cafeA = String(a.values[3]).trim().toUpperCase();
      const cafeB = String(b.values[3]).trim().toUpperCase();

      // 카페 O 우선 (O가 앞으로)
      if (cafeA === 'O' && cafeB !== 'O') return -1;
      if (cafeA !== 'O' && cafeB === 'O') return 1;

      // 같은 카페 상태 내에서 M검색량 높은순
      const mA = parseFloat(String(a.values[2]).replace(/,/g, '')) || 0;
      const mB = parseFloat(String(b.values[2]).replace(/,/g, '')) || 0;
      return mB - mA;
    });
  });

  return grouped;
}

/**
 * 각 분류 영역에 데이터 배치 - copyTo 사용으로 드롭다운 칩 색상까지 복사
 */
function distributeData(sheet, groupedData, destinations) {
  Object.keys(destinations).forEach(category => {
    const dest = destinations[category];
    const startCol = dest.startCol;

    // 기존 데이터 및 유효성 검사, 서식 모두 초기화
    const clearRange = sheet.getRange(
      CONFIG.DEST_START_ROW,
      startCol,
      CONFIG.MAX_ROWS,
      CONFIG.DATA_COLS
    );
    clearRange.clearContent();
    clearRange.clearDataValidations();
    clearRange.clearFormat();

    if (groupedData[category] && groupedData[category].length > 0) {
      // 각 행을 원본에서 직접 copyTo로 복사 (드롭다운 칩 색상 포함)
      groupedData[category].forEach((item, index) => {
        const sourceRow = item.sourceRow;
        const destRow = CONFIG.DEST_START_ROW + index;

        // 원본 행(A~E) 전체를 목적지로 복사
        const sourceRange = sheet.getRange(sourceRow, 1, 1, CONFIG.DATA_COLS);
        const destRange = sheet.getRange(destRow, startCol, 1, CONFIG.DATA_COLS);

        // copyTo: 값, 서식, 드롭다운 모두 복사
        sourceRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      });

      // 전체 영역 가운데 정렬
      const dataRange = sheet.getRange(
        CONFIG.DEST_START_ROW,
        startCol,
        groupedData[category].length,
        CONFIG.DATA_COLS
      );
      dataRange.setHorizontalAlignment('center');

      // PC, M 열에 숫자 콤마 형식 적용
      sheet.getRange(
        CONFIG.DEST_START_ROW,
        startCol + 1,
        groupedData[category].length,
        2
      ).setNumberFormat('#,##0');
    }
  });
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
  report += '\n(카페O 우선 + 검색량순 정렬, 서식/드롭다운 포함)';

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
 * 현재 A~E열의 분류 목록 확인
 */
function 분류목록확인() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const destinations = detectDestinations(sheet);
  const sourceData = getSourceData(sheet);

  const categories = {};
  sourceData.forEach(item => {
    const cat = String(item.data[4]).trim();  // E열: 분류 (인덱스 4)
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
      const endLetter = columnToLetter(dest.startCol + CONFIG.DATA_COLS - 1);
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
  info += '소스 데이터: A~E열 (키워드, PC, M, 카페탭, 분류)\n';
  info += '데이터 시작 행: ' + CONFIG.SOURCE.START_ROW + '행\n';
  info += '헤더 스캔 행: ' + CONFIG.HEADER_SCAN_ROW + '행\n';
  info += '정렬: 카페O 우선 → M(모바일) 검색량 높은순\n\n';
  info += '[헤더에서 자동 감지된 분류 영역]\n';

  if (Object.keys(destinations).length === 0) {
    info += '(감지된 분류 없음)\n';
    info += '\n2행의 J, P, V열 등에 분류명을 입력하세요.';
  } else {
    Object.keys(destinations).forEach(cat => {
      const dest = destinations[cat];
      const startLetter = columnToLetter(dest.startCol);
      const endLetter = columnToLetter(dest.startCol + CONFIG.DATA_COLS - 1);
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
'A~E열 데이터를 우측 분류 영역에 배치합니다.\n' +
'카페O 우선, 그 안에서 M검색량 높은순 정렬됩니다.\n\n' +
'[구조]\n' +
'1행: 날짜/제목\n' +
'2행: 헤더 (키워드, PC, M, 카페탭, 분류)\n' +
'3행~: 데이터\n\n' +
'A~E열: 빅데이터 (모든 키워드)\n' +
'F열: 구분자\n' +
'G열~: 분류별 자동 미러링 (6열 단위)\n\n' +
'[자동 감지 방식]\n' +
'2행의 J, P, V열 등(6열 단위의 4번째)에\n' +
'분류명을 입력하면 자동으로 인식됩니다.\n\n' +
'예: J2="메인" -> G-K열이 메인 영역';

  showMessage(help);
}

// ========================================
// 댓글 변환 기능
// ========================================

/**
 * 댓글 데이터 변환 실행
 * 1행: 고정 헤더 (유지)
 * 2행: Input 헤더 (매칭용)
 * 3행~: Input 데이터
 *
 * 실행 시: 2행 input 헤더를 1행 고정 헤더와 매칭하여
 * 3행~ 데이터를 1행 헤더 순서에 맞게 재배치
 * 결과는 2행부터 출력 (input 헤더/데이터 삭제 후)
 */
function 댓글변환() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  try {
    // 1행: 고정 헤더 읽기
    const lastCol = sheet.getLastColumn();
    const fixedHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // 2행: Input 헤더 읽기
    const inputHeaders = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

    // 3행~: Input 데이터 읽기
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) {
      showMessage('데이터가 없습니다.\n3행부터 데이터를 입력해주세요.');
      return;
    }

    const inputData = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

    // Input 헤더에서 실제 값이 있는 열만 찾기
    const activeInputCols = [];
    for (let i = 0; i < inputHeaders.length; i++) {
      const header = String(inputHeaders[i]).trim();
      if (header) {
        activeInputCols.push({ col: i, header: header });
      }
    }

    if (activeInputCols.length === 0) {
      showMessage('2행에 Input 헤더가 없습니다.\n2행에 헤더명을 입력해주세요.');
      return;
    }

    // Input 헤더 -> 고정 헤더 매핑 생성
    // 공백 제거하여 비교 (예: "댓글 1" == "댓글1")
    const headerMap = {}; // inputColIndex -> fixedColIndex
    const matchedHeaders = [];
    const unmatchedHeaders = [];

    const normalize = (str) => str.replace(/\s+/g, ''); // 공백 모두 제거

    for (const item of activeInputCols) {
      let found = false;
      const normalizedInput = normalize(item.header);

      for (let j = 0; j < fixedHeaders.length; j++) {
        const fixedHeader = String(fixedHeaders[j]).trim();
        const normalizedFixed = normalize(fixedHeader);

        if (normalizedInput === normalizedFixed) {
          headerMap[item.col] = j;
          matchedHeaders.push(item.header + ' → ' + fixedHeader);
          found = true;
          break;
        }
      }
      if (!found) {
        unmatchedHeaders.push(item.header);
      }
    }

    if (Object.keys(headerMap).length === 0) {
      showMessage('매칭되는 헤더가 없습니다.\n\n' +
        '1행 헤더와 2행 헤더의 이름이 정확히 일치해야 합니다.\n\n' +
        '[2행 Input 헤더]\n' + activeInputCols.map(x => x.header).join(', '));
      return;
    }

    // 변환된 데이터 생성
    const outputData = [];
    for (const row of inputData) {
      // 첫 번째 input 열에 데이터가 있는 행만 처리
      const firstInputCol = activeInputCols[0].col;
      if (!row[firstInputCol] && row[firstInputCol] !== 0) continue;

      const newRow = new Array(lastCol).fill('');
      for (const [inputCol, fixedCol] of Object.entries(headerMap)) {
        newRow[fixedCol] = row[inputCol];
      }
      outputData.push(newRow);
    }

    if (outputData.length === 0) {
      showMessage('변환할 데이터가 없습니다.');
      return;
    }

    // 2행부터 전부 지우기
    if (lastRow >= 2) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    }

    // 변환된 데이터 2행부터 출력
    sheet.getRange(2, 1, outputData.length, lastCol).setValues(outputData);

    // 결과 메시지
    let report = '✓ 변환 완료!\n\n';
    report += '• 변환된 행: ' + outputData.length + '개\n';
    report += '• 매칭된 헤더: ' + matchedHeaders.length + '개\n';
    report += '  (' + matchedHeaders.join(', ') + ')\n';

    if (unmatchedHeaders.length > 0) {
      report += '\n[주의] 매칭 안된 헤더:\n';
      report += unmatchedHeaders.join(', ');
    }

    showMessage(report);

  } catch (error) {
    showMessage('오류 발생: ' + error.message);
    console.error(error);
  }
}

/**
 * 댓글 도구 도움말
 */
function 댓글도움말() {
  const help =
'=== 댓글 변환 도구 ===\n\n' +
'[사용법]\n' +
'1행: 고정 헤더 (변환 목표)\n' +
'2행: Input 헤더 (입력할 데이터의 헤더)\n' +
'3행~: Input 데이터\n\n' +
'[실행 시]\n' +
'2행 헤더와 1행 헤더를 이름으로 매칭하여\n' +
'3행~ 데이터를 1행 헤더 순서에 맞게 배치합니다.\n\n' +
'[예시]\n' +
'1행: 제목 | 본문 | 댓글1 | 댓글2 | 댓글3\n' +
'2행: 댓글1 | 댓글3\n' +
'3행: 좋아요 | 최고예요\n\n' +
'→ 변환 후 2행:\n' +
'(빈칸) | (빈칸) | 좋아요 | (빈칸) | 최고예요\n\n' +
'[주의]\n' +
'• 헤더 이름이 정확히 일치해야 합니다\n' +
'• 공백, 띄어쓰기 주의';

  showMessage(help);
}
