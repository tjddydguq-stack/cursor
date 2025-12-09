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
  // 네이버 검색광고 API 설정
  NAVER_API: {
    ACCESS_LICENSE: '01000000007a25c6f02f5f40ab2252dd3712bcebe04647e919809a2348f9fd4d10e7feb348',
    SECRET_KEY: 'AQAAAAB6JcbwL19AqyJS3TcSvOvgtdSrGzqG+zlrq0GVt6J9Sw==',
    CUSTOMER_ID: '3526315',
    BASE_URL: 'https://api.naver.com'
  },

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
  // 데이터: 키워드, PC, M, 분류 (4열) + 구분자 (1열)
  COLS_PER_SECTION: 5,
  DATA_COLS: 4,  // 실제 데이터 열 수

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
 * 2행 헤더를 스캔하여 분류 영역을 자동 감지
 * 5열 단위 (4열 데이터 + 1열 구분자)로 스캔
 * 4번째 열(분류열)에 있는 값을 분류명으로 인식
 */
function detectDestinations(sheet) {
  const headerRow = sheet.getRange(CONFIG.HEADER_SCAN_ROW, 1, 1, CONFIG.MAX_SCAN_COLS).getValues()[0];
  const destinations = {};

  // F열(6)부터 5열 단위로 스캔 (F-I + J구분자, K-N + O구분자, ...)
  for (let col = CONFIG.DEST_START_COL; col < CONFIG.MAX_SCAN_COLS; col += CONFIG.COLS_PER_SECTION) {
    const categoryCol = col + CONFIG.DATA_COLS - 1; // 4번째 열 (분류열): F+3=I, K+3=N, ...
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
      // 문자열/숫자 모두 처리 (쉼표 제거 후 숫자 변환)
      const mA = parseFloat(String(a[2]).replace(/,/g, '')) || 0;
      const mB = parseFloat(String(b[2]).replace(/,/g, '')) || 0;
      return mB - mA; // 내림차순 (높은순)
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

    // 해당 영역 초기화 (DATA_COLS열: 키워드, PC, M, 분류)
    sheet.getRange(
      CONFIG.DEST_START_ROW,
      startCol,
      CONFIG.MAX_ROWS,
      CONFIG.DATA_COLS
    ).clearContent();

    // 해당 분류 데이터가 있으면 입력
    if (groupedData[category] && groupedData[category].length > 0) {
      const dataRange = sheet.getRange(
        CONFIG.DEST_START_ROW,
        startCol,
        groupedData[category].length,
        CONFIG.DATA_COLS
      );
      dataRange.setValues(groupedData[category]);

      // PC, M 열에 숫자 콤마 형식 적용 (2번째, 3번째 열)
      sheet.getRange(
        CONFIG.DEST_START_ROW,
        startCol + 1,  // PC 열
        groupedData[category].length,
        2  // PC, M 두 열
      ).setNumberFormat('#,##0');
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
 * A~D열 중복 제거
 * A열 키워드 기준으로 중복 행 삭제 (위에 있는 행 유지)
 */
function 중복제거() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // 확인 메시지
  const confirm = ui.alert(
    '중복 제거',
    'A열 키워드 기준으로 중복된 행을 삭제합니다.\n' +
    '(띄어쓰기까지 완전히 같은 키워드만 중복 처리)\n' +
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

    // A열 키워드 가져오기
    const keywordRange = sheet.getRange(
      CONFIG.SOURCE.START_ROW,
      CONFIG.SOURCE.KEYWORD_COL,
      lastRow - CONFIG.SOURCE.START_ROW + 1,
      1
    );
    const keywordValues = keywordRange.getValues();

    // 중복 체크 - 이미 나온 키워드 추적
    const seen = new Set();
    const duplicateRows = [];  // 삭제할 행 번호들

    for (let i = 0; i < keywordValues.length; i++) {
      const keyword = String(keywordValues[i][0]);  // 띄어쓰기 유지

      if (keyword === '') continue;  // 빈 셀 무시

      if (seen.has(keyword)) {
        // 중복 발견 - 삭제 대상
        duplicateRows.push(CONFIG.SOURCE.START_ROW + i);
      } else {
        // 처음 나온 키워드 - 유지
        seen.add(keyword);
      }
    }

    if (duplicateRows.length === 0) {
      showMessage('중복된 키워드가 없습니다.');
      return;
    }

    // 아래에서 위로 삭제 (행 번호 유지를 위해)
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
    .addItem('📥 데이터 변환 (입력→출력)', '데이터변환')
    .addSeparator()
    .addItem('키워드 분류 실행 (현재 시트)', '키워드분류')
    .addItem('키워드 분류 실행 (전체 시트)', '전체시트분류')
    .addSeparator()
    .addItem('중복 제거 (A~D열)', '중복제거')
    .addSeparator()
    .addItem('검색량 조회 (현재 시트)', '검색량조회')
    .addItem('검색량 조회 (전체 시트)', '전체시트검색량조회')
    .addItem('검색량 조회 (선택 키워드)', '선택키워드검색량조회')
    .addSeparator()
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

// ========================================
// 네이버 검색광고 API 연동
// ========================================

/**
 * 네이버 API 서명 생성
 * 참고: https://github.com/naver/searchad-apidoc
 * message = "{timestamp}.{method}.{uri}"
 * HMAC-SHA256(message, secret_key) -> base64
 */
function generateSignature(timestamp, method, path) {
  const message = timestamp + '.' + method + '.' + path;

  // Google Apps Script에서 HMAC-SHA256 생성
  // message와 secret_key 모두 UTF-8 바이트로 변환
  const messageBytes = Utilities.newBlob(message).getBytes();
  const keyBytes = Utilities.newBlob(CONFIG.NAVER_API.SECRET_KEY).getBytes();

  const signature = Utilities.computeHmacSha256Signature(messageBytes, keyBytes);
  return Utilities.base64Encode(signature);
}

/**
 * 네이버 API 호출
 */
function callNaverApi(method, path, payload) {
  const timestamp = String(new Date().getTime());
  const signature = generateSignature(timestamp, method, path);

  const options = {
    method: method,
    headers: {
      'X-Timestamp': timestamp,
      'X-API-KEY': CONFIG.NAVER_API.ACCESS_LICENSE,
      'X-Customer': CONFIG.NAVER_API.CUSTOMER_ID,
      'X-Signature': signature,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (payload) {
    options.payload = JSON.stringify(payload);
  }

  const url = CONFIG.NAVER_API.BASE_URL + path;
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    throw new Error('API 오류 (' + responseCode + '): ' + responseText);
  }

  return JSON.parse(responseText);
}

/**
 * 키워드 검색량 조회 (네이버 API)
 */
function getKeywordStats(keywords) {
  if (!keywords || keywords.length === 0) return {};

  // 최대 100개씩 처리
  const results = {};
  const chunkSize = 100;

  for (let i = 0; i < keywords.length; i += chunkSize) {
    const chunk = keywords.slice(i, i + chunkSize);
    const payload = {
      hintKeywords: chunk,
      showDetail: '1'
    };

    try {
      const response = callNaverApi('POST', '/keywordstool', payload);

      if (response.keywordList) {
        response.keywordList.forEach(item => {
          results[item.relKeyword] = {
            pc: item.monthlyPcQcCnt === '< 10' ? 0 : parseInt(item.monthlyPcQcCnt) || 0,
            mobile: item.monthlyMobileQcCnt === '< 10' ? 0 : parseInt(item.monthlyMobileQcCnt) || 0
          };
        });
      }
    } catch (e) {
      console.error('API 호출 실패:', e.message);
    }

    // API 속도 제한 방지
    if (i + chunkSize < keywords.length) {
      Utilities.sleep(500);
    }
  }

  return results;
}

/**
 * 키워드 검색량 조회 (진행 상황 표시 버전)
 */
function getKeywordStatsWithProgress(keywords, ss) {
  if (!keywords || keywords.length === 0) return {};

  const results = {};
  const chunkSize = 100;
  const totalChunks = Math.ceil(keywords.length / chunkSize);

  for (let i = 0; i < keywords.length; i += chunkSize) {
    const chunkNum = Math.floor(i / chunkSize) + 1;
    const chunk = keywords.slice(i, i + chunkSize);

    // 진행 상황 토스트 표시
    ss.toast(
      'API 호출 중... (' + chunkNum + '/' + totalChunks + ' 배치)\n' +
      '처리: ' + Math.min(i + chunkSize, keywords.length) + '/' + keywords.length + '개',
      '네이버 API',
      -1
    );

    const payload = {
      hintKeywords: chunk,
      showDetail: '1'
    };

    try {
      const response = callNaverApi('POST', '/keywordstool', payload);

      if (response.keywordList) {
        response.keywordList.forEach(item => {
          results[item.relKeyword] = {
            pc: item.monthlyPcQcCnt === '< 10' ? 0 : parseInt(item.monthlyPcQcCnt) || 0,
            mobile: item.monthlyMobileQcCnt === '< 10' ? 0 : parseInt(item.monthlyMobileQcCnt) || 0
          };
        });
      }
    } catch (e) {
      console.error('API 호출 실패:', e.message);
      ss.toast('API 오류: ' + e.message, '오류', 3);
    }

    // API 속도 제한 방지
    if (i + chunkSize < keywords.length) {
      Utilities.sleep(500);
    }
  }

  return results;
}

/**
 * A열 키워드의 검색량 자동 조회 (B, C열에 입력)
 */
function 검색량조회() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // 확인 메시지
  const confirm = ui.alert(
    '검색량 조회',
    'A열의 키워드 검색량을 네이버 API로 조회합니다.\n' +
    'B열(PC)과 C열(M)에 검색량이 입력됩니다.\n\n' +
    '계속하시겠습니까?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.SOURCE.START_ROW) {
      showMessage('조회할 키워드가 없습니다.');
      return;
    }

    // A열에서 키워드 가져오기
    const keywordRange = sheet.getRange(
      CONFIG.SOURCE.START_ROW,
      CONFIG.SOURCE.KEYWORD_COL,
      lastRow - CONFIG.SOURCE.START_ROW + 1,
      1
    );
    const keywordValues = keywordRange.getValues();
    const keywords = keywordValues.map(row => String(row[0]).trim()).filter(k => k);

    if (keywords.length === 0) {
      showMessage('조회할 키워드가 없습니다.');
      return;
    }

    // 토스트로 시작 알림
    ss.toast('검색량 조회 시작... (' + keywords.length + '개 키워드)', '진행 중', -1);

    // 네이버 API로 검색량 조회 (진행 상황 표시 버전)
    const stats = getKeywordStatsWithProgress(keywords, ss);

    // B, C열에 검색량 입력 (실시간 업데이트)
    ss.toast('검색량을 시트에 입력 중...', '진행 중', -1);
    let updatedCount = 0;
    for (let i = 0; i < keywordValues.length; i++) {
      const keyword = String(keywordValues[i][0]).trim();
      if (keyword && stats[keyword]) {
        const row = CONFIG.SOURCE.START_ROW + i;
        sheet.getRange(row, CONFIG.SOURCE.PC_COL).setValue(stats[keyword].pc);
        sheet.getRange(row, CONFIG.SOURCE.M_COL).setValue(stats[keyword].mobile);
        updatedCount++;

        // 10개마다 화면 갱신
        if (updatedCount % 10 === 0) {
          SpreadsheetApp.flush();
          ss.toast(updatedCount + '개 입력 완료...', '진행 중', -1);
        }
      }
    }
    SpreadsheetApp.flush();

    showMessage('검색량 조회 완료!\n\n' +
      '- 조회 키워드: ' + keywords.length + '개\n' +
      '- 업데이트: ' + updatedCount + '개');

  } catch (error) {
    showMessage('오류 발생: ' + error.message);
    console.error(error);
  }
}

/**
 * 선택한 셀의 키워드만 검색량 조회
 */
function 선택키워드검색량조회() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const selection = sheet.getActiveRange();

  if (!selection) {
    showMessage('키워드를 선택해주세요.');
    return;
  }

  try {
    const values = selection.getValues();
    const keywords = [];

    values.forEach(row => {
      row.forEach(cell => {
        const keyword = String(cell).trim();
        if (keyword) keywords.push(keyword);
      });
    });

    if (keywords.length === 0) {
      showMessage('선택한 영역에 키워드가 없습니다.');
      return;
    }

    // 네이버 API로 검색량 조회
    const stats = getKeywordStats(keywords);

    // 결과 표시
    let result = '=== 검색량 조회 결과 ===\n\n';
    keywords.forEach(keyword => {
      if (stats[keyword]) {
        result += keyword + '\n';
        result += '  PC: ' + stats[keyword].pc.toLocaleString() + '\n';
        result += '  M: ' + stats[keyword].mobile.toLocaleString() + '\n\n';
      } else {
        result += keyword + ': 데이터 없음\n\n';
      }
    });

    showMessage(result);

  } catch (error) {
    showMessage('오류 발생: ' + error.message);
    console.error(error);
  }
}

// ========================================
// 전체 시트 일괄 처리
// ========================================

/**
 * 전체 시트 키워드 분류 실행
 */
function 전체시트분류() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const ui = SpreadsheetApp.getUi();

  // 시트 목록 확인
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

      // 분류 영역이 없으면 스킵
      if (Object.keys(destinations).length === 0) {
        totalReport += '[' + sheet.getName() + '] 분류 영역 없음 (스킵)\n';
        return;
      }

      const sourceData = getSourceDataFromSheet(sheet);

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

/**
 * 전체 시트 검색량 조회
 */
function 전체시트검색량조회() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const ui = SpreadsheetApp.getUi();

  // 시트 목록 확인
  const sheetNames = sheets.map(s => s.getName()).join('\n- ');
  const confirm = ui.alert(
    '전체 시트 검색량 조회',
    '다음 시트들에서 검색량을 조회합니다:\n\n- ' + sheetNames + '\n\n' +
    '각 시트의 A열 키워드를 조회하여 B,C열에 입력합니다.\n계속하시겠습니까?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  let totalReport = '=== 전체 시트 검색량 조회 결과 ===\n\n';
  let totalKeywords = 0;

  sheets.forEach(sheet => {
    try {
      const lastRow = sheet.getLastRow();
      if (lastRow < CONFIG.SOURCE.START_ROW) {
        totalReport += '[' + sheet.getName() + '] 데이터 없음 (스킵)\n';
        return;
      }

      // A열에서 키워드 가져오기
      const keywordRange = sheet.getRange(
        CONFIG.SOURCE.START_ROW,
        CONFIG.SOURCE.KEYWORD_COL,
        lastRow - CONFIG.SOURCE.START_ROW + 1,
        1
      );
      const keywordValues = keywordRange.getValues();
      const keywords = keywordValues.map(row => String(row[0]).trim()).filter(k => k);

      if (keywords.length === 0) {
        totalReport += '[' + sheet.getName() + '] 키워드 없음 (스킵)\n';
        return;
      }

      // 네이버 API로 검색량 조회
      const stats = getKeywordStats(keywords);

      // B, C열에 검색량 입력
      let updatedCount = 0;
      for (let i = 0; i < keywordValues.length; i++) {
        const keyword = String(keywordValues[i][0]).trim();
        if (keyword && stats[keyword]) {
          const row = CONFIG.SOURCE.START_ROW + i;
          sheet.getRange(row, CONFIG.SOURCE.PC_COL).setValue(stats[keyword].pc);
          sheet.getRange(row, CONFIG.SOURCE.M_COL).setValue(stats[keyword].mobile);
          updatedCount++;
        }
      }

      totalReport += '[' + sheet.getName() + '] ' + updatedCount + '/' + keywords.length + '개 업데이트\n';
      totalKeywords += updatedCount;

    } catch (e) {
      totalReport += '[' + sheet.getName() + '] 오류: ' + e.message + '\n';
    }
  });

  totalReport += '\n총 ' + totalKeywords + '개 키워드 업데이트됨';
  showMessage(totalReport);
}

/**
 * 특정 시트에서 소스 데이터 가져오기
 */
function getSourceDataFromSheet(sheet) {
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

// ========================================
// 입력 데이터 변환 기능
// ========================================

/**
 * 입력 데이터 변환
 * - A열 빈칸에 위의 검색키워드 채우기
 * - 카페가 있는 키워드 우선 정렬 (각 그룹 내에서)
 * - 검색키워드 중복 시 상단 그룹만 유지
 * - PC/M 검색량은 공란으로
 *
 * 입력 형식: A(검색키워드), B(주제키워드), C~E(콘텐츠영역)
 * 출력 형식: A(검색키워드), B(PC), C(M), D(주제키워드), E~G(콘텐츠영역)
 */
function 데이터변환() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // 입력 데이터는 1행=헤더, 2행부터 데이터
  const INPUT_START_ROW = 2;

  // 확인 메시지
  const confirm = ui.alert(
    '데이터 변환',
    '입력 데이터를 변환합니다.\n\n' +
    '[수행 작업]\n' +
    '1. A열 빈칸에 위의 검색키워드 채우기\n' +
    '2. 카페가 있는 키워드를 각 그룹 상위로 정렬\n' +
    '3. 검색키워드 중복 시 상단 그룹만 유지\n' +
    '4. PC/M 검색량은 공란으로\n\n' +
    '계속하시겠습니까?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  try {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < INPUT_START_ROW) {
      showMessage('데이터가 없습니다.');
      return;
    }

    ss.toast('데이터 읽는 중...', '진행 중', -1);

    // 데이터 범위 읽기 (A~E열 또는 더 많은 열) - 2행부터
    const dataRange = sheet.getRange(
      INPUT_START_ROW,
      1,
      lastRow - INPUT_START_ROW + 1,
      Math.max(lastCol, 5)
    );
    const rawData = dataRange.getValues();

    // 1. A열 빈칸 채우기 + 그룹화
    ss.toast('A열 빈칸 채우기 및 그룹화 중...', '진행 중', -1);
    const groups = fillAndGroup(rawData);

    // 2. 중복 검색키워드 제거 (상단 그룹만 유지)
    ss.toast('중복 제거 중...', '진행 중', -1);
    const uniqueGroups = removeDuplicateGroups(groups);

    // 3. 각 그룹 내에서 카페 우선 정렬
    ss.toast('카페 우선 정렬 중...', '진행 중', -1);
    const sortedGroups = sortGroupsByCafe(uniqueGroups);

    // 4. 출력 형식으로 변환 (PC/M은 공란)
    ss.toast('출력 형식으로 변환 중...', '진행 중', -1);
    const outputData = convertToOutputFormat(sortedGroups);

    // 5. 기존 데이터 삭제 후 새 데이터 입력
    ss.toast('시트에 출력 중...', '진행 중', -1);

    // A~G열까지의 기존 데이터 삭제 (1행 헤더 포함)
    sheet.getRange(1, 1, lastRow, 7).clearContent();

    // 헤더 입력 (1행)
    const header = ['검색키워드', 'PC', 'M', '주제키워드', '콘텐츠영역', '콘텐츠영역', '콘텐츠영역'];
    sheet.getRange(1, 1, 1, 7).setValues([header]);

    // 새 데이터 입력 (2행부터)
    if (outputData.length > 0) {
      const dataRange = sheet.getRange(
        INPUT_START_ROW,
        1,
        outputData.length,
        outputData[0].length
      );
      dataRange.setValues(outputData);

      // 전체 범위 가운데 정렬 (헤더 + 데이터)
      sheet.getRange(1, 1, outputData.length + 1, 7)
        .setHorizontalAlignment('center');
    }

    // 결과 보고
    const originalRows = rawData.filter(row => row[0] || row[1]).length;
    const resultRows = outputData.length;
    const removedGroups = groups.length - uniqueGroups.length;

    showMessage('데이터 변환 완료!\n\n' +
      '- 원본 행: ' + originalRows + '개\n' +
      '- 결과 행: ' + resultRows + '개\n' +
      '- 제거된 중복 그룹: ' + removedGroups + '개\n' +
      '- 고유 검색키워드: ' + uniqueGroups.length + '개');

  } catch (error) {
    showMessage('오류 발생: ' + error.message);
    console.error(error);
  }
}

/**
 * A열 빈칸 채우기 및 그룹화
 * 빈칸인 경우 위의 검색키워드를 채움
 * 같은 검색키워드를 가진 행들을 그룹으로 묶음
 */
function fillAndGroup(data) {
  const groups = [];
  let currentKeyword = '';
  let currentGroup = [];

  data.forEach(row => {
    const cellA = String(row[0] || '').trim();
    const cellB = String(row[1] || '').trim();

    // A열에 값이 있으면 새 그룹 시작
    if (cellA !== '') {
      // 이전 그룹 저장
      if (currentGroup.length > 0) {
        groups.push({
          keyword: currentKeyword,
          rows: currentGroup
        });
      }
      currentKeyword = cellA;
      currentGroup = [row];
    } else if (cellB !== '') {
      // A열은 비어있고 B열에 값이 있으면 현재 그룹에 추가
      currentGroup.push(row);
    }
    // A, B 모두 비어있으면 무시
  });

  // 마지막 그룹 저장
  if (currentGroup.length > 0) {
    groups.push({
      keyword: currentKeyword,
      rows: currentGroup
    });
  }

  return groups;
}

/**
 * 중복 검색키워드 제거 (상단 그룹만 유지)
 */
function removeDuplicateGroups(groups) {
  const seen = new Set();
  const uniqueGroups = [];

  groups.forEach(group => {
    if (!seen.has(group.keyword)) {
      seen.add(group.keyword);
      uniqueGroups.push(group);
    }
  });

  return uniqueGroups;
}

/**
 * 각 그룹 내에서 카페 개수 기준 정렬
 * 콘텐츠영역(C~E열)에 '카페'가 많은 행일수록 상위로
 * 3개 > 2개 > 1개 > 0개 (카페 없는 행도 하단에 유지)
 */
function sortGroupsByCafe(groups) {
  return groups.map(group => {
    // 카페 개수 많은 순으로 정렬 (내림차순)
    // 카페 없는 행(0개)도 하단에 유지
    const sortedRows = [...group.rows].sort((a, b) => {
      const cafeCountA = countCafe(a);
      const cafeCountB = countCafe(b);
      return cafeCountB - cafeCountA;
    });

    return {
      keyword: group.keyword,
      rows: sortedRows
    };
  });
}

/**
 * 행에서 '카페' 개수 세기
 * C, D, E열 (인덱스 2, 3, 4)에서 카페 포함된 셀 개수 반환
 */
function countCafe(row) {
  let count = 0;
  for (let i = 2; i <= 4; i++) {
    if (String(row[i] || '').includes('카페')) {
      count++;
    }
  }
  return count;
}

/**
 * 출력 형식으로 변환
 * 입력: A(검색키워드), B(주제키워드), C~E(콘텐츠영역)
 * 출력: A(검색키워드), B(PC-공란), C(M-공란), D(주제키워드), E~G(콘텐츠영역)
 *
 * 모든 행에 검색키워드를 채워넣음 (빈칸 없이)
 */
function convertToOutputFormat(groups) {
  const output = [];

  groups.forEach(group => {
    group.rows.forEach(row => {
      const topicKeyword = row[1] || '';  // B열: 주제키워드
      const content1 = row[2] || '';  // C열: 콘텐츠영역
      const content2 = row[3] || '';  // D열: 콘텐츠영역
      const content3 = row[4] || '';  // E열: 콘텐츠영역

      output.push([
        group.keyword, // A: 검색키워드 (모든 행에 채움)
        '',            // B: PC (공란)
        '',            // C: M (공란)
        topicKeyword,  // D: 주제키워드
        content1,      // E: 콘텐츠영역
        content2,      // F: 콘텐츠영역.1
        content3       // G: 콘텐츠영역.2
      ]);
    });
  });

  return output;
}
