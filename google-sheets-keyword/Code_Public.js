/**
 * 키워드 자동 분류 시스템 (배포용 - 검색량 조회 기능 제외)
 *
 * [주요 기능]
 * 1. 데이터 변환 (입력→출력): 카페 키워드 필터링 및 정렬
 * 2. 데이터 재배열: 카페 우선 재정렬
 * 3. 중복 제거: A열 키워드 중복 제거
 *
 * [사용 방법]
 * 1. 구글 시트에서 "확장 프로그램 > Apps Script" 메뉴 클릭
 * 2. 이 코드를 전체 복사하여 붙여넣기
 * 3. 저장 후 시트로 돌아가서 새로고침 (F5)
 * 4. 상단 메뉴에 "키워드 도구" 메뉴 표시됨
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
    .addItem('🔄 데이터 재배열 (카페 우선)', '데이터재배열')
    .addSeparator()
    .addItem('중복 제거 (A~D열)', '중복제거')
    .addSeparator()
    .addItem('도움말', '도움말')
    .addToUi();
}

/**
 * 도움말
 */
function 도움말() {
  const help =
'=== 키워드 분류 도구 (배포용) ===\n\n' +
'[데이터 변환 기능]\n' +
'입력 데이터를 카페 키워드 중심으로 정리합니다.\n\n' +
'1. A열 빈칸에 위의 검색키워드 자동 채우기\n' +
'2. 카페가 있는 키워드만 필터링\n' +
'3. 검색키워드 중복 시 상단 그룹만 유지\n' +
'4. 카페 개수 많은 순으로 정렬 (3개 > 2개 > 1개)\n\n' +
'[데이터 재배열 기능]\n' +
'이미 변환된 데이터를 다시 정렬합니다.\n\n' +
'1. 검색키워드별로 그룹화\n' +
'2. 카페가 없는 키워드 제거\n' +
'3. 카페 개수 많은 순으로 정렬\n\n' +
'[중복 제거 기능]\n' +
'A열 키워드 기준으로 중복된 행을 삭제합니다.\n\n' +
'※ PC/M 검색량은 네이버 API가 필요하므로 배포용에서는 제외되었습니다.';

  showMessage(help);
}

// ========================================
// 입력 데이터 변환 기능
// ========================================

/**
 * 입력 데이터 변환
 * - A열 빈칸에 위의 검색키워드 채우기
 * - 카페가 있는 키워드만 필터링
 * - 검색키워드 중복 시 상단 그룹만 유지
 * - 카페 개수 많은 순으로 정렬 (3개 > 2개 > 1개)
 *
 * 입력 형식: A(검색키워드), B(주제키워드), C~E(콘텐츠영역)
 * 출력 형식: A(검색키워드), B(PC-공란), C(M-공란), D(주제키워드), E~G(콘텐츠영역)
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
    '2. 카페가 있는 키워드만 필터링\n' +
    '3. 검색키워드 중복 시 상단 그룹만 유지\n' +
    '4. 카페 개수 많은 순으로 정렬 (3개 > 2개 > 1개)\n\n' +
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
    const cafeFilteredGroups = uniqueGroups.length - sortedGroups.length;

    showMessage('데이터 변환 완료!\n\n' +
      '- 원본 행: ' + originalRows + '개\n' +
      '- 결과 행: ' + resultRows + '개\n' +
      '- 제거된 중복 그룹: ' + removedGroups + '개\n' +
      '- 제거된 그룹 (카페 없음): ' + cafeFilteredGroups + '개\n' +
      '- 고유 검색키워드: ' + sortedGroups.length + '개');

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
 * 각 그룹 내에서 카페 개수 기준 정렬 + 카페 필터링
 * 콘텐츠영역(C~E열)에 '카페'가 있는 행만 유지
 * 카페 개수: 3개 > 2개 > 1개 순으로 정렬
 */
function sortGroupsByCafe(groups) {
  return groups.map(group => {
    // 카페가 1개 이상 있는 행만 필터링
    const cafeRows = group.rows.filter(row => countCafe(row) > 0);

    // 카페 개수 많은 순으로 정렬 (내림차순)
    const sortedRows = [...cafeRows].sort((a, b) => {
      const cafeCountA = countCafe(a);
      const cafeCountB = countCafe(b);
      return cafeCountB - cafeCountA;
    });

    return {
      keyword: group.keyword,
      rows: sortedRows
    };
  }).filter(group => group.rows.length > 0); // 카페가 하나도 없는 그룹은 제외
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

// ========================================
// 재배열 기능 (카페 우선 정렬)
// ========================================

/**
 * 출력 데이터 재배열 (카페 우선)
 * A~G열의 기존 데이터를 읽어서 카페 우선으로 다시 정렬
 */
function 데이터재배열() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // 출력 데이터는 1행=헤더, 2행부터 데이터
  const OUTPUT_START_ROW = 2;

  // 확인 메시지
  const response = ui.alert(
    '데이터 재배열',
    '현재 시트의 데이터를 카페 우선으로 재정렬합니다.\n\n' +
    '[수행 작업]\n' +
    '1. 검색키워드별로 그룹화\n' +
    '2. 카페가 없는 키워드 제거\n' +
    '3. 카페 개수 많은 순으로 정렬 (3개 > 2개 > 1개)\n\n' +
    '계속하시겠습니까?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    const lastRow = sheet.getLastRow();

    if (lastRow < OUTPUT_START_ROW) {
      showMessage('재배열할 데이터가 없습니다.');
      return;
    }

    ss.toast('데이터 읽는 중...', '진행 중', -1);

    // A~G열 읽기 (검색키워드, PC, M, 주제키워드, 콘텐츠영역x3)
    const dataRange = sheet.getRange(
      OUTPUT_START_ROW,
      1,
      lastRow - OUTPUT_START_ROW + 1,
      7
    );
    const rawData = dataRange.getValues();

    // 빈 행 제거
    const validData = rawData.filter(row => row[0]); // A열(검색키워드)이 있는 행만

    if (validData.length === 0) {
      showMessage('재배열할 데이터가 없습니다.');
      return;
    }

    // 1. 검색키워드별 그룹화
    ss.toast('그룹화 중...', '진행 중', -1);
    const groups = groupOutputData(validData);

    // 2. 카페 필터링 + 정렬
    ss.toast('카페 우선 정렬 중...', '진행 중', -1);
    const sortedGroups = sortOutputGroupsByCafe(groups);

    // 3. 다시 평탄화
    const sortedData = flattenOutputGroups(sortedGroups);

    if (sortedData.length === 0) {
      showMessage('카페가 있는 키워드가 없습니다.');
      return;
    }

    // 4. 기존 데이터 삭제 후 새 데이터 입력
    ss.toast('시트에 출력 중...', '진행 중', -1);

    // A~G열 기존 데이터 삭제 (헤더 제외, 2행부터)
    sheet.getRange(OUTPUT_START_ROW, 1, lastRow - OUTPUT_START_ROW + 1, 7).clearContent();

    // 새 데이터 입력
    if (sortedData.length > 0) {
      const outputRange = sheet.getRange(
        OUTPUT_START_ROW,
        1,
        sortedData.length,
        7
      );
      outputRange.setValues(sortedData);

      // 전체 범위 가운데 정렬
      sheet.getRange(OUTPUT_START_ROW, 1, sortedData.length, 7)
        .setHorizontalAlignment('center');
    }

    // 결과 보고
    const originalRows = validData.length;
    const resultRows = sortedData.length;
    const removedRows = originalRows - resultRows;

    showMessage('데이터 재배열 완료!\n\n' +
      '- 원본 행: ' + originalRows + '개\n' +
      '- 결과 행: ' + resultRows + '개\n' +
      '- 제거된 행 (카페 없음): ' + removedRows + '개');

  } catch (error) {
    showMessage('오류 발생: ' + error.message);
    console.error(error);
  }
}

/**
 * 출력 데이터를 검색키워드별로 그룹화
 * 입력: [검색키워드, PC, M, 주제키워드, 콘텐츠1, 콘텐츠2, 콘텐츠3]
 */
function groupOutputData(data) {
  const groups = [];
  let currentKeyword = '';
  let currentGroup = [];

  data.forEach(row => {
    const searchKeyword = String(row[0]).trim();

    if (searchKeyword === '') return; // 빈 행 무시

    // 새 키워드면 새 그룹 시작
    if (searchKeyword !== currentKeyword) {
      // 이전 그룹 저장
      if (currentGroup.length > 0) {
        groups.push({
          keyword: currentKeyword,
          rows: currentGroup
        });
      }
      currentKeyword = searchKeyword;
      currentGroup = [row];
    } else {
      // 같은 키워드면 현재 그룹에 추가
      currentGroup.push(row);
    }
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
 * 출력 데이터 그룹을 카페 우선으로 정렬
 * 콘텐츠영역(E, F, G열 = 인덱스 4, 5, 6)에서 카페 검사
 */
function sortOutputGroupsByCafe(groups) {
  return groups.map(group => {
    // 카페가 1개 이상 있는 행만 필터링
    const cafeRows = group.rows.filter(row => countCafeInOutput(row) > 0);

    // 카페 개수 많은 순으로 정렬 (내림차순)
    const sortedRows = [...cafeRows].sort((a, b) => {
      const cafeCountA = countCafeInOutput(a);
      const cafeCountB = countCafeInOutput(b);
      return cafeCountB - cafeCountA;
    });

    return {
      keyword: group.keyword,
      rows: sortedRows
    };
  }).filter(group => group.rows.length > 0); // 카페가 하나도 없는 그룹은 제외
}

/**
 * 출력 데이터 행에서 '카페' 개수 세기
 * E, F, G열 (인덱스 4, 5, 6)에서 카페 포함된 셀 개수 반환
 */
function countCafeInOutput(row) {
  let count = 0;
  for (let i = 4; i <= 6; i++) {
    if (String(row[i] || '').includes('카페')) {
      count++;
    }
  }
  return count;
}

/**
 * 그룹을 평탄화하여 2차원 배열로 변환
 */
function flattenOutputGroups(groups) {
  const output = [];
  groups.forEach(group => {
    group.rows.forEach(row => {
      output.push(row);
    });
  });
  return output;
}

// ========================================
// 중복 제거 기능
// ========================================

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
