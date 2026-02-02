/**
 * 바이럴 GFS - SET 기반 댓글 프롬프트 시스템
 */

/** 메뉴 추가 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 SET 관리')
    .addItem('✅ SET 드롭다운 & 수식 적용', 'applySetSystem')
    .addItem('📋 SET 2,3 프롬프트 영역 생성', 'createSetTemplates')
    .addToUi();
}

/**
 * R열에 드롭다운 추가 + 댓글 수식을 INDIRECT 기반으로 변경
 */
function applySetSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('LI1★');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ LI1★ 시트를 찾을 수 없습니다.');
    return;
  }

  // 데이터 범위 확인 (3행부터 마지막 데이터까지)
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    SpreadsheetApp.getUi().alert('❌ 데이터가 없습니다.');
    return;
  }

  // 1. R열에 드롭다운 추가 (R3부터) - "1번", "2번", "3번" 형식
  const rRange = sheet.getRange(3, 18, lastRow - 2); // R열 = 18번째
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1번', '2번', '3번'], true)
    .setAllowInvalid(false)
    .build();
  rRange.setDataValidation(rule);

  // 기존 값이 없는 셀만 기본값 설정
  const currentValues = rRange.getValues();
  const newValues = currentValues.map(row => {
    if (!row[0] || row[0] === '') return ['1번'];
    return row;
  });
  rRange.setValues(newValues);

  // 2. 댓글 수식 적용 (S~AB열)
  // S=19, T=20, U=21, V=22, W=23, X=24, Y=25, Z=26, AA=27, AB=28
  // VALUE(LEFT($R행,1)) 로 "1번"에서 숫자 1 추출

  const commentFormulas = [
    // S열: 댓글1 (AI18 기반)
    { col: 19, baseRow: 18, simple: true },
    // T열: 댓글2 (AI19 기반)
    { col: 20, baseRow: 19, simple: true },
    // U열: 댓글3 (AI20 기반)
    { col: 21, baseRow: 20, simple: true },
    // V열: 댓글4 (AI21 기반)
    { col: 22, baseRow: 21, simple: true },
    // W열: 댓글1-1 (AI22 기반)
    { col: 23, baseRow: 22, simple: true },
    // X열: 댓글2-1 (AI23 기반)
    { col: 24, baseRow: 23, simple: true },
    // Y열: 댓글3-1 (AI24 기반)
    { col: 25, baseRow: 24, simple: true },
    // Z열: 댓글4-1 (AI25 기반)
    { col: 26, baseRow: 25, simple: true },
    // AA열: 댓글2-2 (AI26 기반) - 복잡한 수식
    { col: 27, baseRow: 26, simple: false },
    // AB열: 댓글2-3 (AI27 기반)
    { col: 28, baseRow: 27, simple: true }
  ];

  for (let row = 3; row <= lastRow; row++) {
    commentFormulas.forEach(cf => {
      let formula;
      if (cf.simple) {
        // 단순 수식: VALUE(LEFT())로 "1번"→1 변환
        formula = `=gpt(INDIRECT("$AI$"&(${cf.baseRow}+(VALUE(LEFT($R${row},1))-1)*10)),$D${row},1)`;
      } else {
        // AA열 복잡한 수식
        formula = `=GPT("이전 대화 맥락을 기반으로 댓글 2-2를 작성해줘. 맥락은 아래와 같아:"&CHAR(10)&CHAR(10)&"[댓글 2 내용] "&$T${row}&CHAR(10)&"[댓글 2-1 내용] "&$X${row}&CHAR(10)&CHAR(10)&"이때 반드시 지켜야 할 조건은 다음과 같아:"&CHAR(10)&"- 댓글 2에서 언급된 '전문가' 관계(예: 이모부)를 그대로 유지할 것"&CHAR(10)&"- 댓글 2-1에서 제품명을 물었으므로, 이 제품명을 자연스럽고 친절하게 명시할 것"&CHAR(10)&"- 전체적인 구어체 톤과 이모티콘, 체중 변화 수치는 공통지침을 따를 것"&CHAR(10)&CHAR(10)&INDIRECT("$AI$"&(${cf.baseRow}+(VALUE(LEFT($R${row},1))-1)*10)))`;
      }
      sheet.getRange(row, cf.col).setFormula(formula);
    });
  }

  SpreadsheetApp.getUi().alert(`✅ 완료!\n\nR열 드롭다운: 1번, 2번, 3번\n댓글 수식 ${lastRow - 2}행 적용됨\n\n이제 R열에서 SET 번호를 바꾸면 자동으로 해당 프롬프트가 적용됩니다.\n\n※ SET 2, 3 프롬프트가 없으면 메뉴에서 '📋 SET 2,3 프롬프트 영역 생성'을 실행하세요.`);
}

/**
 * SET 2, 3 프롬프트 영역 생성 (AI28~37, AI38~47)
 */
function createSetTemplates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('LI1★');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ LI1★ 시트를 찾을 수 없습니다.');
    return;
  }

  // SET 1 프롬프트 복사 (AI18:AI27)
  const set1Range = sheet.getRange('AI18:AI27');
  const set1Values = set1Range.getValues();

  // SET 2 영역에 복사 (AI28:AI37)
  sheet.getRange('AI28:AI37').setValues(set1Values);
  sheet.getRange('AI28').setNote('SET 2 프롬프트 시작');

  // SET 3 영역에 복사 (AI38:AI47)
  sheet.getRange('AI38:AI47').setValues(set1Values);
  sheet.getRange('AI38').setNote('SET 3 프롬프트 시작');

  // 배경색 구분
  sheet.getRange('AI18:AI27').setBackground('#E8F5E9'); // SET 1: 연한 초록
  sheet.getRange('AI28:AI37').setBackground('#E3F2FD'); // SET 2: 연한 파랑
  sheet.getRange('AI38:AI47').setBackground('#FFF3E0'); // SET 3: 연한 주황

  SpreadsheetApp.getUi().alert(`✅ SET 2, 3 프롬프트 영역 생성 완료!\n\nSET 1: AI18~AI27 (초록)\nSET 2: AI28~AI37 (파랑)\nSET 3: AI38~AI47 (주황)\n\n각 SET의 프롬프트를 원하는대로 수정하세요.`);
}
