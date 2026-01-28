/**
 * 섬고짚 네이버광고 분석 시스템
 * 키워드/검색어 보고서를 분석하고 종합 제언을 생성합니다.
 */

// ============================================
// 1. 메뉴 생성
// ============================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 광고 분석')
    .addItem('📋 종합 분석 (최근 6개월)', 'runSummaryAnalysis')
    .addItem('📅 월별 분석', 'runMonthlyAnalysis')
    .addSeparator()
    .addItem('🗑️ 종합탭 초기화', 'clearSummarySheet')
    .addItem('🗑️ 월별탭 초기화', 'clearMonthlySheet')
    .addToUi();
}

// ============================================
// 2. 종합 분석 함수 (최근 6개월)
// ============================================
function runSummaryAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('시트를 탐색하는 중...', '분석 시작', 3);

    // 동적으로 모든 보고서 시트 찾기
    const { campaignSheets, searchSheets } = findReportSheets(ss);

    if (searchSheets.length === 0) {
      ui.alert('오류', '검색어 보고서 시트를 찾을 수 없습니다.\n시트 이름에 "검색어 보고서"가 포함되어야 합니다.', ui.ButtonSet.OK);
      return;
    }

    ss.toast('데이터를 불러오는 중...', '진행 중', 3);

    // 최근 6개월만 선택
    const recentSearchSheets = searchSheets.slice(-6);
    const recentCampaignSheets = campaignSheets.slice(-6);

    // 데이터 수집
    // 총계용: 모든 데이터 포함 (filterInvalid=false)
    // 검색어 분석용: 유효한 키워드만 (filterInvalid=true)
    const searchDataAll = {};  // 총계 계산용 (전체)
    const searchDataFiltered = {};  // 검색어 분석용 (필터링)
    const campaignData = {};

    recentSearchSheets.forEach(s => {
      searchDataAll[s.month] = getReportData(ss, s.name, false); // 전체 데이터
      searchDataFiltered[s.month] = getReportData(ss, s.name, true); // 유효 키워드만
    });

    recentCampaignSheets.forEach(s => {
      campaignData[s.month] = getReportData(ss, s.name, false); // 전체 데이터
    });

    const months = recentSearchSheets.map(s => s.month);
    const latestMonth = months[months.length - 1];
    const prevMonth = months.length > 1 ? months[months.length - 2] : null;

    ss.toast('데이터 분석 중...', '진행 중', 3);

    // 분석 실행
    const analysis = {
      months: months,
      latestMonth: latestMonth,
      prevMonth: prevMonth,
      summaryByMonth: {},
      byTypeByMonth: {},
      byDistanceByMonth: {},
      // 검색어 보고서 기준 Top/Bottom (필터링된 데이터 사용)
      topKeywords: searchDataFiltered[latestMonth] ? getTopItems(searchDataFiltered[latestMonth], 'ctr', 10, true) : [],
      bottomKeywords: searchDataFiltered[latestMonth] ? getTopItems(searchDataFiltered[latestMonth], 'ctr', 10, false) : [],
      topCpcKeywords: searchDataFiltered[latestMonth] ? getTopItems(searchDataFiltered[latestMonth], 'cpc', 10, false) : [],
      bottomCpcKeywords: searchDataFiltered[latestMonth] ? getTopItems(searchDataFiltered[latestMonth], 'cpc', 10, true) : [],
      // 검색어 인사이트 (필터링된 데이터 사용)
      topSearchTerms: searchDataFiltered[latestMonth] ? getTopSearchTerms(searchDataFiltered[latestMonth], 10) : [],
      bottomSearchTerms: searchDataFiltered[latestMonth] ? getBottomSearchTerms(searchDataFiltered[latestMonth], 10) : [],
      monthComparison: null
    };

    // 각 월별 분석 - 총계는 전체 데이터 사용 (캠페인 또는 검색어 전체)
    months.forEach(month => {
      const data = campaignData[month] || searchDataAll[month];
      if (data) {
        analysis.summaryByMonth[month] = calculateSummary(data);
        analysis.byTypeByMonth[month] = analyzeByType(data);
        analysis.byDistanceByMonth[month] = analyzeByDistance(data);
      }
    });

    // 최근 2개월 비교
    if (prevMonth && analysis.summaryByMonth[prevMonth] && analysis.summaryByMonth[latestMonth]) {
      analysis.monthComparison = compareMonths(
        analysis.summaryByMonth[prevMonth],
        analysis.summaryByMonth[latestMonth]
      );
    }

    ss.toast('제언 생성 중...', '진행 중', 3);
    const recommendations = generateRecommendations(analysis);

    ss.toast('결과 출력 중...', '진행 중', 3);
    writeToSummarySheet(ss, analysis, recommendations);

    ss.toast('종합 분석이 완료되었습니다!', '완료', 5);

  } catch (error) {
    ui.alert('오류 발생', '분석 중 오류가 발생했습니다: ' + error.message, ui.ButtonSet.OK);
    Logger.log(error);
  }
}

// ============================================
// 2-1. 월별 분석 함수
// ============================================
function runMonthlyAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    // 사용 가능한 월 찾기
    const { campaignSheets, searchSheets } = findReportSheets(ss);
    const availableMonths = searchSheets.map(s => s.month);

    if (availableMonths.length === 0) {
      ui.alert('오류', '검색어 보고서 시트를 찾을 수 없습니다.', ui.ButtonSet.OK);
      return;
    }

    // 월 선택 프롬프트
    const response = ui.prompt(
      '📅 월별 분석',
      '분석할 월을 입력하세요.\n\n사용 가능: ' + availableMonths.join(', ') + '\n\n예: 1월',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    const selectedMonth = response.getResponseText().trim();

    // 유효성 검사
    if (!availableMonths.includes(selectedMonth)) {
      ui.alert('오류', '잘못된 월입니다. 다시 시도해주세요.\n사용 가능: ' + availableMonths.join(', '), ui.ButtonSet.OK);
      return;
    }

    ss.toast(selectedMonth + ' 데이터를 불러오는 중...', '진행 중', 3);

    // 해당 월 데이터 가져오기
    const searchSheet = searchSheets.find(s => s.month === selectedMonth);
    const campaignSheet = campaignSheets.find(s => s.month === selectedMonth);

    // 총계용: 전체 데이터, 검색어 분석용: 필터링된 데이터
    const searchDataAll = searchSheet ? getReportData(ss, searchSheet.name, false) : [];
    const searchDataFiltered = searchSheet ? getReportData(ss, searchSheet.name, true) : [];
    const campaignData = campaignSheet ? getReportData(ss, campaignSheet.name, false) : [];

    // 이전 월 찾기 (비교용)
    const monthIndex = availableMonths.indexOf(selectedMonth);
    const prevMonth = monthIndex > 0 ? availableMonths[monthIndex - 1] : null;
    let prevSearchDataAll = null;
    let prevCampaignData = null;

    if (prevMonth) {
      const prevSearchSheet = searchSheets.find(s => s.month === prevMonth);
      const prevCampaignSheet = campaignSheets.find(s => s.month === prevMonth);
      prevSearchDataAll = prevSearchSheet ? getReportData(ss, prevSearchSheet.name, false) : [];
      prevCampaignData = prevCampaignSheet ? getReportData(ss, prevCampaignSheet.name, false) : [];
    }

    ss.toast('데이터 분석 중...', '진행 중', 3);

    // 총계 계산용 데이터 선택
    const summaryData = campaignData.length > 0 ? campaignData : searchDataAll;
    const prevSummaryData = prevMonth ? (prevCampaignData && prevCampaignData.length > 0 ? prevCampaignData : prevSearchDataAll) : null;

    // 월별 분석
    const analysis = {
      selectedMonth: selectedMonth,
      prevMonth: prevMonth,
      // 현재 월 요약 (전체 데이터)
      summary: calculateSummary(summaryData),
      prevSummary: prevSummaryData ? calculateSummary(prevSummaryData) : null,
      // 캠페인유형별 (전체 데이터)
      byType: analyzeByType(summaryData),
      // 거리별 (전체 데이터)
      byDistance: analyzeByDistance(summaryData),
      // 키워드 분석 (필터링된 검색어 데이터)
      topKeywords: getTopItems(searchDataFiltered, 'ctr', 10, true),
      bottomKeywords: getTopItems(searchDataFiltered, 'ctr', 10, false),
      topCpcKeywords: getTopItems(searchDataFiltered, 'cpc', 10, false),
      bottomCpcKeywords: getTopItems(searchDataFiltered, 'cpc', 10, true),
      // 검색어 분석 (필터링된 데이터)
      topSearchTerms: getTopSearchTerms(searchDataFiltered, 10),
      bottomSearchTerms: getBottomSearchTerms(searchDataFiltered, 10),
      // 월별 비교
      monthComparison: null
    };

    if (prevMonth && analysis.prevSummary) {
      analysis.monthComparison = compareMonths(analysis.prevSummary, analysis.summary);
    }

    ss.toast('결과 출력 중...', '진행 중', 3);
    writeToMonthlySheet(ss, analysis);

    ss.toast(selectedMonth + ' 분석이 완료되었습니다!', '완료', 5);

  } catch (error) {
    ui.alert('오류 발생', '분석 중 오류가 발생했습니다: ' + error.message, ui.ButtonSet.OK);
    Logger.log(error);
  }
}

// 보고서 시트 찾기
function findReportSheets(ss) {
  const sheets = ss.getSheets();
  const campaignSheets = [];
  const searchSheets = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    // 키워드 보고서 = 캠페인 보고서로 취급
    if (name.includes('키워드 보고서') || name.includes('캠페인 보고서')) {
      const month = extractMonth(name);
      if (month) campaignSheets.push({ name, month, order: getMonthOrder(month) });
    } else if (name.includes('검색어 보고서')) {
      const month = extractMonth(name);
      if (month) searchSheets.push({ name, month, order: getMonthOrder(month) });
    }
  });

  // 월 순서대로 정렬
  campaignSheets.sort((a, b) => a.order - b.order);
  searchSheets.sort((a, b) => a.order - b.order);

  return { campaignSheets, searchSheets };
}

// 시트 이름에서 월 추출 (예: "12월 키워드 보고서" -> "12월")
function extractMonth(sheetName) {
  const match = sheetName.match(/(\d{1,2}월)/);
  return match ? match[1] : null;
}

// 월을 정렬 순서로 변환 (1월=1, 12월=12)
function getMonthOrder(month) {
  const num = parseInt(month.replace('월', ''));
  return num;
}

// ============================================
// 3. 데이터 읽기 함수
// ============================================
function getReportData(ss, sheetName, filterInvalid) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('시트를 찾을 수 없습니다: ' + sheetName);
  }

  const data = sheet.getDataRange().getValues();
  const result = [];

  // 헤더 행 찾기 (캠페인유형이 있는 행)
  let headerRow = -1;
  for (let i = 0; i < Math.min(5, data.length); i++) {
    if (data[i][0] === '캠페인유형') {
      headerRow = i;
      break;
    }
  }

  if (headerRow === -1) {
    throw new Error('헤더를 찾을 수 없습니다: ' + sheetName);
  }

  // 데이터 파싱
  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || row[0] === '') continue; // 빈 행 스킵

    const keyword = row[3];

    // 이상치 필터링: 키워드가 "-" 이거나 빈 값이면 스킵
    if (filterInvalid && (!keyword || keyword === '-' || keyword === '')) {
      continue;
    }

    result.push({
      type: row[0],           // 캠페인유형
      campaign: row[1],       // 캠페인
      adGroup: row[2],        // 광고그룹
      keyword: keyword,       // 키워드 또는 검색어
      impressions: parseNumber(row[4]),  // 노출수
      clicks: parseNumber(row[5]),       // 클릭수
      ctr: parsePercent(row[6]),         // 클릭률
      cpc: parseNumber(row[7]),          // 평균클릭비용
      cost: parseNumber(row[8])          // 총비용
    });
  }

  return result;
}

function parseNumber(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;
  return parseFloat(String(value).replace(/,/g, '')) || 0;
}

function parsePercent(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;
  return parseFloat(String(value).replace('%', '')) || 0;
}

// ============================================
// 4. 분석 로직
// ============================================

// 전체 요약 계산
function calculateSummary(data) {
  const summary = {
    totalImpressions: 0,
    totalClicks: 0,
    totalCost: 0,
    count: data.length
  };

  data.forEach(item => {
    summary.totalImpressions += item.impressions;
    summary.totalClicks += item.clicks;
    summary.totalCost += item.cost;
  });

  summary.avgCtr = summary.totalImpressions > 0
    ? (summary.totalClicks / summary.totalImpressions * 100)
    : 0;
  summary.avgCpc = summary.totalClicks > 0
    ? (summary.totalCost / summary.totalClicks)
    : 0;

  return summary;
}

// 캠페인유형별 분석
function analyzeByType(data) {
  const types = {};

  data.forEach(item => {
    if (!types[item.type]) {
      types[item.type] = {
        impressions: 0,
        clicks: 0,
        cost: 0,
        count: 0
      };
    }
    types[item.type].impressions += item.impressions;
    types[item.type].clicks += item.clicks;
    types[item.type].cost += item.cost;
    types[item.type].count++;
  });

  // CTR, CPC 계산
  Object.keys(types).forEach(type => {
    const t = types[type];
    t.ctr = t.impressions > 0 ? (t.clicks / t.impressions * 100) : 0;
    t.cpc = t.clicks > 0 ? (t.cost / t.clicks) : 0;
  });

  return types;
}

// 거리별 분석 (캠페인명에서 추출)
function analyzeByDistance(data) {
  const distances = {
    '1km': { impressions: 0, clicks: 0, cost: 0, count: 0 },
    '3km': { impressions: 0, clicks: 0, cost: 0, count: 0 },
    '5km': { impressions: 0, clicks: 0, cost: 0, count: 0 },
    '10km': { impressions: 0, clicks: 0, cost: 0, count: 0 },
    '전국': { impressions: 0, clicks: 0, cost: 0, count: 0 },
    '기타': { impressions: 0, clicks: 0, cost: 0, count: 0 }
  };

  data.forEach(item => {
    let distance = '기타';
    const campaign = item.campaign || '';

    if (campaign.includes('1km')) distance = '1km';
    else if (campaign.includes('3km')) distance = '3km';
    else if (campaign.includes('5km')) distance = '5km';
    else if (campaign.includes('10km')) distance = '10km';
    else if (campaign.includes('전국')) distance = '전국';

    distances[distance].impressions += item.impressions;
    distances[distance].clicks += item.clicks;
    distances[distance].cost += item.cost;
    distances[distance].count++;
  });

  // CTR, CPC 계산
  Object.keys(distances).forEach(d => {
    const dist = distances[d];
    dist.ctr = dist.impressions > 0 ? (dist.clicks / dist.impressions * 100) : 0;
    dist.cpc = dist.clicks > 0 ? (dist.cost / dist.clicks) : 0;
  });

  return distances;
}

// Top/Bottom 아이템 추출 (클릭 10회 이상만)
function getTopItems(data, metric, count, descending) {
  if (!data || data.length === 0) return [];

  // 클릭 10회 이상, 키워드 유효한 것만 필터링
  const filtered = data.filter(item =>
    item.clicks >= 10 &&
    item.keyword &&
    item.keyword !== '-' &&
    item.keyword !== ''
  );

  const sorted = [...filtered].sort((a, b) => {
    const valA = a[metric] || 0;
    const valB = b[metric] || 0;
    return descending ? valB - valA : valA - valB;
  });

  return sorted.slice(0, count);
}

// 검색어 Top 추출 (CPC 낮은 순 = 효율 좋은 것, 클릭 10회 이상)
function getTopSearchTerms(data, count) {
  if (!data || data.length === 0) return [];

  const filtered = data.filter(item =>
    item.clicks >= 10 &&
    item.keyword &&
    item.keyword !== '-'
  );

  // CPC 낮은 순 = 효율 좋은 검색어
  const sorted = filtered.sort((a, b) => a.cpc - b.cpc);
  return sorted.slice(0, count);
}

// 검색어 Bottom 추출 (CPC 높은 순 = 비효율, 클릭 10회 이상)
function getBottomSearchTerms(data, count) {
  if (!data || data.length === 0) return [];

  const filtered = data.filter(item =>
    item.clicks >= 10 &&
    item.keyword &&
    item.keyword !== '-'
  );

  // CPC 높은 순 = 비효율 검색어
  const sorted = filtered.sort((a, b) => b.cpc - a.cpc);
  return sorted.slice(0, count);
}

// 월별 비교
function compareMonths(decSummary, janSummary) {
  const calcChange = (dec, jan) => {
    if (dec === 0) return jan > 0 ? 100 : 0;
    return ((jan - dec) / dec * 100);
  };

  return {
    impressionsChange: calcChange(decSummary.totalImpressions, janSummary.totalImpressions),
    clicksChange: calcChange(decSummary.totalClicks, janSummary.totalClicks),
    costChange: calcChange(decSummary.totalCost, janSummary.totalCost),
    ctrChange: janSummary.avgCtr - decSummary.avgCtr,
    cpcChange: calcChange(decSummary.avgCpc, janSummary.avgCpc)
  };
}

// ============================================
// 5. 종합 제언 생성 (CPC 중심)
// ============================================
function generateRecommendations(analysis) {
  const recommendations = [];
  const latestMonth = analysis.latestMonth;
  const prevMonth = analysis.prevMonth;

  // 1. 캠페인유형별 제언 (CPC 기준)
  const latestTypes = analysis.byTypeByMonth[latestMonth];
  if (latestTypes) {
    const typeEntries = Object.entries(latestTypes)
      .filter(([k, v]) => v.clicks >= 10) // 클릭 10회 이상만
      .sort((a, b) => a[1].cpc - b[1].cpc); // CPC 낮은 순

    if (typeEntries.length >= 2) {
      const bestType = typeEntries[0]; // CPC 가장 낮은 유형
      const worstType = typeEntries[typeEntries.length - 1]; // CPC 가장 높은 유형

      if (worstType[1].cpc > bestType[1].cpc * 1.5) {
        recommendations.push({
          category: '예산 재배분',
          content: `${bestType[0]}의 CPC(${formatCurrency(bestType[1].cpc)})가 ${worstType[0]}(${formatCurrency(worstType[1].cpc)})보다 효율적입니다. ${bestType[0]} 예산 확대를 권고합니다.`
        });
      }
    }
  }

  // 2. 거리별 제언 (CPC 기준)
  const latestDist = analysis.byDistanceByMonth[latestMonth];
  if (latestDist) {
    const distEntries = Object.entries(latestDist)
      .filter(([k, v]) => v.clicks >= 10 && k !== '기타')
      .sort((a, b) => a[1].cpc - b[1].cpc); // CPC 낮은 순

    if (distEntries.length > 0) {
      const bestDist = distEntries[0];
      recommendations.push({
        category: '타겟팅 최적화',
        content: `${bestDist[0]} 반경 타겟팅의 CPC(${formatCurrency(bestDist[1].cpc)})가 가장 효율적입니다. 해당 반경 중심으로 예산 집중을 권고합니다.`
      });
    }
  }

  // 3. 고비용 비효율 검색어 제언 (CPC 높은 것)
  if (analysis.bottomSearchTerms && analysis.bottomSearchTerms.length > 0) {
    const highCpcTerms = analysis.bottomSearchTerms
      .filter(k => k.cost >= 20000) // 비용 2만원 이상
      .slice(0, 3);

    if (highCpcTerms.length > 0) {
      const keywordList = highCpcTerms.map(k => `${k.keyword}(CPC ${formatCurrency(k.cpc)})`).join(', ');
      recommendations.push({
        category: '비용 효율 점검',
        content: `CPC 높은 검색어: ${keywordList}. 입찰가 조정 또는 제외 검색어 등록을 검토하세요.`
      });
    }
  }

  // 4. 고효율 검색어 제언 (CPC 낮은 것)
  if (analysis.topSearchTerms && analysis.topSearchTerms.length > 0) {
    const lowCpcTerms = analysis.topSearchTerms.slice(0, 3);

    if (lowCpcTerms.length > 0) {
      const termList = lowCpcTerms.map(s => `${s.keyword}(CPC ${formatCurrency(s.cpc)})`).join(', ');
      recommendations.push({
        category: '강화 권고',
        content: `CPC 효율 좋은 검색어: ${termList}. 정확검색 키워드 추가 및 예산 확대를 권고합니다.`
      });
    }
  }

  // 5. 월별 CPC 변화 제언
  if (analysis.monthComparison && prevMonth) {
    const change = analysis.monthComparison;

    if (change.cpcChange > 20) {
      recommendations.push({
        category: 'CPC 상승 주의',
        content: `${latestMonth} CPC가 ${prevMonth} 대비 ${change.cpcChange.toFixed(1)}% 상승했습니다. 경쟁 심화 또는 품질지수 하락 가능성을 점검하세요.`
      });
    } else if (change.cpcChange < -10) {
      recommendations.push({
        category: 'CPC 개선',
        content: `${latestMonth} CPC가 ${prevMonth} 대비 ${Math.abs(change.cpcChange).toFixed(1)}% 하락했습니다. 긍정적인 변화입니다.`
      });
    }

    // 비용 대비 클릭 효율 체크
    if (change.costChange > 20 && change.clicksChange < 10) {
      recommendations.push({
        category: '비용 효율 주의',
        content: `비용은 ${change.costChange.toFixed(1)}% 증가했으나 클릭은 ${change.clicksChange.toFixed(1)}% 변화했습니다. 비용 대비 효율을 점검하세요.`
      });
    }
  }

  return recommendations;
}

// ============================================
// 6. 결과 출력
// ============================================
function writeToSummarySheet(ss, analysis, recommendations) {
  let sheet = ss.getSheetByName('종합');
  if (!sheet) {
    sheet = ss.insertSheet('종합');
  }

  // 시트 초기화
  sheet.clear();

  const output = [];
  const now = new Date();
  const dateStr = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  const months = analysis.months;
  const latestMonth = analysis.latestMonth;
  const prevMonth = analysis.prevMonth;

  // 제목
  output.push(['📊 섬고짚 네이버광고 종합 분석 리포트']);
  output.push(['분석일시: ' + dateStr]);
  output.push(['분석 대상: ' + months.join(', ')]);
  output.push(['']);

  // ═══════════════════════════════════════
  // 전체 요약 (모든 월 표시)
  // ═══════════════════════════════════════
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['📌 전체 요약 (월별 추이)']);
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['']);

  // 동적 헤더 생성
  const summaryHeader = ['지표', ...months];
  if (prevMonth) summaryHeader.push('증감(' + prevMonth + '→' + latestMonth + ')');
  output.push(summaryHeader);

  // 총 노출수
  const impressionsRow = ['총 노출수'];
  months.forEach(m => impressionsRow.push(formatNumber(analysis.summaryByMonth[m].totalImpressions)));
  if (analysis.monthComparison) impressionsRow.push(formatChange(analysis.monthComparison.impressionsChange));
  output.push(impressionsRow);

  // 총 클릭수
  const clicksRow = ['총 클릭수'];
  months.forEach(m => clicksRow.push(formatNumber(analysis.summaryByMonth[m].totalClicks)));
  if (analysis.monthComparison) clicksRow.push(formatChange(analysis.monthComparison.clicksChange));
  output.push(clicksRow);

  // 총 비용
  const costRow = ['총 비용'];
  months.forEach(m => costRow.push(formatCurrency(analysis.summaryByMonth[m].totalCost)));
  if (analysis.monthComparison) costRow.push(formatChange(analysis.monthComparison.costChange));
  output.push(costRow);

  // 평균 CTR
  const ctrRow = ['평균 CTR'];
  months.forEach(m => ctrRow.push(analysis.summaryByMonth[m].avgCtr.toFixed(2) + '%'));
  if (analysis.monthComparison) ctrRow.push(formatChangePoint(analysis.monthComparison.ctrChange));
  output.push(ctrRow);

  // 평균 CPC
  const cpcRow = ['평균 CPC'];
  months.forEach(m => cpcRow.push(formatCurrency(analysis.summaryByMonth[m].avgCpc)));
  if (analysis.monthComparison) cpcRow.push(formatChange(analysis.monthComparison.cpcChange));
  output.push(cpcRow);

  output.push(['']);

  // ═══════════════════════════════════════
  // 캠페인유형별 성과 (최신 월)
  // ═══════════════════════════════════════
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['📈 캠페인유형별 성과 (' + latestMonth + ')']);
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['']);

  output.push(['유형', '노출수', '클릭수', 'CTR', '비용', 'CPC']);
  const latestTypeData = analysis.byTypeByMonth[latestMonth];
  if (latestTypeData) {
    Object.entries(latestTypeData).forEach(([type, data]) => {
      output.push([
        type,
        formatNumber(data.impressions),
        formatNumber(data.clicks),
        data.ctr.toFixed(2) + '%',
        formatCurrency(data.cost),
        formatCurrency(data.cpc)
      ]);
    });
  }
  output.push(['']);

  // ═══════════════════════════════════════
  // 거리별 성과 (최신 월)
  // ═══════════════════════════════════════
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['🎯 타겟 반경별 성과 (' + latestMonth + ')']);
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['']);

  output.push(['반경', '노출수', '클릭수', 'CTR', '비용', 'CPC']);
  const latestDistData = analysis.byDistanceByMonth[latestMonth];
  if (latestDistData) {
    ['1km', '3km', '5km', '10km', '전국'].forEach(dist => {
      const data = latestDistData[dist];
      if (data && data.count > 0) {
        output.push([
          dist,
          formatNumber(data.impressions),
          formatNumber(data.clicks),
          data.ctr.toFixed(2) + '%',
          formatCurrency(data.cost),
          formatCurrency(data.cpc)
        ]);
      }
    });
  }
  output.push(['']);

  // ═══════════════════════════════════════
  // Top 10 검색어 (CPC 효율 좋은 순)
  // ═══════════════════════════════════════
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['⭐ Top 10 검색어 (CPC 효율 좋은 순, 클릭 10회↑, ' + latestMonth + ')']);
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['']);

  output.push(['순위', '검색어', '캠페인유형', '클릭수', 'CPC', 'CTR', '비용']);
  analysis.topSearchTerms.forEach((item, idx) => {
    output.push([
      idx + 1,
      item.keyword,
      item.type,
      formatNumber(item.clicks),
      formatCurrency(item.cpc),
      item.ctr.toFixed(2) + '%',
      formatCurrency(item.cost)
    ]);
  });
  output.push(['']);

  // ═══════════════════════════════════════
  // 비효율 검색어 (CPC 높은 순)
  // ═══════════════════════════════════════
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['⚠️ 비효율 검색어 (CPC 높은 순, 클릭 10회↑, ' + latestMonth + ')']);
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['']);

  output.push(['순위', '검색어', '캠페인유형', '클릭수', 'CPC', 'CTR', '비용']);
  analysis.bottomSearchTerms.forEach((item, idx) => {
    output.push([
      idx + 1,
      item.keyword,
      item.type,
      formatNumber(item.clicks),
      formatCurrency(item.cpc),
      item.ctr.toFixed(2) + '%',
      formatCurrency(item.cost)
    ]);
  });
  output.push(['']);

  // ═══════════════════════════════════════
  // 캠페인유형별 CPC 비교
  // ═══════════════════════════════════════
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['💰 캠페인유형별 CPC 효율 비교 (' + latestMonth + ')']);
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['']);

  output.push(['유형', '클릭수', 'CPC', '총비용', '효율등급']);
  if (latestTypeData) {
    const typesByCpc = Object.entries(latestTypeData)
      .filter(([k, v]) => v.clicks >= 10)
      .sort((a, b) => a[1].cpc - b[1].cpc);

    typesByCpc.forEach(([type, data], idx) => {
      const grade = idx === 0 ? '🟢 최고' : idx === typesByCpc.length - 1 ? '🔴 개선필요' : '🟡 보통';
      output.push([
        type,
        formatNumber(data.clicks),
        formatCurrency(data.cpc),
        formatCurrency(data.cost),
        grade
      ]);
    });
  }
  output.push(['']);

  // ═══════════════════════════════════════
  // 종합 제언
  // ═══════════════════════════════════════
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['💡 종합 제언']);
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['']);

  if (recommendations.length === 0) {
    output.push(['현재 특별한 제언 사항이 없습니다. 전반적으로 양호한 성과입니다.']);
  } else {
    recommendations.forEach((rec, idx) => {
      output.push([`${idx + 1}. [${rec.category}]`]);
      output.push([`   ${rec.content}`]);
      output.push(['']);
    });
  }

  // 데이터 출력
  sheet.getRange(1, 1, output.length, 7).setValues(
    output.map(row => {
      while (row.length < 7) row.push('');
      return row;
    })
  );

  // 서식 적용
  applyFormatting(sheet, output.length);
}

// 서식 적용
function applyFormatting(sheet, rowCount) {
  // 제목 스타일
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A2').setFontSize(10).setFontColor('#666666');

  // 구분선 스타일
  const separatorRows = [];
  const headerRows = [];

  for (let i = 1; i <= rowCount; i++) {
    const cell = sheet.getRange(i, 1);
    const value = cell.getValue();

    if (String(value).includes('═══')) {
      separatorRows.push(i);
      cell.setFontColor('#1a73e8');
    }

    if (String(value).includes('📌') || String(value).includes('📈') ||
        String(value).includes('🎯') || String(value).includes('⭐') ||
        String(value).includes('⚠️') || String(value).includes('💰') ||
        String(value).includes('🔍') || String(value).includes('💡')) {
      headerRows.push(i);
      sheet.getRange(i, 1).setFontSize(12).setFontWeight('bold');
    }
  }

  // 열 너비 조정
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 120);
}

// ============================================
// 유틸리티 함수
// ============================================
function formatNumber(num) {
  if (!num) return '0';
  return Math.round(num).toLocaleString('ko-KR');
}

function formatCurrency(num) {
  if (!num) return '₩0';
  return '₩' + Math.round(num).toLocaleString('ko-KR');
}

function formatChange(percent) {
  if (!percent) return '-';
  const sign = percent >= 0 ? '+' : '';
  return sign + percent.toFixed(1) + '%';
}

function formatChangePoint(point) {
  if (!point) return '-';
  const sign = point >= 0 ? '+' : '';
  return sign + point.toFixed(2) + '%p';
}

// 종합탭 초기화
function clearSummarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('종합');
  if (sheet) {
    sheet.clear();
    ss.toast('종합 탭이 초기화되었습니다.', '완료', 3);
  }
}

// 월별탭 초기화
function clearMonthlySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('월별');
  if (sheet) {
    sheet.clear();
    ss.toast('월별 탭이 초기화되었습니다.', '완료', 3);
  }
}

// ============================================
// 7. 월별 분석 결과 출력
// ============================================
function writeToMonthlySheet(ss, analysis) {
  let sheet = ss.getSheetByName('월별');
  if (!sheet) {
    sheet = ss.insertSheet('월별');
  }

  sheet.clear();

  const output = [];
  const now = new Date();
  const dateStr = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  const month = analysis.selectedMonth;
  const prevMonth = analysis.prevMonth;

  // 제목
  output.push(['📅 ' + month + ' 네이버광고 상세 분석 리포트']);
  output.push(['분석일시: ' + dateStr]);
  output.push(['']);

  // ═══════════════════════════════════════
  // 월간 요약
  // ═══════════════════════════════════════
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['📌 ' + month + ' 요약']);
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['']);

  if (prevMonth && analysis.monthComparison) {
    output.push(['지표', month, prevMonth, '증감']);
    output.push(['총 노출수', formatNumber(analysis.summary.totalImpressions), formatNumber(analysis.prevSummary.totalImpressions), formatChange(analysis.monthComparison.impressionsChange)]);
    output.push(['총 클릭수', formatNumber(analysis.summary.totalClicks), formatNumber(analysis.prevSummary.totalClicks), formatChange(analysis.monthComparison.clicksChange)]);
    output.push(['총 비용', formatCurrency(analysis.summary.totalCost), formatCurrency(analysis.prevSummary.totalCost), formatChange(analysis.monthComparison.costChange)]);
    output.push(['평균 CTR', analysis.summary.avgCtr.toFixed(2) + '%', analysis.prevSummary.avgCtr.toFixed(2) + '%', formatChangePoint(analysis.monthComparison.ctrChange)]);
    output.push(['평균 CPC', formatCurrency(analysis.summary.avgCpc), formatCurrency(analysis.prevSummary.avgCpc), formatChange(analysis.monthComparison.cpcChange)]);
  } else {
    output.push(['지표', month]);
    output.push(['총 노출수', formatNumber(analysis.summary.totalImpressions)]);
    output.push(['총 클릭수', formatNumber(analysis.summary.totalClicks)]);
    output.push(['총 비용', formatCurrency(analysis.summary.totalCost)]);
    output.push(['평균 CTR', analysis.summary.avgCtr.toFixed(2) + '%']);
    output.push(['평균 CPC', formatCurrency(analysis.summary.avgCpc)]);
  }
  output.push(['']);

  // ═══════════════════════════════════════
  // 캠페인유형별 성과
  // ═══════════════════════════════════════
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['📈 캠페인유형별 성과']);
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['']);

  output.push(['유형', '노출수', '클릭수', 'CTR', '비용', 'CPC']);
  if (analysis.byType) {
    Object.entries(analysis.byType).forEach(([type, data]) => {
      output.push([
        type,
        formatNumber(data.impressions),
        formatNumber(data.clicks),
        data.ctr.toFixed(2) + '%',
        formatCurrency(data.cost),
        formatCurrency(data.cpc)
      ]);
    });
  }
  output.push(['']);

  // ═══════════════════════════════════════
  // 타겟 반경별 성과
  // ═══════════════════════════════════════
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['🎯 타겟 반경별 성과']);
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['']);

  output.push(['반경', '노출수', '클릭수', 'CTR', '비용', 'CPC']);
  if (analysis.byDistance) {
    ['1km', '3km', '5km', '10km', '전국'].forEach(dist => {
      const data = analysis.byDistance[dist];
      if (data && data.count > 0) {
        output.push([
          dist,
          formatNumber(data.impressions),
          formatNumber(data.clicks),
          data.ctr.toFixed(2) + '%',
          formatCurrency(data.cost),
          formatCurrency(data.cpc)
        ]);
      }
    });
  }
  output.push(['']);

  // ═══════════════════════════════════════
  // Top 10 검색어 (CPC 효율 좋은 순)
  // ═══════════════════════════════════════
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['⭐ Top 10 검색어 (CPC 효율 좋은 순, 클릭 10회↑)']);
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['']);

  output.push(['순위', '검색어', '캠페인유형', '클릭수', 'CPC', 'CTR', '비용']);
  analysis.topSearchTerms.forEach((item, idx) => {
    output.push([
      idx + 1,
      item.keyword,
      item.type,
      formatNumber(item.clicks),
      formatCurrency(item.cpc),
      item.ctr.toFixed(2) + '%',
      formatCurrency(item.cost)
    ]);
  });
  output.push(['']);

  // ═══════════════════════════════════════
  // 비효율 검색어 (CPC 높은 순)
  // ═══════════════════════════════════════
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['⚠️ 비효율 검색어 (CPC 높은 순, 클릭 10회↑)']);
  output.push(['═══════════════════════════════════════════════════════════════']);
  output.push(['']);

  output.push(['순위', '검색어', '캠페인유형', '클릭수', 'CPC', 'CTR', '비용']);
  analysis.bottomSearchTerms.forEach((item, idx) => {
    output.push([
      idx + 1,
      item.keyword,
      item.type,
      formatNumber(item.clicks),
      formatCurrency(item.cpc),
      item.ctr.toFixed(2) + '%',
      formatCurrency(item.cost)
    ]);
  });
  output.push(['']);

  // 데이터 출력
  const maxCols = 7;
  sheet.getRange(1, 1, output.length, maxCols).setValues(
    output.map(row => {
      while (row.length < maxCols) row.push('');
      return row;
    })
  );

  // 서식 적용
  applyFormatting(sheet, output.length);
}
