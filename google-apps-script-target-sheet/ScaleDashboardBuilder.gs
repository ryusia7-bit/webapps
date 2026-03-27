/** @OnlyCurrentDoc */

/**
 * 대상자 검색형 척도대시보드를 재구축합니다.
 * `척도검사기록 -> 실무자보기 -> 척도대시보드` 흐름을 기준으로
 * 대상자 1명의 날짜별 척도 변화와 최근 검사 이력을 한 화면에서 보도록 구성합니다.
 *
 * @returns {{dashboardSheetName:string, selectedClient:string, chartCount:number}}
 */
function buildDashboard() {
  console.time("buildDashboard");
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = getSpreadsheet_();
    const dashboardSheet =
      spreadsheet.getSheetByName(getScaleScreeningDashboardSheetName_()) ||
      spreadsheet.insertSheet(getScaleScreeningDashboardSheetName_());

    const summary = buildDashboard_(dashboardSheet);
    SpreadsheetApp.flush();
    console.log(JSON.stringify(summary));
    return summary;
  } catch (error) {
    console.error("buildDashboard failed", {
      message: error && error.message ? error.message : String(error),
      stack: error && error.stack ? error.stack : ""
    });
    throw error;
  } finally {
    lock.releaseLock();
    console.timeEnd("buildDashboard");
  }
}

/**
 * 워크스페이스 재구축 중 동일한 대시보드 레이아웃을 재사용하기 위한 내부 래퍼입니다.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} existingSheet
 * @returns {{dashboardSheetName:string, selectedClient:string, chartCount:number}}
 */
function buildDashboard_(existingSheet) {
  const spreadsheet = getSpreadsheet_();
  const dashboardSheet =
    existingSheet ||
    spreadsheet.getSheetByName(getScaleScreeningDashboardSheetName_()) ||
    spreadsheet.insertSheet(getScaleScreeningDashboardSheetName_());

  buildScaleScreeningDashboardSheet_(dashboardSheet);
  repairScaleDashboardSummaryCards_(dashboardSheet);

  return {
    dashboardSheetName: dashboardSheet.getName(),
    selectedClient: normalizeText_(dashboardSheet.getRange(SCALE_SCREENING_SYNC_CONFIG.dashboard.clientNameCell).getDisplayValue()),
    chartCount: dashboardSheet.getCharts().length
  };
}

/**
 * 대시보드 구성 확인용 스냅샷입니다.
 *
 * @returns {{
 *   selectedClient:string,
 *   chartCount:number,
 *   longitudinalRowCount:number,
 *   firstLongitudinalDate:string,
 *   firstLongitudinalScale:string,
 *   firstDeltaText:string
 * }}
 */
function getDashboardSnapshot() {
  const sheet = getSpreadsheet_().getSheetByName(getScaleScreeningDashboardSheetName_());
  if (!sheet) {
    throw new Error("척도대시보드 시트를 찾을 수 없습니다.");
  }

  const dashboardConfig = SCALE_SCREENING_SYNC_CONFIG.dashboard;
  const longitudinalRows = sheet.getRange(
    dashboardConfig.detailStartRow,
    1,
    Math.max(1, 220 - dashboardConfig.detailStartRow),
    10
  ).getDisplayValues().filter(function(row) {
    return row.some(Boolean);
  });

  return {
    selectedClient: normalizeText_(sheet.getRange(dashboardConfig.clientNameCell).getDisplayValue()),
    chartCount: sheet.getCharts().length,
    longitudinalRowCount: longitudinalRows.length,
    firstLongitudinalDate: longitudinalRows.length ? normalizeText_(longitudinalRows[0][0]) : "",
    firstLongitudinalScale: longitudinalRows.length ? normalizeText_(longitudinalRows[0][1]) : "",
    firstDeltaText: longitudinalRows.length ? normalizeText_(longitudinalRows[0][6]) : ""
  };
}

function repairScaleDashboardSummaryCards_(sheet) {
  const dashboardConfig = SCALE_SCREENING_SYNC_CONFIG.dashboard;
  const startRow = dashboardConfig.detailStartRow;
  const rowCount = Math.max(1, sheet.getLastRow() - startRow + 1);
  const rows = sheet.getRange(startRow, 1, rowCount, 10).getDisplayValues().filter(function(row) {
    return normalizeText_(row[0]) && normalizeText_(row[0]) !== "검색 결과가 없습니다.";
  });

  const inspectionCount = rows.length;
  const latestSessionDate = rows.length
    ? rows.map(function(row) { return normalizeText_(row[0]); }).sort().slice(-1)[0]
    : "-";
  const highRiskCount = rows.filter(function(row) {
    return /고|severe/i.test(normalizeText_(row[7]));
  }).length;
  const latestRow = rows.length
    ? rows
        .slice()
        .sort(function(a, b) {
          const dateA = formatScaleDashboardDateDisplay_(a[0]) || "";
          const dateB = formatScaleDashboardDateDisplay_(b[0]) || "";
          if (dateA === dateB) {
            return normalizeText_(a[1]).localeCompare(normalizeText_(b[1]), "ko");
          }
          return dateA.localeCompare(dateB, "ko");
        })
        .slice(-1)[0]
    : null;
  const latestWorker = latestRow ? normalizeText_(latestRow[8]) : "";
  sheet.getRange("G8:I10").setValue("검사 건수\n" + (inspectionCount ? inspectionCount + "건" : "-"));
  sheet.getRange("J8:L10").setValue("최근 검사일\n" + (inspectionCount ? latestSessionDate : "-"));
  sheet.getRange("G11:I13").setValue("고위험 건수\n" + (highRiskCount ? highRiskCount + "건" : "-"));
  sheet.getRange("J11:L13").setValue("최근 담당자\n" + (latestWorker || "-"));
}
