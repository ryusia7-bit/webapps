/** @OnlyCurrentDoc */

/**
 * @typedef {Object} ScaleDashboardRecord
 * @property {string} recordId
 * @property {Date|null} sessionDate
 * @property {string} sessionDateText
 * @property {string} questionnaireId
 * @property {string} questionnaireTitle
 * @property {string} scoreText
 * @property {number|null} normalizedScore
 * @property {string} bandText
 * @property {string} workerName
 * @property {string} clientLabel
 * @property {string} gender
 * @property {string} ageGroup
 * @property {string} flags
 * @property {boolean} isHighRisk
 */

/**
 * @typedef {Object} ScaleDashboardModel
 * @property {{totalRecords:number,highRiskCount:number,uniqueClientCount:number,phqAverage:number}} kpi
 * @property {Array<Array<string|number>>} monthlyRows
 * @property {Array<Array<string|number>>} riskRows
 * @property {Array<Array<string|number>>} workerRows
 * @property {Array<Array<string|number>>} recentHighRiskRows
 */

const SCALE_DASHBOARD_BUILD_CONFIG = {
  sourceSheetName: "척도검사기록",
  dashboardSheetName: "척도대시보드",
  sourceHeaders: {
    recordId: "record_id",
    sessionDate: "session_date",
    questionnaireId: "questionnaire_id",
    questionnaireTitle: "questionnaire_title",
    scoreText: "score_text",
    normalizedScore: "normalized_score",
    bandText: "band_text",
    workerName: "worker_name",
    clientLabel: "client_label",
    gender: "gender",
    ageGroup: "age_group",
    flags: "flags"
  },
  trackedQuestionnaires: [
    { id: "phq-9", label: "PHQ-9", color: "#5B8FF9" },
    { id: "gad-7", label: "GAD-7", color: "#F6BD16" },
    { id: "mkpq-16", label: "MKPQ-16", color: "#9270CA" }
  ],
  cardColors: [
    { background: "#E8F1FF", border: "#B7CCFF", font: "#1D4ED8" },
    { background: "#FFF1E0", border: "#F7C992", font: "#B45309" },
    { background: "#FDECEC", border: "#F5B7B1", font: "#C0392B" },
    { background: "#EAF7EE", border: "#B8E0C2", font: "#1E8449" }
  ],
  chartColors: ["#5B8FF9", "#F6BD16", "#9270CA", "#EF4444", "#10B981", "#F97316"]
};

/**
 * 척도대시보드를 요청 사양대로 완전히 재구축합니다.
 * 원본 데이터는 `척도검사기록` 시트 1행 헤더를 기준으로 읽고,
 * 집계는 메모리에서 처리한 뒤 한 번에 출력합니다.
 *
 * @returns {{dashboardSheetName:string,totalRecords:number,highRiskCount:number,uniqueClientCount:number,phqAverage:number,chartCount:number}}
 */
function buildDashboard() {
  console.time("buildDashboard");
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = getSpreadsheet_();
    const sourceSheet = spreadsheet.getSheetByName(SCALE_DASHBOARD_BUILD_CONFIG.sourceSheetName);
    if (!sourceSheet) {
      throw new Error("원본 데이터 시트 '척도검사기록'을 찾을 수 없습니다.");
    }

    const dashboardSheet =
      spreadsheet.getSheetByName(SCALE_DASHBOARD_BUILD_CONFIG.dashboardSheetName) ||
      spreadsheet.insertSheet(SCALE_DASHBOARD_BUILD_CONFIG.dashboardSheetName);

    const records = readScaleDashboardRecords_(sourceSheet);
    const model = buildScaleDashboardModel_(records);
    renderScaleDashboardSheet_(dashboardSheet, model);
    SpreadsheetApp.flush();

    const summary = {
      dashboardSheetName: dashboardSheet.getName(),
      totalRecords: model.kpi.totalRecords,
      highRiskCount: model.kpi.highRiskCount,
      uniqueClientCount: model.kpi.uniqueClientCount,
      phqAverage: Number(model.kpi.phqAverage.toFixed(1)),
      chartCount: dashboardSheet.getCharts().length
    };

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
 * 메뉴/워크스페이스 재구축 흐름에서도 같은 대시보드 레이아웃을 쓰기 위한 내부 래퍼입니다.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet=} existingSheet
 * @returns {{dashboardSheetName:string,totalRecords:number,highRiskCount:number,uniqueClientCount:number,phqAverage:number,chartCount:number}}
 */
function buildDashboard_(existingSheet) {
  const spreadsheet = getSpreadsheet_();
  const dashboardSheet =
    existingSheet ||
    spreadsheet.getSheetByName(SCALE_DASHBOARD_BUILD_CONFIG.dashboardSheetName) ||
    spreadsheet.insertSheet(SCALE_DASHBOARD_BUILD_CONFIG.dashboardSheetName);

  const sourceSheet = spreadsheet.getSheetByName(SCALE_DASHBOARD_BUILD_CONFIG.sourceSheetName);
  if (!sourceSheet) {
    throw new Error("원본 데이터 시트 '척도검사기록'을 찾을 수 없습니다.");
  }

  const records = readScaleDashboardRecords_(sourceSheet);
  const model = buildScaleDashboardModel_(records);
  renderScaleDashboardSheet_(dashboardSheet, model);
  SpreadsheetApp.flush();

  return {
    dashboardSheetName: dashboardSheet.getName(),
    totalRecords: model.kpi.totalRecords,
    highRiskCount: model.kpi.highRiskCount,
    uniqueClientCount: model.kpi.uniqueClientCount,
    phqAverage: Number(model.kpi.phqAverage.toFixed(1)),
    chartCount: dashboardSheet.getCharts().length
  };
}

/**
 * 실행 결과를 빠르게 확인하기 위한 점검용 함수입니다.
 *
 * @returns {{title:string,kpis:Array<string>,monthlyRows:number,riskRows:number,workerRows:number,recentRows:number,chartCount:number}}
 */
function getDashboardSnapshot() {
  const sheet = getSpreadsheet_().getSheetByName(SCALE_DASHBOARD_BUILD_CONFIG.dashboardSheetName);
  if (!sheet) {
    throw new Error("척도대시보드 시트를 찾을 수 없습니다.");
  }

  return {
    title: normalizeText_(sheet.getRange("A8").getDisplayValue()),
    kpis: [
      normalizeText_(sheet.getRange("A1").getDisplayValue()),
      normalizeText_(sheet.getRange("D1").getDisplayValue()),
      normalizeText_(sheet.getRange("G1").getDisplayValue()),
      normalizeText_(sheet.getRange("J1").getDisplayValue())
    ],
    monthlyRows: Math.max(sheet.getRange("A10:E21").getDisplayValues().filter((row) => row.some(Boolean)).length, 0),
    riskRows: Math.max(sheet.getRange("G10:J12").getDisplayValues().filter((row) => row.some(Boolean)).length, 0),
    workerRows: Math.max(sheet.getRange("L10:N40").getDisplayValues().filter((row) => row.some(Boolean)).length, 0),
    recentRows: Math.max(sheet.getRange("A32:F51").getDisplayValues().filter((row) => row.some(Boolean)).length, 0),
    chartCount: sheet.getCharts().length
  };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sourceSheet
 * @returns {ScaleDashboardRecord[]}
 */
function readScaleDashboardRecords_(sourceSheet) {
  console.time("readScaleDashboardRecords");
  const lastRow = sourceSheet.getLastRow();
  const lastColumn = Math.max(sourceSheet.getLastColumn(), 28);
  if (lastRow < 2 || lastColumn < 1) {
    console.timeEnd("readScaleDashboardRecords");
    return [];
  }

  const values = sourceSheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = values[0].map((header) => normalizeText_(header).toLowerCase());
  const headerMap = new Map(headers.map((header, index) => [header, index]));
  const indexes = getScaleDashboardHeaderIndexes_(headerMap);

  const records = values.slice(1).reduce((list, row) => {
    const recordId = normalizeText_(row[indexes.recordId]);
    const clientLabel = normalizeText_(row[indexes.clientLabel]);
    const questionnaireId = normalizeText_(row[indexes.questionnaireId]).toLowerCase();

    if (!recordId && !clientLabel && !questionnaireId) {
      return list;
    }

    const sessionDate = toDashboardDate_(row[indexes.sessionDate]);
    const bandText = normalizeText_(row[indexes.bandText]);

    list.push({
      recordId: recordId || `record-${list.length + 1}`,
      sessionDate,
      sessionDateText: sessionDate ? formatDashboardDate_(sessionDate) : "",
      questionnaireId,
      questionnaireTitle: normalizeText_(row[indexes.questionnaireTitle]),
      scoreText: normalizeText_(row[indexes.scoreText]),
      normalizedScore: toDashboardNumber_(row[indexes.normalizedScore]),
      bandText,
      workerName: normalizeText_(row[indexes.workerName]) || "(미지정)",
      clientLabel: clientLabel || "(이름없음)",
      gender: normalizeText_(row[indexes.gender]),
      ageGroup: normalizeText_(row[indexes.ageGroup]),
      flags: normalizeText_(row[indexes.flags]),
      isHighRisk: isDashboardHighRisk_(bandText)
    });
    return list;
  }, []);

  console.timeEnd("readScaleDashboardRecords");
  return records;
}

/**
 * @param {Map<string, number>} headerMap
 * @returns {{recordId:number,sessionDate:number,questionnaireId:number,questionnaireTitle:number,scoreText:number,normalizedScore:number,bandText:number,workerName:number,clientLabel:number,gender:number,ageGroup:number,flags:number}}
 */
function getScaleDashboardHeaderIndexes_(headerMap) {
  const config = SCALE_DASHBOARD_BUILD_CONFIG.sourceHeaders;
  const indexes = {
    recordId: headerMap.get(config.recordId),
    sessionDate: headerMap.get(config.sessionDate),
    questionnaireId: headerMap.get(config.questionnaireId),
    questionnaireTitle: headerMap.get(config.questionnaireTitle),
    scoreText: headerMap.get(config.scoreText),
    normalizedScore: headerMap.get(config.normalizedScore),
    bandText: headerMap.get(config.bandText),
    workerName: headerMap.get(config.workerName),
    clientLabel: headerMap.get(config.clientLabel),
    gender: headerMap.get(config.gender),
    ageGroup: headerMap.get(config.ageGroup),
    flags: headerMap.get(config.flags)
  };

  const missingHeaders = Object.entries(config)
    .filter(([, headerName]) => headerMap.get(headerName) === undefined)
    .map(([, headerName]) => headerName);

  if (missingHeaders.length) {
    throw new Error("원본 데이터 시트에 필요한 헤더가 없습니다: " + missingHeaders.join(", "));
  }

  return indexes;
}

/**
 * @param {ScaleDashboardRecord[]} records
 * @returns {ScaleDashboardModel}
 */
function buildScaleDashboardModel_(records) {
  console.time("buildScaleDashboardModel");
  const highRiskRegex = /고|severe/i;
  const uniqueClients = new Set();
  const workerMap = new Map();
  const monthlyMap = new Map();
  const riskMap = new Map(
    SCALE_DASHBOARD_BUILD_CONFIG.trackedQuestionnaires.map((item) => [item.id, { label: item.label, total: 0, high: 0 }])
  );

  let phqScoreSum = 0;
  let phqScoreCount = 0;
  let highRiskCount = 0;
  let maxDate = null;

  records.forEach((record) => {
    if (record.clientLabel && record.clientLabel !== "(이름없음)") {
      uniqueClients.add(record.clientLabel);
    }

    if (record.isHighRisk || highRiskRegex.test(record.bandText)) {
      highRiskCount += 1;
    }

    if (record.questionnaireId === "phq-9" && Number.isFinite(record.normalizedScore)) {
      phqScoreSum += Number(record.normalizedScore);
      phqScoreCount += 1;
    }

    if (record.sessionDate && (!maxDate || record.sessionDate.getTime() > maxDate.getTime())) {
      maxDate = record.sessionDate;
    }

    const workerBucket = workerMap.get(record.workerName) || { name: record.workerName, total: 0, high: 0 };
    workerBucket.total += 1;
    workerBucket.high += record.isHighRisk ? 1 : 0;
    workerMap.set(record.workerName, workerBucket);

    const riskBucket = riskMap.get(record.questionnaireId);
    if (riskBucket) {
      riskBucket.total += 1;
      riskBucket.high += record.isHighRisk ? 1 : 0;
    }
  });

  const monthKeys = buildRecentMonthKeys_(maxDate || new Date(), 12);
  monthKeys.forEach((monthKey) => {
    monthlyMap.set(monthKey, { phq: 0, gad: 0, mkpq: 0, total: 0 });
  });

  records.forEach((record) => {
    if (!record.sessionDate) {
      return;
    }

    const monthKey = formatDashboardMonth_(record.sessionDate);
    const bucket = monthlyMap.get(monthKey);
    if (!bucket) {
      return;
    }

    bucket.total += 1;
    if (record.questionnaireId === "phq-9") {
      bucket.phq += 1;
    } else if (record.questionnaireId === "gad-7") {
      bucket.gad += 1;
    } else if (record.questionnaireId === "mkpq-16") {
      bucket.mkpq += 1;
    }
  });

  const monthlyRows = monthKeys.map((monthKey) => {
    const bucket = monthlyMap.get(monthKey) || { phq: 0, gad: 0, mkpq: 0, total: 0 };
    return [monthKey, bucket.phq, bucket.gad, bucket.mkpq, bucket.total];
  });

  const riskRows = SCALE_DASHBOARD_BUILD_CONFIG.trackedQuestionnaires.map((item) => {
    const bucket = riskMap.get(item.id) || { total: 0, high: 0 };
    const ratio = bucket.total > 0 ? bucket.high / bucket.total : 0;
    return [item.label, bucket.total, bucket.high, ratio];
  });

  const workerRows = Array.from(workerMap.values())
    .sort((a, b) => {
      if (b.total !== a.total) {
        return b.total - a.total;
      }
      return a.name.localeCompare(b.name, "ko");
    })
    .map((item) => [item.name, item.total, item.high]);

  const recentHighRiskRows = records
    .filter((record) => record.isHighRisk)
    .sort((a, b) => {
      const timeA = a.sessionDate ? a.sessionDate.getTime() : 0;
      const timeB = b.sessionDate ? b.sessionDate.getTime() : 0;
      return timeB - timeA;
    })
    .slice(0, 20)
    .map((record) => [
      record.sessionDateText,
      record.clientLabel,
      record.questionnaireTitle || record.questionnaireId,
      record.scoreText || (Number.isFinite(record.normalizedScore) ? String(record.normalizedScore) : ""),
      record.bandText,
      record.workerName
    ]);

  const model = {
    kpi: {
      totalRecords: records.length,
      highRiskCount,
      uniqueClientCount: uniqueClients.size,
      phqAverage: phqScoreCount > 0 ? phqScoreSum / phqScoreCount : 0
    },
    monthlyRows,
    riskRows,
    workerRows: workerRows.length ? workerRows : [["(데이터 없음)", 0, 0]],
    recentHighRiskRows: recentHighRiskRows.length
      ? recentHighRiskRows
      : [["", "데이터 없음", "", "", "", ""]]
  };

  console.timeEnd("buildScaleDashboardModel");
  return model;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {ScaleDashboardModel} model
 * @returns {void}
 */
function renderScaleDashboardSheet_(sheet, model) {
  console.time("renderScaleDashboardSheet");
  sheet.getCharts().forEach((chart) => sheet.removeChart(chart));
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  sheet.clearConditionalFormatRules();
  ensureSheetSize_(sheet, 70, 24);
  sheet.showColumns(1, 24);
  sheet.showRows(1, 70);
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(9);
  sheet.setFrozenColumns(0);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setFontFamily("Malgun Gothic");

  renderScaleDashboardCards_(sheet, model.kpi);
  renderScaleDashboardMonthlySection_(sheet, model.monthlyRows);
  renderScaleDashboardRiskSection_(sheet, model.riskRows);
  renderScaleDashboardWorkerSection_(sheet, model.workerRows);
  renderScaleDashboardRecentSection_(sheet, model.recentHighRiskRows);
  styleScaleDashboardLayout_(sheet, model);
  insertScaleDashboardCharts_(sheet, model);
  console.timeEnd("renderScaleDashboardSheet");
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {{totalRecords:number,highRiskCount:number,uniqueClientCount:number,phqAverage:number}} kpi
 */
function renderScaleDashboardCards_(sheet, kpi) {
  const cardSpecs = [
    { label: "전체 검사 건수", value: `${kpi.totalRecords}건`, labelRange: "A1:C1", valueRange: "A2:C4", color: SCALE_DASHBOARD_BUILD_CONFIG.cardColors[0] },
    { label: "고위험 건수", value: `${kpi.highRiskCount}건`, labelRange: "D1:F1", valueRange: "D2:F4", color: SCALE_DASHBOARD_BUILD_CONFIG.cardColors[1] },
    { label: "전체 대상자 수", value: `${kpi.uniqueClientCount}명`, labelRange: "G1:I1", valueRange: "G2:I4", color: SCALE_DASHBOARD_BUILD_CONFIG.cardColors[2] },
    { label: "PHQ-9 평균 점수", value: `${kpi.phqAverage.toFixed(1)}점`, labelRange: "J1:L1", valueRange: "J2:L4", color: SCALE_DASHBOARD_BUILD_CONFIG.cardColors[3] }
  ];

  cardSpecs.forEach((card) => {
    sheet.getRange(card.labelRange).merge().setValue(card.label);
    sheet.getRange(card.valueRange).merge().setValue(card.value);
    [card.labelRange, card.valueRange].forEach((rangeA1) => {
      sheet
        .getRange(rangeA1)
        .setBackground(card.color.background)
        .setBorder(true, true, true, true, true, true, card.color.border, SpreadsheetApp.BorderStyle.SOLID);
    });
    sheet.getRange(card.labelRange).setFontWeight("bold").setFontSize(12).setFontColor(card.color.font).setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange(card.valueRange).setFontWeight("bold").setFontSize(22).setFontColor(card.color.font).setHorizontalAlignment("center").setVerticalAlignment("middle");
  });
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<Array<string|number>>} rows
 */
function renderScaleDashboardMonthlySection_(sheet, rows) {
  sheet.getRange("A8:E8").merge().setValue("월별 검사 건수");
  sheet.getRange("A9:E9").setValues([["월", "PHQ-9", "GAD-7", "MKPQ-16", "합계"]]);
  sheet.getRange(10, 1, rows.length, 5).setValues(rows);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<Array<string|number>>} rows
 */
function renderScaleDashboardRiskSection_(sheet, rows) {
  sheet.getRange("G8:J8").merge().setValue("고위험군 현황");
  sheet.getRange("G9:J9").setValues([["척도", "전체건수", "고위험건수", "고위험비율"]]);
  sheet.getRange(10, 7, rows.length, 4).setValues(rows);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<Array<string|number>>} rows
 */
function renderScaleDashboardWorkerSection_(sheet, rows) {
  sheet.getRange("L8:N8").merge().setValue("담당자별 실적");
  sheet.getRange("L9:N9").setValues([["담당자", "전체 검사 건수", "고위험 발굴 건수"]]);
  sheet.getRange(10, 12, rows.length, 3).setValues(rows);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<Array<string|number>>} rows
 */
function renderScaleDashboardRecentSection_(sheet, rows) {
  sheet.getRange("A30:F30").merge().setValue("고위험군 대상자 목록 (최근 20건)");
  sheet.getRange("A31:F31").setValues([["검사일", "대상자명", "척도", "점수", "위험구간", "담당자"]]);
  sheet.getRange(32, 1, rows.length, 6).setValues(rows);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {ScaleDashboardModel} model
 */
function styleScaleDashboardLayout_(sheet, model) {
  const titleColor = "#173F5F";
  const titleTextColor = "#FFFFFF";
  const sectionRanges = ["A8:E8", "G8:J8", "L8:N8", "A30:F30"];
  const headerRanges = ["A9:E9", "G9:J9", "L9:N9", "A31:F31"];
  const workerEndRow = 9 + Math.max(model.workerRows.length, 1);
  const recentEndRow = 31 + Math.max(model.recentHighRiskRows.length, 1);

  sheet.getRangeList(sectionRanges.concat(headerRanges))
    .setBackground(titleColor)
    .setFontColor(titleTextColor)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.getRange("A8:N8").setFontSize(12);
  sheet.getRange("A9:N9").setFontSize(10);
  sheet.getRange("A30:F31").setFontSize(10);

  sheet.getRange("A10:E21").setHorizontalAlignment("center");
  sheet.getRange("H10:J12").setHorizontalAlignment("center");
  sheet.getRange(`M10:N${workerEndRow}`).setHorizontalAlignment("center");
  sheet.getRange(`A32:F${recentEndRow}`).setVerticalAlignment("middle");

  sheet.getRange("J2:L4").setNumberFormat("0.0점");
  sheet.getRange("J10:J12").setNumberFormat("0.0%");
  sheet.getRange("A10:A21").setNumberFormat("@");
  sheet.getRange(`A32:A${recentEndRow}`).setNumberFormat("yyyy-mm-dd");

  sheet.getRange(`A32:F${recentEndRow}`)
    .setBackground("#FFF5F5")
    .setFontColor("#7F1D1D");

  ["A8:E21", "G8:J12", `L8:N${workerEndRow}`, `A30:F${recentEndRow}`].forEach((rangeA1) => {
    sheet
      .getRange(rangeA1)
      .setBorder(true, true, true, true, true, true, "#D1D5DB", SpreadsheetApp.BorderStyle.SOLID);
  });

  const widths = [
    [1, 92], [2, 120], [3, 110], [4, 90], [5, 82],
    [7, 110], [8, 82], [9, 92], [10, 96],
    [12, 130], [13, 92], [14, 108]
  ];

  widths.forEach(([column, width]) => sheet.setColumnWidth(column, width));
  sheet.setColumnWidth(15, 16);
  sheet.setColumnWidth(16, 16);
  sheet.setColumnWidth(17, 16);
  sheet.setColumnWidth(18, 16);
  sheet.setColumnWidth(19, 16);
  sheet.setColumnWidth(20, 16);
  sheet.setColumnWidth(21, 16);
  sheet.setColumnWidth(22, 16);
  sheet.setColumnWidth(23, 16);
  sheet.setColumnWidth(24, 16);

  sheet.setRowHeights(1, 4, 28);
  sheet.setRowHeights(8, 2, 24);
  sheet.setRowHeights(30, 2, 24);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {ScaleDashboardModel} model
 */
function insertScaleDashboardCharts_(sheet, model) {
  const monthlyChartRowCount = model.monthlyRows.length + 1;
  const riskChartRowCount = model.riskRows.length + 1;
  const workerChartRowCount = Math.max(model.workerRows.length + 1, 2);

  const monthlyChart = sheet.newChart()
    .asColumnChart()
    .addRange(sheet.getRange(9, 1, monthlyChartRowCount, 4))
    .setNumHeaders(1)
    .setPosition(22, 1, 0, 0)
    .setOption("title", "월별 검사 건수")
    .setOption("legend.position", "top")
    .setOption("backgroundColor", "#FFFFFF")
    .setOption("isStacked", true)
    .setOption("height", 170)
    .setOption("width", 540)
    .setOption("chartArea", { left: 60, top: 40, width: "72%", height: "58%" })
    .setOption("hAxis", { slantedText: true, slantedTextAngle: 45 })
    .setOption("vAxis", { minValue: 0 })
    .setOption("series", {
      0: { color: SCALE_DASHBOARD_BUILD_CONFIG.trackedQuestionnaires[0].color },
      1: { color: SCALE_DASHBOARD_BUILD_CONFIG.trackedQuestionnaires[1].color },
      2: { color: SCALE_DASHBOARD_BUILD_CONFIG.trackedQuestionnaires[2].color }
    })
    .build();

  const riskChart = sheet.newChart()
    .asPieChart()
    .addRange(sheet.getRange(9, 7, riskChartRowCount, 1))
    .addRange(sheet.getRange(9, 9, riskChartRowCount, 1))
    .setNumHeaders(1)
    .setPosition(14, 7, 0, 0)
    .setOption("title", "척도별 고위험군 비율")
    .setOption("legend.position", "right")
    .setOption("backgroundColor", "#FFFFFF")
    .setOption("pieHole", 0.35)
    .setOption("height", 210)
    .setOption("width", 380)
    .setOption("series", {
      0: { color: SCALE_DASHBOARD_BUILD_CONFIG.trackedQuestionnaires[0].color },
      1: { color: SCALE_DASHBOARD_BUILD_CONFIG.trackedQuestionnaires[1].color },
      2: { color: SCALE_DASHBOARD_BUILD_CONFIG.trackedQuestionnaires[2].color }
    })
    .build();

  const workerChart = sheet.newChart()
    .asBarChart()
    .addRange(sheet.getRange(9, 12, workerChartRowCount, 3))
    .setNumHeaders(1)
    .setPosition(11, 15, 0, 0)
    .setOption("title", "담당자별 검사 실적")
    .setOption("legend.position", "top")
    .setOption("backgroundColor", "#FFFFFF")
    .setOption("height", 320)
    .setOption("width", 500)
    .setOption("hAxis", { minValue: 0 })
    .setOption("series", {
      0: { color: "#5B8FF9" },
      1: { color: "#EF4444" }
    })
    .build();

  sheet.insertChart(monthlyChart);
  sheet.insertChart(riskChart);
  sheet.insertChart(workerChart);
}

/**
 * @param {Date} endDate
 * @param {number} count
 * @returns {string[]}
 */
function buildRecentMonthKeys_(endDate, count) {
  const months = [];
  const cursor = new Date(endDate.getFullYear(), endDate.getMonth(), 1);

  for (let index = count - 1; index >= 0; index -= 1) {
    const date = new Date(cursor.getFullYear(), cursor.getMonth() - index, 1);
    months.push(formatDashboardMonth_(date));
  }

  return months;
}

/**
 * @param {Date} date
 * @returns {string}
 */
function formatDashboardMonth_(date) {
  return Utilities.formatDate(date, APP_CONFIG.timezone, "yyyy-MM");
}

/**
 * @param {Date} date
 * @returns {string}
 */
function formatDashboardDate_(date) {
  return Utilities.formatDate(date, APP_CONFIG.timezone, "yyyy-MM-dd");
}

/**
 * @param {*} value
 * @returns {Date|null}
 */
function toDashboardDate_(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    const millis = Math.round((value - 25569) * 86400 * 1000);
    const date = new Date(millis);
    return Number.isNaN(date.getTime()) ? null : new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  const text = normalizeText_(value);
  if (!text) {
    return null;
  }

  const match = text.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})$/);
  if (match) {
    const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

/**
 * @param {*} value
 * @returns {number|null}
 */
function toDashboardNumber_(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  const text = normalizeText_(value).replace(/[^0-9.\-]/g, "");
  if (!text) {
    return null;
  }

  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : null;
}

/**
 * @param {string} bandText
 * @returns {boolean}
 */
function isDashboardHighRisk_(bandText) {
  return /고|severe/i.test(normalizeText_(bandText));
}
