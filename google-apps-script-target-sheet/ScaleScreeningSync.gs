const SCALE_SCREENING_SYNC_CONFIG = {
  propertyKeys: {
    token: "scale_screening_sync_token",
    spreadsheetId: "scale_screening_target_spreadsheet_id",
    recordSheetName: "scale_screening_record_sheet_name",
    answerSheetName: "scale_screening_answer_sheet_name",
    questionnaireSheetName: "scale_screening_questionnaire_sheet_name",
    fieldSheetName: "scale_screening_field_sheet_name",
    optionSheetName: "scale_screening_option_sheet_name",
    workerViewSheetName: "scale_screening_worker_view_sheet_name",
    riskViewSheetName: "scale_screening_risk_view_sheet_name",
    dashboardSheetName: "scale_screening_dashboard_sheet_name",
    settingsSheetName: "scale_screening_settings_sheet_name",
    workspaceVersion: "scale_screening_workspace_version"
  },
  defaults: {
    targetSpreadsheetId: "11y5p7Cp_yN2vggMOlCwn4pKNBEmio-CmkK25Nyd2nIk",
    recordSheetName: "척도검사기록",
    answerSheetName: "척도문항응답",
    questionnaireSheetName: "척도마스터",
    fieldSheetName: "척도문항마스터",
    optionSheetName: "척도선택지마스터",
    workerViewSheetName: "실무자보기",
    riskViewSheetName: "고위험군보기",
    dashboardSheetName: "척도대시보드",
    settingsSheetName: "척도설정"
  },
  recordHeaders: [
    "record_id",
    "exported_at",
    "sync_scope",
    "source_app",
    "organization_name",
    "team_name",
    "contact_note",
    "record_created_at",
    "session_date",
    "questionnaire_id",
    "questionnaire_title",
    "questionnaire_short_title",
    "score_text",
    "normalized_score",
    "band_text",
    "worker_name",
    "client_label",
    "birth_date",
    "gender",
    "age_group",
    "progress_summary",
    "progress_percent",
    "progress_answered",
    "progress_total",
    "signature_present",
    "session_note",
    "highlights",
    "flags",
    "respondent_summary",
    "breakdown_summary",
    "record_json"
  ],
  answerHeaders: [
    "detail_key",
    "record_id",
    "exported_at",
    "session_date",
    "questionnaire_id",
    "questionnaire_title",
    "worker_name",
    "client_label",
    "birth_date",
    "is_subquestion",
    "parent_question_id",
    "question_id",
    "question_number",
    "question_text",
    "answer_label",
    "score",
    "raw_json"
  ],
  questionnaireHeaders: [
    "questionnaire_id",
    "self_seq",
    "title",
    "short_title",
    "recommended_age",
    "question_count",
    "respondent_field_count",
    "question_prompt",
    "intro_text",
    "source_reference_page",
    "source_institution",
    "source_citation",
    "scoring_type",
    "scoring_json",
    "extraction_notes_json",
    "questionnaire_json"
  ],
  fieldHeaders: [
    "field_key",
    "questionnaire_id",
    "field_scope",
    "parent_field_id",
    "field_id",
    "field_number",
    "field_label",
    "field_text",
    "field_type",
    "is_required",
    "option_count",
    "field_json"
  ],
  optionHeaders: [
    "option_key",
    "questionnaire_id",
    "field_scope",
    "parent_field_id",
    "field_id",
    "option_order",
    "option_value",
    "option_label",
    "option_score",
    "option_json"
  ],
  workerViewHeaders: [
    "검사일",
    "대상자",
    "생년월일",
    "척도",
    "원점수",
    "정규화점수",
    "점수표시",
    "결과구간",
    "담당자",
    "비고",
    "경고여부",
    "기록고유값"
  ],
  riskViewHeaders: [
    "검사일",
    "대상자",
    "생년월일",
    "척도",
    "점수표시",
    "결과구간",
    "담당자",
    "경고내용",
    "기록고유값"
  ],
  dashboard: {
    clientNameCell: "B4",
    birthDateCell: "E4",
    namesHelperColumn: 27,
    detailHeaderRow: 11,
    detailStartRow: 12,
    trendStartRow: 11,
    trendStartColumn: 14,
    chartAnchorRow: 1,
    chartAnchorColumn: 12
  },
  settingsRows: [
    ["항목", "값", "설명"],
    ["기록 시트", "척도검사기록", "원본 검사 결과가 저장되는 시트"],
    ["문항응답 시트", "척도문항응답", "문항별 응답 원본 시트"],
    ["실무자 보기 시트", "실무자보기", "실무자가 결과를 조회하는 시트"],
    ["고위험군 보기 시트", "고위험군보기", "고위험 결과만 모아보는 시트"],
    ["대시보드 시트", "척도대시보드", "대상자 검색 및 점수 변화 그래프 시트"],
    ["척도 마스터 시트", "척도마스터", "척도 메타데이터 시트"],
    ["문항 마스터 시트", "척도문항마스터", "문항 정의 시트"],
    ["선택지 마스터 시트", "척도선택지마스터", "선택지 정의 시트"],
    ["검색 사용법", "대상자명을 입력", "척도대시보드 B4에 대상자명을 입력하면 비교표와 그래프가 갱신됩니다."],
    ["생년월일 필터", "선택 입력", "동명이인 구분이 필요할 때 척도대시보드 E4에 생년월일을 입력합니다."]
  ]
};

const SCALE_SCREENING_WORKSPACE_VERSION = "2026-03-27-v3";
const SCALE_SCREENING_STATUS_CACHE_KEY = "scale_screening_sync_status_v2";

function setupScaleScreeningSyncSheets() {
  setScaleScreeningTargetToCurrentSpreadsheet_(false);

  const recordSheet = ensureScaleScreeningSyncSheet_(
    getScaleScreeningRecordSheetName_(),
    SCALE_SCREENING_SYNC_CONFIG.recordHeaders
  );
  const answerSheet = ensureScaleScreeningSyncSheet_(
    getScaleScreeningAnswerSheetName_(),
    SCALE_SCREENING_SYNC_CONFIG.answerHeaders
  );
  const questionnaireSheet = ensureScaleScreeningSyncSheet_(
    getScaleScreeningQuestionnaireSheetName_(),
    SCALE_SCREENING_SYNC_CONFIG.questionnaireHeaders
  );
  const fieldSheet = ensureScaleScreeningSyncSheet_(
    getScaleScreeningFieldSheetName_(),
    SCALE_SCREENING_SYNC_CONFIG.fieldHeaders
  );
  const optionSheet = ensureScaleScreeningSyncSheet_(
    getScaleScreeningOptionSheetName_(),
    SCALE_SCREENING_SYNC_CONFIG.optionHeaders
  );

  formatScaleScreeningSyncSheet_(recordSheet, SCALE_SCREENING_SYNC_CONFIG.recordHeaders.length);
  formatScaleScreeningSyncSheet_(answerSheet, SCALE_SCREENING_SYNC_CONFIG.answerHeaders.length);
  formatScaleScreeningSyncSheet_(questionnaireSheet, SCALE_SCREENING_SYNC_CONFIG.questionnaireHeaders.length);
  formatScaleScreeningSyncSheet_(fieldSheet, SCALE_SCREENING_SYNC_CONFIG.fieldHeaders.length);
  formatScaleScreeningSyncSheet_(optionSheet, SCALE_SCREENING_SYNC_CONFIG.optionHeaders.length);
  const workspaceResult = buildScaleScreeningWorkspace_();

  SpreadsheetApp.getUi().alert(
    "척도검사 시트 준비",
    [
      "척도검사 구글 시트 연동용 시트를 준비했습니다.",
      "기록 시트: " + recordSheet.getName(),
      "문항 시트: " + answerSheet.getName(),
      "척도 시트: " + questionnaireSheet.getName(),
      "문항마스터 시트: " + fieldSheet.getName(),
      "선택지마스터 시트: " + optionSheet.getName(),
      "실무자 보기 시트: " + workspaceResult.workerViewSheetName,
      "고위험군 보기 시트: " + workspaceResult.riskViewSheetName,
      "대시보드 시트: " + workspaceResult.dashboardSheetName,
      "설정 시트: " + workspaceResult.settingsSheetName,
      "다음 단계로 '척도검사 토큰 저장'을 실행한 뒤 웹앱에 같은 토큰을 입력하세요."
    ].join("\n"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function refreshScaleScreeningWorkspace() {
  const result = buildScaleScreeningWorkspace_();

  SpreadsheetApp.getUi().alert(
    "척도검사 분석 시트 새로고침",
    [
      "조회/대시보드 시트를 다시 구성했습니다.",
      "실무자 보기 시트: " + result.workerViewSheetName,
      "고위험군 보기 시트: " + result.riskViewSheetName,
      "대시보드 시트: " + result.dashboardSheetName,
      "설정 시트: " + result.settingsSheetName
    ].join("\n"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function setScaleScreeningTargetToCurrentSpreadsheet() {
  const spreadsheet = setScaleScreeningTargetToCurrentSpreadsheet_(true);

  SpreadsheetApp.getUi().alert(
    "척도검사 대상 시트 지정",
    [
      "현재 스프레드시트를 척도검사 DB 대상으로 저장했습니다.",
      "시트명: " + spreadsheet.getName(),
      "시트 ID: " + spreadsheet.getId()
    ].join("\n"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function setScaleScreeningTargetToCurrentSpreadsheet_(showToast) {
  const spreadsheet = getSpreadsheet_();
  PropertiesService.getScriptProperties().setProperty(
    SCALE_SCREENING_SYNC_CONFIG.propertyKeys.spreadsheetId,
    spreadsheet.getId()
  );

  if (showToast) {
    showToast_("현재 스프레드시트를 척도검사 DB 대상으로 저장했습니다.");
  }

  return spreadsheet;
}

function promptScaleScreeningSyncToken() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "척도검사 토큰 저장",
    "웹앱과 동일하게 사용할 임의의 토큰을 입력하세요.",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const token = normalizeText_(response.getResponseText());
  if (!token) {
    ui.alert("토큰 저장", "비어 있는 토큰은 저장할 수 없습니다.", ui.ButtonSet.OK);
    return;
  }

  PropertiesService.getScriptProperties().setProperty(SCALE_SCREENING_SYNC_CONFIG.propertyKeys.token, token);
  ui.alert("토큰 저장", "척도검사 동기화 토큰을 저장했습니다.", ui.ButtonSet.OK);
}

function showScaleScreeningSyncStatus() {
  const status = buildScaleScreeningSyncStatus_();

  SpreadsheetApp.getUi().alert(
    "척도검사 연동 상태",
    [
      "웹앱 배포 준비 상태를 확인했습니다.",
      "대상 스프레드시트: " + status.spreadsheetName,
      "대상 스프레드시트 ID: " + status.spreadsheetId,
      "토큰 저장 여부: " + (status.tokenConfigured ? "예" : "아니오"),
      "기록 시트: " + status.recordSheetName + " (" + status.recordRowCount + "행)",
      "문항 시트: " + status.answerSheetName + " (" + status.answerRowCount + "행)",
      "척도 시트: " + status.questionnaireSheetName + " (" + status.questionnaireRowCount + "행)",
      "문항마스터 시트: " + status.fieldSheetName + " (" + status.fieldRowCount + "행)",
      "선택지마스터 시트: " + status.optionSheetName + " (" + status.optionRowCount + "행)",
      "실무자 보기 시트: " + status.workerViewSheetName + " (" + status.workerViewRowCount + "행)",
      "고위험군 보기 시트: " + status.riskViewSheetName + " (" + status.riskViewRowCount + "행)",
      "대시보드 시트: " + status.dashboardSheetName,
      "설정 시트: " + status.settingsSheetName,
      "웹앱 URL은 '배포 > 새 배포 > 유형: 웹 앱'으로 발급하세요."
    ].join("\n"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  const callbackName = params.callback;
  const action = normalizeText_(params.action).toLowerCase();

  try {
    if (action === "authorize") {
      return buildScaleScreeningAuthorizePage_();
    }

    validateScaleScreeningSyncToken_(params.token);

    if (action === "searchrecords" || action === "search_records" || action === "search") {
      return createScaleScreeningJsonOutput_(searchScaleScreeningRecords_(params), callbackName);
    }

    if (action === "refreshworkspace" || action === "refresh_workspace") {
      const workspace = buildScaleScreeningWorkspace_();
      applyScaleScreeningDisplayFormats_();
      invalidateScaleScreeningSyncCache_();
      return createScaleScreeningJsonOutput_(Object.assign({
        ok: true
      }, workspace), callbackName);
    }

    return createScaleScreeningJsonOutput_(getCachedScaleScreeningSyncStatus_(), callbackName);
  } catch (error) {
    console.error("척도검사 GET 요청 처리 실패:", {
      action: action || "status",
      message: error.message
    });
    return createScaleScreeningJsonOutput_({
      ok: false,
      error: error.message
    }, callbackName);
  }
}

function buildScaleScreeningAuthorizePage_() {
  try {
    const spreadsheet = getScaleScreeningTargetSpreadsheet_();
    const html = [
      "<!doctype html>",
      "<html lang=\"ko\"><head><meta charset=\"utf-8\">",
      "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">",
      "<title>구글 시트 권한 연결</title>",
      "<style>",
      "body{margin:0;padding:32px;font-family:Arial,'Malgun Gothic',sans-serif;background:#f4f7f4;color:#143127;}",
      ".card{max-width:720px;margin:0 auto;padding:28px;border-radius:24px;background:#fff;border:1px solid rgba(20,49,39,.12);box-shadow:0 16px 40px rgba(16,50,39,.12);}",
      "h1{margin:0 0 12px;font-size:28px;}p{line-height:1.7;margin:0 0 12px;color:#486157;}ul{margin:16px 0 0;padding-left:20px;line-height:1.7;}",
      ".ok{display:inline-block;padding:8px 12px;border-radius:999px;background:#e8f3ed;color:#126b57;font-weight:700;margin-bottom:16px;}",
      ".error{display:inline-block;padding:8px 12px;border-radius:999px;background:#fdecec;color:#b42318;font-weight:700;margin-bottom:16px;}",
      "a{color:#126b57;font-weight:700;text-decoration:none;}",
      "</style></head><body><main class=\"card\">"
    ];

    html.push("<span class=\"ok\">권한 연결 확인 완료</span>");
    html.push("<h1>구글 시트 연동을 사용할 준비가 되었습니다.</h1>");
    html.push("<p>현재 계정으로 대상 스프레드시트에 접근할 수 있습니다. 이제 공개 웹앱으로 돌아가 시트 저장과 조회 기능을 사용할 수 있습니다.</p>");
    html.push("<ul>");
    html.push("<li>연결 시트: " + escapeHtmlScale_(spreadsheet.getName()) + "</li>");
    html.push("<li>연결 시트 ID: " + escapeHtmlScale_(spreadsheet.getId()) + "</li>");
    html.push("<li>공개 웹앱에서 '시트 기능 사용'을 켜고 저장 또는 비교 분석을 실행하세요.</li>");
    html.push("</ul></main></body></html>");
    return HtmlService.createHtmlOutput(html.join("")).setTitle("구글 시트 권한 연결");
  } catch (error) {
    const html = [
      "<!doctype html>",
      "<html lang=\"ko\"><head><meta charset=\"utf-8\">",
      "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">",
      "<title>구글 시트 권한 연결 실패</title>",
      "<style>",
      "body{margin:0;padding:32px;font-family:Arial,'Malgun Gothic',sans-serif;background:#f8f4f4;color:#143127;}",
      ".card{max-width:720px;margin:0 auto;padding:28px;border-radius:24px;background:#fff;border:1px solid rgba(180,54,54,.18);box-shadow:0 16px 40px rgba(16,50,39,.08);}",
      "h1{margin:0 0 12px;font-size:28px;}p{line-height:1.7;margin:0 0 12px;color:#6b4d4d;}",
      ".error{display:inline-block;padding:8px 12px;border-radius:999px;background:#fdecec;color:#b42318;font-weight:700;margin-bottom:16px;}",
      "</style></head><body><main class=\"card\">",
      "<span class=\"error\">권한 연결 실패</span>",
      "<h1>현재 계정으로 구글 시트에 접근할 수 없습니다.</h1>",
      "<p>이 기능은 대상 스프레드시트에 공유 권한이 있는 Google 계정만 사용할 수 있습니다.</p>",
      "<p>오류 내용: " + escapeHtmlScale_(error.message) + "</p>",
      "</main></body></html>"
    ];
    return HtmlService.createHtmlOutput(html.join("")).setTitle("구글 시트 권한 연결 실패");
  }
}

function doPost(e) {
  const lock = LockService.getDocumentLock();
  let hasLock = false;

  try {
    lock.waitLock(30000);
    hasLock = true;

    const payload = parseScaleScreeningSyncPayload_(e);
    validateScaleScreeningSyncToken_(payload.token);

    const result = upsertScaleScreeningPayload_(payload);
    applyScaleScreeningDisplayFormats_();
    invalidateScaleScreeningSyncCache_();
    return createScaleScreeningJsonOutput_({
      ok: true,
      recordSheetName: result.recordSheetName,
      answerSheetName: result.answerSheetName,
      questionnaireSheetName: result.questionnaireSheetName,
      fieldSheetName: result.fieldSheetName,
      optionSheetName: result.optionSheetName,
      recordsInserted: result.recordsInserted,
      recordsUpdated: result.recordsUpdated,
      answersInserted: result.answersInserted,
      answersUpdated: result.answersUpdated,
      questionnairesInserted: result.questionnairesInserted,
      questionnairesUpdated: result.questionnairesUpdated,
      fieldsInserted: result.fieldsInserted,
      fieldsUpdated: result.fieldsUpdated,
      optionsInserted: result.optionsInserted,
      optionsUpdated: result.optionsUpdated
    });
  } catch (error) {
    console.error("척도검사 시트 저장 실패:", error);
    return createScaleScreeningJsonOutput_({
      ok: false,
      error: error.message
    });
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

function buildScaleScreeningSyncStatus_() {
  ensureScaleScreeningWorkspaceSchema_();
  applyScaleScreeningDisplayFormats_();

  const spreadsheet = getScaleScreeningTargetSpreadsheet_();
  const recordSheetName = getScaleScreeningRecordSheetName_();
  const answerSheetName = getScaleScreeningAnswerSheetName_();
  const questionnaireSheetName = getScaleScreeningQuestionnaireSheetName_();
  const fieldSheetName = getScaleScreeningFieldSheetName_();
  const optionSheetName = getScaleScreeningOptionSheetName_();
  const workerViewSheetName = getScaleScreeningWorkerViewSheetName_();
  const riskViewSheetName = getScaleScreeningRiskViewSheetName_();
  const dashboardSheetName = getScaleScreeningDashboardSheetName_();
  const settingsSheetName = getScaleScreeningSettingsSheetName_();
  const recordSheet = getScaleScreeningSheetIfExists_(recordSheetName);
  const answerSheet = getScaleScreeningSheetIfExists_(answerSheetName);
  const questionnaireSheet = getScaleScreeningSheetIfExists_(questionnaireSheetName);
  const fieldSheet = getScaleScreeningSheetIfExists_(fieldSheetName);
  const optionSheet = getScaleScreeningSheetIfExists_(optionSheetName);
  const workerViewSheet = getScaleScreeningSheetIfExists_(workerViewSheetName);
  const riskViewSheet = getScaleScreeningSheetIfExists_(riskViewSheetName);

  return {
    ok: true,
    tokenConfigured: Boolean(getScaleScreeningSyncToken_()),
    spreadsheetId: spreadsheet.getId(),
    spreadsheetName: spreadsheet.getName(),
    recordSheetName: recordSheetName,
    answerSheetName: answerSheetName,
    questionnaireSheetName: questionnaireSheetName,
    fieldSheetName: fieldSheetName,
    optionSheetName: optionSheetName,
    workerViewSheetName: workerViewSheetName,
    riskViewSheetName: riskViewSheetName,
    dashboardSheetName: dashboardSheetName,
    settingsSheetName: settingsSheetName,
    recordRowCount: Math.max((recordSheet && recordSheet.getLastRow()) || 0, 1) - 1,
    answerRowCount: Math.max((answerSheet && answerSheet.getLastRow()) || 0, 1) - 1,
    questionnaireRowCount: Math.max((questionnaireSheet && questionnaireSheet.getLastRow()) || 0, 1) - 1,
    fieldRowCount: Math.max((fieldSheet && fieldSheet.getLastRow()) || 0, 1) - 1,
    optionRowCount: Math.max((optionSheet && optionSheet.getLastRow()) || 0, 1) - 1,
    workerViewRowCount: Math.max((workerViewSheet && workerViewSheet.getLastRow()) || 0, 1) - 1,
    riskViewRowCount: Math.max((riskViewSheet && riskViewSheet.getLastRow()) || 0, 1) - 1
  };
}

function getCachedScaleScreeningSyncStatus_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(SCALE_SCREENING_STATUS_CACHE_KEY);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (error) {
      cache.remove(SCALE_SCREENING_STATUS_CACHE_KEY);
    }
  }

  const status = buildScaleScreeningSyncStatus_();
  cache.put(SCALE_SCREENING_STATUS_CACHE_KEY, JSON.stringify(status), 60);
  return status;
}

function invalidateScaleScreeningSyncCache_() {
  CacheService.getScriptCache().remove(SCALE_SCREENING_STATUS_CACHE_KEY);
}

/**
 * 척도검사 기록 시트에서 대상자 기준 검색 결과를 반환합니다.
 *
 * @param {{name?: string, birthDate?: string, questionnaireId?: string, limit?: string}} params 조회 파라미터
 * @returns {{ok: boolean, query: object, recordCount: number, records: object[]}}
 */
function searchScaleScreeningRecords_(params) {
  console.time("searchScaleScreeningRecords");

  try {
    const nameQuery = normalizeScaleLookupText_(params.name);
    const birthDateQuery = normalizeScaleSearchDate_(params.birthDate);
    const questionnaireIdQuery = normalizeScaleLookupText_(params.questionnaireId);
    const limit = Math.min(Math.max(parseInt(params.limit, 10) || 200, 1), 500);

    if (!nameQuery && !birthDateQuery) {
      throw new Error("대상자 이름 또는 생년월일을 입력해주세요.");
    }

    const recordSheet = getScaleScreeningSheetIfExists_(getScaleScreeningRecordSheetName_());
    if (!recordSheet || recordSheet.getLastRow() < 2) {
      return {
        ok: true,
        query: {
          name: normalizeText_(params.name),
          birthDate: birthDateQuery,
          questionnaireId: normalizeText_(params.questionnaireId)
        },
        recordCount: 0,
        records: []
      };
    }

    const values = recordSheet.getRange(1, 1, recordSheet.getLastRow(), recordSheet.getLastColumn()).getDisplayValues();
    const headerIndexMap = buildScaleHeaderIndexMap_(values[0]);
    const matchedRecords = [];

    for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
      const row = values[rowIndex];
      const clientLabel = normalizeScaleLookupText_(row[headerIndexMap.client_label]);
      const birthDate = normalizeScaleSearchDate_(row[headerIndexMap.birth_date]);
      const questionnaireId = normalizeScaleLookupText_(row[headerIndexMap.questionnaire_id]);

      if (nameQuery && clientLabel !== nameQuery) {
        continue;
      }
      if (birthDateQuery && birthDate !== birthDateQuery) {
        continue;
      }
      if (questionnaireIdQuery && questionnaireId !== questionnaireIdQuery) {
        continue;
      }

      const parsedRecord = parseScaleScreeningSearchRecord_(row, headerIndexMap);
      if (parsedRecord) {
        matchedRecords.push(parsedRecord);
      }
    }

    matchedRecords.sort(function(a, b) {
      return getScaleSortDateValue_(a).localeCompare(getScaleSortDateValue_(b));
    });

    const limitedRecords = matchedRecords.length > limit
      ? matchedRecords.slice(matchedRecords.length - limit)
      : matchedRecords;

    return {
      ok: true,
      query: {
        name: normalizeText_(params.name),
        birthDate: birthDateQuery,
        questionnaireId: normalizeText_(params.questionnaireId)
      },
      recordCount: limitedRecords.length,
      records: limitedRecords
    };
  } finally {
    console.timeEnd("searchScaleScreeningRecords");
  }
}

function buildScaleHeaderIndexMap_(headers) {
  return (headers || []).reduce(function(result, header, index) {
    const key = normalizeText_(header);
    if (key) {
      result[key] = index;
    }
    return result;
  }, {});
}

function parseScaleScreeningSearchRecord_(row, headerIndexMap) {
  const recordId = normalizeText_(row[headerIndexMap.record_id]);
  if (!recordId) {
    return null;
  }

  const parsedRecord = safeParseScaleJson_(row[headerIndexMap.record_json]) || {};
  const rawEvaluation = parsedRecord.evaluation || {};
  const rawMeta = parsedRecord.meta || {};
  const rawFlags = Array.isArray(rawEvaluation.flags) ? rawEvaluation.flags : [];
  const normalizedScore = rawEvaluation.normalizedScore !== null && rawEvaluation.normalizedScore !== undefined
    ? Number(rawEvaluation.normalizedScore)
    : parseScaleNumericValue_(row[headerIndexMap.normalized_score]);

  return {
    id: recordId,
    questionnaireId: normalizeText_(row[headerIndexMap.questionnaire_id]),
    questionnaireTitle: normalizeText_(row[headerIndexMap.questionnaire_title]),
    shortTitle: normalizeText_(row[headerIndexMap.questionnaire_short_title]),
    createdAt: normalizeText_(parsedRecord.createdAt) || normalizeText_(row[headerIndexMap.record_created_at]),
    meta: {
      sessionDate: normalizeScaleSearchDate_(row[headerIndexMap.session_date]),
      workerName: normalizeText_(row[headerIndexMap.worker_name]) || normalizeText_(rawMeta.workerName),
      clientLabel: normalizeText_(row[headerIndexMap.client_label]) || normalizeText_(rawMeta.clientLabel),
      birthDate: normalizeScaleSearchDate_(row[headerIndexMap.birth_date]) || normalizeScaleSearchDate_(rawMeta.birthDate),
      sessionNote: normalizeText_(row[headerIndexMap.session_note]) || normalizeText_(rawMeta.sessionNote)
    },
    progress: {
      percent: parseScaleNumericValue_(row[headerIndexMap.progress_percent]),
      answered: parseScaleNumericValue_(row[headerIndexMap.progress_answered]),
      total: parseScaleNumericValue_(row[headerIndexMap.progress_total]),
      summary: normalizeText_(row[headerIndexMap.progress_summary])
    },
    evaluation: {
      score: rawEvaluation.score !== null && rawEvaluation.score !== undefined
        ? Number(rawEvaluation.score)
        : parseScaleNumericValue_(row[headerIndexMap.score_text]),
      maxScore: rawEvaluation.maxScore !== null && rawEvaluation.maxScore !== undefined
        ? Number(rawEvaluation.maxScore)
        : null,
      normalizedScore: Number.isFinite(normalizedScore) ? normalizedScore : 0,
      scoreText: normalizeText_(row[headerIndexMap.score_text]) || normalizeText_(rawEvaluation.scoreText),
      bandText: normalizeText_(row[headerIndexMap.band_text]) || normalizeText_(rawEvaluation.bandText),
      highlights: Array.isArray(rawEvaluation.highlights) ? rawEvaluation.highlights : [],
      flags: rawFlags.length
        ? rawFlags.map(function(flag) {
          return {
            level: normalizeText_(flag && flag.level) || "warn",
            text: normalizeText_(flag && flag.text)
          };
        }).filter(function(flag) {
          return flag.text;
        })
        : normalizeText_(row[headerIndexMap.flags])
            .split("|")
            .map(function(text) {
              return normalizeText_(text);
            })
            .filter(Boolean)
            .map(function(text) {
              return { level: "warn", text: text };
            })
    }
  };
}

function safeParseScaleJson_(value) {
  const text = normalizeText_(value);
  if (!text) {
    return null;
  }

  try {
    return JSON.parse(text);
  } catch (error) {
    return null;
  }
}

function normalizeScaleLookupText_(value) {
  return normalizeText_(value).replace(/\s+/g, " ").toLowerCase();
}

function normalizeScaleSearchDate_(value) {
  const text = normalizeText_(value);
  if (!text) {
    return "";
  }
  if (/^\d{5}(?:\.\d+)?$/.test(text)) {
    return convertScaleSerialDateToText_(Number(text));
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return text;
  }

  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, "Asia/Seoul", "yyyy-MM-dd");
  }

  return text.slice(0, 10);
}

function convertScaleSerialDateToText_(serialValue) {
  const serial = Number(serialValue);
  if (!isFinite(serial)) {
    return "";
  }

  const milliseconds = Math.round((serial - 25569) * 86400 * 1000);
  const date = new Date(milliseconds);
  if (isNaN(date.getTime())) {
    return "";
  }

  return Utilities.formatDate(date, "Asia/Seoul", "yyyy-MM-dd");
}

function parseScaleNumericValue_(value) {
  const text = normalizeText_(value);
  if (!text) {
    return null;
  }

  const matched = text.match(/-?\d+(?:\.\d+)?/);
  return matched ? Number(matched[0]) : null;
}

function getScaleSortDateValue_(record) {
  return normalizeText_(
    record &&
    record.meta &&
    record.meta.sessionDate
  ) || normalizeText_(record && record.createdAt);
}

function parseScaleScreeningSyncPayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error("요청 본문이 비어 있습니다.");
  }

  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (error) {
    throw new Error("JSON 본문을 해석할 수 없습니다.");
  }

  const hasRecords = payload && Array.isArray(payload.records) && payload.records.length;
  const hasQuestionnaires = payload && Array.isArray(payload.questionnaires) && payload.questionnaires.length;

  if (!hasRecords && !hasQuestionnaires) {
    throw new Error("전송할 검사 기록 또는 척도 마스터가 없습니다.");
  }

  return payload;
}

function validateScaleScreeningSyncToken_(token) {
  const configuredToken = getScaleScreeningSyncToken_();

  if (!configuredToken) {
    return;
  }

  if (normalizeText_(token) !== configuredToken) {
    throw new Error("동기화 토큰이 일치하지 않습니다.");
  }
}

function upsertScaleScreeningPayload_(payload) {
  const result = {
    recordSheetName: "",
    answerSheetName: "",
    questionnaireSheetName: "",
    fieldSheetName: "",
    optionSheetName: "",
    recordsInserted: 0,
    recordsUpdated: 0,
    answersInserted: 0,
    answersUpdated: 0,
    questionnairesInserted: 0,
    questionnairesUpdated: 0,
    fieldsInserted: 0,
    fieldsUpdated: 0,
    optionsInserted: 0,
    optionsUpdated: 0
  };

  if (Array.isArray(payload.records) && payload.records.length) {
    mergeScaleScreeningResult_(result, upsertScaleScreeningRecordPayload_(payload));
  }

  if (Array.isArray(payload.questionnaires) && payload.questionnaires.length) {
    mergeScaleScreeningResult_(result, upsertScaleScreeningQuestionnairePayload_(payload));
  }

  return result;
}

function mergeScaleScreeningResult_(target, partial) {
  Object.keys(partial).forEach(function(key) {
    if (typeof partial[key] === "number") {
      target[key] = (target[key] || 0) + partial[key];
      return;
    }

    if (partial[key]) {
      target[key] = partial[key];
    }
  });
}

function upsertScaleScreeningRecordPayload_(payload) {
  const recordSheet = ensureScaleScreeningSyncSheet_(
    getScaleScreeningRecordSheetName_(),
    SCALE_SCREENING_SYNC_CONFIG.recordHeaders
  );
  const answerSheet = ensureScaleScreeningSyncSheet_(
    getScaleScreeningAnswerSheetName_(),
    SCALE_SCREENING_SYNC_CONFIG.answerHeaders
  );

  const recordRows = payload.records.map(function(record) {
    return buildScaleScreeningRecordRow_(record, payload);
  });

  const answerRows = [];
  payload.records.forEach(function(record) {
    buildScaleScreeningAnswerRows_(record, payload).forEach(function(row) {
      answerRows.push(row);
    });
  });

  const recordResult = upsertRowsByKey_(
    recordSheet,
    SCALE_SCREENING_SYNC_CONFIG.recordHeaders,
    recordRows,
    "record_id"
  );
  const answerResult = upsertRowsByKey_(
    answerSheet,
    SCALE_SCREENING_SYNC_CONFIG.answerHeaders,
    answerRows,
    "detail_key"
  );

  return {
    recordSheetName: recordSheet.getName(),
    answerSheetName: answerSheet.getName(),
    recordsInserted: recordResult.inserted,
    recordsUpdated: recordResult.updated,
    answersInserted: answerResult.inserted,
    answersUpdated: answerResult.updated
  };
}

function upsertScaleScreeningQuestionnairePayload_(payload) {
  const questionnaireSheet = ensureScaleScreeningSyncSheet_(
    getScaleScreeningQuestionnaireSheetName_(),
    SCALE_SCREENING_SYNC_CONFIG.questionnaireHeaders
  );
  const fieldSheet = ensureScaleScreeningSyncSheet_(
    getScaleScreeningFieldSheetName_(),
    SCALE_SCREENING_SYNC_CONFIG.fieldHeaders
  );
  const optionSheet = ensureScaleScreeningSyncSheet_(
    getScaleScreeningOptionSheetName_(),
    SCALE_SCREENING_SYNC_CONFIG.optionHeaders
  );

  const questionnaireRows = payload.questionnaires.map(function(questionnaire) {
    return buildScaleScreeningQuestionnaireRow_(questionnaire);
  });

  const fieldRows = [];
  const optionRows = [];
  payload.questionnaires.forEach(function(questionnaire) {
    buildScaleScreeningFieldRows_(questionnaire).forEach(function(row) {
      fieldRows.push(row);
    });
    buildScaleScreeningOptionRows_(questionnaire).forEach(function(row) {
      optionRows.push(row);
    });
  });

  const questionnaireResult = upsertRowsByKey_(
    questionnaireSheet,
    SCALE_SCREENING_SYNC_CONFIG.questionnaireHeaders,
    questionnaireRows,
    "questionnaire_id"
  );
  const fieldResult = upsertRowsByKey_(
    fieldSheet,
    SCALE_SCREENING_SYNC_CONFIG.fieldHeaders,
    fieldRows,
    "field_key"
  );
  const optionResult = upsertRowsByKey_(
    optionSheet,
    SCALE_SCREENING_SYNC_CONFIG.optionHeaders,
    optionRows,
    "option_key"
  );

  return {
    questionnaireSheetName: questionnaireSheet.getName(),
    fieldSheetName: fieldSheet.getName(),
    optionSheetName: optionSheet.getName(),
    questionnairesInserted: questionnaireResult.inserted,
    questionnairesUpdated: questionnaireResult.updated,
    fieldsInserted: fieldResult.inserted,
    fieldsUpdated: fieldResult.updated,
    optionsInserted: optionResult.inserted,
    optionsUpdated: optionResult.updated
  };
}

function buildScaleScreeningQuestionnaireRow_(questionnaire) {
  return {
    questionnaire_id: toCellText_(questionnaire.id),
    self_seq: toCellText_(questionnaire.selfSeq),
    title: toCellText_(questionnaire.title),
    short_title: toCellText_(questionnaire.shortTitle),
    recommended_age: toCellText_(questionnaire.recommendedAge),
    question_count: String((questionnaire.questions || []).length),
    respondent_field_count: String((questionnaire.respondentFields || []).length),
    question_prompt: toCellText_(questionnaire.questionPrompt),
    intro_text: joinScaleTextList_(questionnaire.intro),
    source_reference_page: toCellText_(questionnaire.source && questionnaire.source.referencePage),
    source_institution: toCellText_(questionnaire.source && questionnaire.source.institution),
    source_citation: toCellText_(questionnaire.source && questionnaire.source.citation),
    scoring_type: toCellText_(questionnaire.scoring && questionnaire.scoring.type),
    scoring_json: safeStringifyScaleValue_(questionnaire.scoring || {}),
    extraction_notes_json: safeStringifyScaleValue_(questionnaire.extractionNotes || []),
    questionnaire_json: safeStringifyScaleValue_(questionnaire)
  };
}

function buildScaleScreeningFieldRows_(questionnaire) {
  const rows = [];

  (questionnaire.respondentFields || []).forEach(function(field) {
    rows.push(buildScaleFieldRow_(questionnaire, "respondent", "", field));
  });

  (questionnaire.questions || []).forEach(function(question) {
    rows.push(buildScaleFieldRow_(questionnaire, "question", "", question));
    (question.subQuestions || []).forEach(function(subQuestion) {
      rows.push(buildScaleFieldRow_(questionnaire, "subquestion", question.id, subQuestion));
    });
  });

  return rows;
}

function buildScaleScreeningOptionRows_(questionnaire) {
  const rows = [];

  (questionnaire.respondentFields || []).forEach(function(field) {
    buildScaleOptionRowsForField_(questionnaire, "respondent", "", field).forEach(function(row) {
      rows.push(row);
    });
  });

  (questionnaire.questions || []).forEach(function(question) {
    buildScaleOptionRowsForField_(questionnaire, "question", "", question).forEach(function(row) {
      rows.push(row);
    });
    (question.subQuestions || []).forEach(function(subQuestion) {
      buildScaleOptionRowsForField_(questionnaire, "subquestion", question.id, subQuestion).forEach(function(row) {
        rows.push(row);
      });
    });
  });

  return rows;
}

function buildScaleFieldRow_(questionnaire, fieldScope, parentFieldId, field) {
  return {
    field_key: toCellText_(questionnaire.id) + "::" + fieldScope + "::" + toCellText_(field.id),
    questionnaire_id: toCellText_(questionnaire.id),
    field_scope: fieldScope,
    parent_field_id: toCellText_(parentFieldId),
    field_id: toCellText_(field.id),
    field_number: toCellText_(field.number),
    field_label: toCellText_(field.label),
    field_text: toCellText_(field.text),
    field_type: toCellText_(field.type || "single_choice"),
    is_required: field.required ? "Y" : "N",
    option_count: String((field.options || []).length),
    field_json: safeStringifyScaleValue_(field)
  };
}

function buildScaleOptionRowsForField_(questionnaire, fieldScope, parentFieldId, field) {
  return (field.options || []).map(function(option, index) {
    return {
      option_key: [
        toCellText_(questionnaire.id),
        fieldScope,
        toCellText_(field.id),
        String(index + 1)
      ].join("::"),
      questionnaire_id: toCellText_(questionnaire.id),
      field_scope: fieldScope,
      parent_field_id: toCellText_(parentFieldId),
      field_id: toCellText_(field.id),
      option_order: String(index + 1),
      option_value: toCellText_(option.value),
      option_label: toCellText_(option.label),
      option_score: option.score === null || option.score === undefined ? "" : String(option.score),
      option_json: safeStringifyScaleValue_(option)
    };
  });
}

function buildScaleScreeningRecordRow_(record, payload) {
  const respondentDisplay = Array.isArray(record.respondentDisplay) ? record.respondentDisplay : [];
  const flags = Array.isArray(record.evaluation && record.evaluation.flags)
    ? record.evaluation.flags.map(function(flag) {
      return flag && flag.text ? flag.text : "";
    }).filter(Boolean)
    : [];

  return {
    record_id: toCellText_(record.id),
    exported_at: toCellText_(payload.sentAt),
    sync_scope: toCellText_(payload.syncScope),
    source_app: toCellText_(payload.source),
    organization_name: toCellText_(payload.appSettings && payload.appSettings.organizationName),
    team_name: toCellText_(payload.appSettings && payload.appSettings.teamName),
    contact_note: toCellText_(payload.appSettings && payload.appSettings.contactNote),
    record_created_at: toCellText_(record.createdAt),
    session_date: toCellText_(record.meta && record.meta.sessionDate),
    questionnaire_id: toCellText_(record.questionnaireId),
    questionnaire_title: toCellText_(record.questionnaireTitle),
    questionnaire_short_title: toCellText_(record.shortTitle),
    score_text: toCellText_(record.evaluation && record.evaluation.scoreText),
    normalized_score: formatScaleNormalizedScore_(getScaleNormalizedScore_(record)),
    band_text: toCellText_(record.evaluation && record.evaluation.bandText),
    worker_name: toCellText_(record.meta && record.meta.workerName),
    client_label: toCellText_(record.meta && record.meta.clientLabel),
    birth_date: toCellText_(record.meta && record.meta.birthDate),
    gender: findDisplayValueByLabel_(respondentDisplay, "성별"),
    age_group: findDisplayValueByLabel_(respondentDisplay, "연령대"),
    progress_summary: buildScaleScreeningProgressSummary_(record && record.progress ? record.progress : {}),
    progress_percent: (record && record.progress && record.progress.percent !== null && record.progress.percent !== undefined) ? String(record.progress.percent) : "",
    progress_answered: (record && record.progress && record.progress.answered !== null && record.progress.answered !== undefined) ? String(record.progress.answered) : "",
    progress_total: (record && record.progress && record.progress.total !== null && record.progress.total !== undefined) ? String(record.progress.total) : "",
    signature_present: record.meta && record.meta.signatureDataUrl ? "Y" : "N",
    session_note: toCellText_(record.meta && record.meta.sessionNote),
    highlights: joinScaleTextList_(record.evaluation && record.evaluation.highlights),
    flags: joinScaleTextList_(flags),
    respondent_summary: buildRespondentSummary_(respondentDisplay),
    breakdown_summary: buildBreakdownSummary_(record.breakdown),
    record_json: safeStringifyScaleValue_(record)
  };
}

function getScaleNormalizedScore_(record) {
  const evaluation = (record && record.evaluation) || {};
  const explicitNormalized = Number(evaluation.normalizedScore);

  if (Number.isFinite(explicitNormalized)) {
    return clampScaleValue_(explicitNormalized, 0, 100);
  }

  const score = Number(evaluation.score);
  const maxScore = Number(evaluation.maxScore);
  if (Number.isFinite(score) && Number.isFinite(maxScore) && maxScore > 0) {
    return clampScaleValue_((score / maxScore) * 100, 0, 100);
  }

  return null;
}

function formatScaleNormalizedScore_(value) {
  if (!Number.isFinite(value)) {
    return "";
  }
  return String(Math.round(value * 10) / 10);
}

function clampScaleValue_(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function buildScaleScreeningAnswerRows_(record, payload) {
  const rows = [];
  const breakdown = Array.isArray(record.breakdown) ? record.breakdown : [];

  breakdown.forEach(function(item) {
    const parentKey = toCellText_(record.id) + "::" + toCellText_(item.id || item.number);

    rows.push({
      detail_key: parentKey,
      record_id: toCellText_(record.id),
      exported_at: toCellText_(payload.sentAt),
      session_date: toCellText_(record.meta && record.meta.sessionDate),
      questionnaire_id: toCellText_(record.questionnaireId),
      questionnaire_title: toCellText_(record.questionnaireTitle),
      worker_name: toCellText_(record.meta && record.meta.workerName),
      client_label: toCellText_(record.meta && record.meta.clientLabel),
      birth_date: toCellText_(record.meta && record.meta.birthDate),
      is_subquestion: "N",
      parent_question_id: "",
      question_id: toCellText_(item.id || item.number),
      question_number: toCellText_(item.number),
      question_text: toCellText_(item.text),
      answer_label: toCellText_(item.answerLabel),
      score: item.score === null || item.score === undefined ? "" : String(item.score),
      raw_json: safeStringifyScaleValue_(item)
    });

    (Array.isArray(item.subAnswers) ? item.subAnswers : []).forEach(function(subItem, index) {
      rows.push({
        detail_key: parentKey + "::sub::" + String(index + 1),
        record_id: toCellText_(record.id),
        exported_at: toCellText_(payload.sentAt),
        session_date: toCellText_(record.meta && record.meta.sessionDate),
        questionnaire_id: toCellText_(record.questionnaireId),
        questionnaire_title: toCellText_(record.questionnaireTitle),
        worker_name: toCellText_(record.meta && record.meta.workerName),
        client_label: toCellText_(record.meta && record.meta.clientLabel),
        birth_date: toCellText_(record.meta && record.meta.birthDate),
        is_subquestion: "Y",
        parent_question_id: toCellText_(item.id || item.number),
        question_id: (toCellText_(item.id || item.number) + "::sub::" + String(index + 1)),
        question_number: toCellText_(subItem.number),
        question_text: toCellText_(subItem.text),
        answer_label: toCellText_(subItem.answerLabel),
        score: subItem.score === null || subItem.score === undefined ? "" : String(subItem.score),
        raw_json: safeStringifyScaleValue_(subItem)
      });
    });
  });

  return rows;
}

function upsertRowsByKey_(sheet, headers, rowObjects, keyField) {
  const keyIndex = headers.indexOf(keyField);
  if (keyIndex === -1) {
    throw new Error("키 필드를 찾을 수 없습니다: " + keyField);
  }

  ensureSheetSize_(sheet, Math.max(sheet.getLastRow(), 2), headers.length);
  const existingRows = getExistingRowNumberMap_(sheet, headers, keyField);
  const appendRows = [];
  const updateRows = {};
  let inserted = 0;
  let updated = 0;
  let existingData = [];

  if (sheet.getLastRow() >= 2) {
    existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getDisplayValues();
  }

  rowObjects.forEach(function(rowObject) {
    const key = normalizeText_(rowObject[keyField]);
    if (!key) {
      return;
    }

    const rowValues = headers.map(function(header) {
      return toCellText_(rowObject[header]);
    });

    if (existingRows[key]) {
      updateRows[existingRows[key] - 2] = rowValues;
      updated += 1;
      return;
    }

    appendRows.push(rowValues);
    existingRows[key] = sheet.getLastRow() + appendRows.length;
    inserted += 1;
  });

  if (appendRows.length) {
    const startRow = sheet.getLastRow() + 1;
    ensureSheetSize_(sheet, startRow + appendRows.length, headers.length);
    sheet.getRange(startRow, 1, appendRows.length, headers.length).setValues(appendRows);
  }

  const updateIndexes = Object.keys(updateRows);
  if (updateIndexes.length && existingData.length) {
    updateIndexes.forEach(function(indexText) {
      const rowIndex = Number(indexText);
      existingData[rowIndex] = updateRows[indexText];
    });
    sheet.getRange(2, 1, existingData.length, headers.length).setValues(existingData);
  }

  return {
    inserted: inserted,
    updated: updated
  };
}

function buildScaleScreeningWorkspace_() {
  console.time("buildScaleScreeningWorkspace");
  const workerViewSheet = getOrCreateScaleScreeningSheet_(getScaleScreeningWorkerViewSheetName_());
  const riskViewSheet = getOrCreateScaleScreeningSheet_(getScaleScreeningRiskViewSheetName_());
  const dashboardSheet = getOrCreateScaleScreeningSheet_(getScaleScreeningDashboardSheetName_());
  const settingsSheet = getOrCreateScaleScreeningSheet_(getScaleScreeningSettingsSheetName_());

  buildScaleScreeningWorkerViewSheet_(workerViewSheet);
  buildScaleScreeningRiskViewSheet_(riskViewSheet);
  buildScaleScreeningDashboardSheet_(dashboardSheet);
  buildScaleScreeningSettingsSheet_(settingsSheet);

  console.timeEnd("buildScaleScreeningWorkspace");
  PropertiesService.getScriptProperties().setProperty(
    SCALE_SCREENING_SYNC_CONFIG.propertyKeys.workspaceVersion,
    SCALE_SCREENING_WORKSPACE_VERSION
  );
  return {
    workerViewSheetName: workerViewSheet.getName(),
    riskViewSheetName: riskViewSheet.getName(),
    dashboardSheetName: dashboardSheet.getName(),
    settingsSheetName: settingsSheet.getName()
  };
}

function ensureScaleScreeningWorkspaceSchema_() {
  const properties = PropertiesService.getScriptProperties();
  const currentVersion = normalizeText_(properties.getProperty(
    SCALE_SCREENING_SYNC_CONFIG.propertyKeys.workspaceVersion
  ));

  if (currentVersion === SCALE_SCREENING_WORKSPACE_VERSION) {
    return false;
  }

  const lock = LockService.getScriptLock();
  let hasLock = false;

  try {
    lock.waitLock(30000);
    hasLock = true;

    const lockedVersion = normalizeText_(properties.getProperty(
      SCALE_SCREENING_SYNC_CONFIG.propertyKeys.workspaceVersion
    ));
    if (lockedVersion === SCALE_SCREENING_WORKSPACE_VERSION) {
      return false;
    }

    buildScaleScreeningWorkspace_();
    applyScaleScreeningDisplayFormats_();
    invalidateScaleScreeningSyncCache_();
    return true;
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

function buildScaleScreeningWorkerViewSheet_(sheet) {
  const recordSheetRef = escapeScaleSheetNameForFormula_(getScaleScreeningRecordSheetName_());
  const formula = '=IFERROR(SORT(FILTER({' +
    recordSheetRef + '!I2:I,' +
    recordSheetRef + '!Q2:Q,' +
    recordSheetRef + '!R2:R,' +
    recordSheetRef + '!K2:K,' +
    'IF(' + recordSheetRef + '!M2:M="",,IFERROR(VALUE(REGEXEXTRACT(' + recordSheetRef + '!M2:M,"-?\\d+(?:\\.\\d+)?")),)),' +
    'IF(' + recordSheetRef + '!N2:N="",,VALUE(' + recordSheetRef + '!N2:N)),' +
    recordSheetRef + '!M2:M,' +
    recordSheetRef + '!O2:O,' +
    recordSheetRef + '!P2:P,' +
    recordSheetRef + '!Z2:Z,' +
    recordSheetRef + '!AB2:AB,' +
    recordSheetRef + '!A2:A' +
    '},' + recordSheetRef + '!A2:A<>""),1,FALSE),"")';

  sheet.clear();
  ensureSheetSize_(sheet, 200, SCALE_SCREENING_SYNC_CONFIG.workerViewHeaders.length);
  sheet.getRange(1, 1, 1, SCALE_SCREENING_SYNC_CONFIG.workerViewHeaders.length)
    .setValues([SCALE_SCREENING_SYNC_CONFIG.workerViewHeaders]);
  sheet.getRange(2, 1).setFormula(formula);
  sheet.setFrozenRows(1);
  styleHeaderRow_(sheet, 1, SCALE_SCREENING_SYNC_CONFIG.workerViewHeaders.length);
  sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("C:C").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("F:F").setNumberFormat("0.0");
  applyScaleBanding_(sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 2), SCALE_SCREENING_SYNC_CONFIG.workerViewHeaders.length));
  applyScaleWorkerViewRules_(sheet);
  setScaleColumnWidths_(sheet, [110, 150, 110, 160, 90, 100, 120, 110, 110, 180, 120, 160]);
}

function buildScaleScreeningRiskViewSheet_(sheet) {
  const workerSheetRef = escapeScaleSheetNameForFormula_(getScaleScreeningWorkerViewSheetName_());
  const formula = '=IFERROR(SORT(FILTER({' +
    workerSheetRef + '!A2:A,' +
    workerSheetRef + '!B2:B,' +
    workerSheetRef + '!C2:C,' +
    workerSheetRef + '!D2:D,' +
    workerSheetRef + '!G2:G,' +
    workerSheetRef + '!H2:H,' +
    workerSheetRef + '!I2:I,' +
    workerSheetRef + '!K2:K,' +
    workerSheetRef + '!L2:L' +
    '},' +
    workerSheetRef + '!L2:L<>"",' +
    '((' + workerSheetRef + '!K2:K<>"")+REGEXMATCH(' + workerSheetRef + '!H2:H,"^(A|B|C)"))>0' +
    '),1,FALSE),"")';

  sheet.clear();
  ensureSheetSize_(sheet, 200, SCALE_SCREENING_SYNC_CONFIG.riskViewHeaders.length);
  sheet.getRange(1, 1, 1, SCALE_SCREENING_SYNC_CONFIG.riskViewHeaders.length)
    .setValues([SCALE_SCREENING_SYNC_CONFIG.riskViewHeaders]);
  sheet.getRange(2, 1).setFormula(formula);
  sheet.setFrozenRows(1);
  styleHeaderRow_(sheet, 1, SCALE_SCREENING_SYNC_CONFIG.riskViewHeaders.length);
  sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("C:C").setNumberFormat("yyyy-mm-dd");
  applyScaleBanding_(sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 2), SCALE_SCREENING_SYNC_CONFIG.riskViewHeaders.length));
  applyScaleRiskViewRules_(sheet);
  setScaleColumnWidths_(sheet, [110, 150, 110, 160, 120, 110, 110, 220, 160]);
}

function buildScaleScreeningSettingsSheet_(sheet) {
  sheet.clear();
  ensureSheetSize_(sheet, SCALE_SCREENING_SYNC_CONFIG.settingsRows.length + 10, 3);
  sheet.getRange(1, 1, SCALE_SCREENING_SYNC_CONFIG.settingsRows.length, 3)
    .setValues(SCALE_SCREENING_SYNC_CONFIG.settingsRows);
  sheet.setFrozenRows(1);
  styleHeaderRow_(sheet, 1, 3);
  sheet.autoResizeColumns(1, 3);
}

function buildScaleScreeningDashboardSheet_(sheet) {
  const workerSheetRef = escapeScaleSheetNameForFormula_(getScaleScreeningWorkerViewSheetName_());
  const dashboardConfig = SCALE_SCREENING_SYNC_CONFIG.dashboard;
  const detailFormula = "=IF($B$4=\"\",\"\",IFERROR(SORT(FILTER({" +
    workerSheetRef + '!A2:A,' +
    workerSheetRef + '!D2:D,' +
    workerSheetRef + '!E2:E,' +
    workerSheetRef + '!F2:F,' +
    workerSheetRef + '!G2:G,' +
    workerSheetRef + '!H2:H,' +
    workerSheetRef + '!I2:I,' +
    workerSheetRef + '!J2:J,' +
    workerSheetRef + '!K2:K,' +
    workerSheetRef + '!L2:L' +
    "}," +
    workerSheetRef + '!B2:B=$B$4,' +
    "IF($E$4=\"\"," + workerSheetRef + '!A2:A<>"",' + workerSheetRef + '!C2:C=TEXT($E$4,\"yyyy-mm-dd\"))' +
    "),1,TRUE),{\"검색 결과가 없습니다.\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\"}))";
  const trendFormula = "=IF($B$4=\"\",\"\",IFERROR(QUERY(FILTER({" +
    workerSheetRef + '!A2:A,' +
    workerSheetRef + '!D2:D,' +
    workerSheetRef + '!F2:F,' +
    workerSheetRef + '!B2:B,' +
    workerSheetRef + '!C2:C' +
    "}," +
    workerSheetRef + '!B2:B=$B$4,' +
    "IF($E$4=\"\"," + workerSheetRef + '!A2:A<>"",' + workerSheetRef + '!C2:C=TEXT($E$4,\"yyyy-mm-dd\"))' +
    "),\"select Col1, max(Col3) where Col3 is not null group by Col1 pivot Col2 label Col1 '검사일', max(Col3) ''\",0),\"\"))";

  sheet.clear();
  ensureSheetSize_(sheet, 320, 35);

  sheet.getRange("A1:L1").merge().setValue("척도 검사 결과 대시보드");
  sheet.getRange("A2:L2").merge().setValue("대상자명과 생년월일을 입력하면 검사 이력과 척도별 점수 변화 그래프를 확인할 수 있습니다.");
  sheet.getRange("A4").setValue("대상자명");
  sheet.getRange("D4").setValue("생년월일");
  sheet.getRange("G4:L4").merge().setValue("동명이인은 생년월일을 함께 입력하면 더 정확하게 조회됩니다.");
  sheet.getRange("A6:C7").merge().setFormula('=IF($B$4="","검사 건수"&CHAR(10)&"-","검사 건수"&CHAR(10)&IFERROR(COUNTA(FILTER(' + workerSheetRef + '!A2:A,' + workerSheetRef + '!B2:B=$B$4,IF($E$4="",'+ workerSheetRef + '!A2:A<>"",' + workerSheetRef + '!C2:C=TEXT($E$4,"yyyy-mm-dd"))))&"건","0건"))');
  sheet.getRange("D6:F7").merge().setFormula('=IF($B$4="","최근 검사일"&CHAR(10)&"-","최근 검사일"&CHAR(10)&IFERROR(INDEX(SORT(FILTER(' + workerSheetRef + '!A2:A,' + workerSheetRef + '!B2:B=$B$4,IF($E$4="",'+ workerSheetRef + '!A2:A<>"",' + workerSheetRef + '!C2:C=TEXT($E$4,"yyyy-mm-dd"))),1,FALSE),1,1),"-"))');
  sheet.getRange("G6:I7").merge().setFormula('=IF($B$4="","평균 정규화점수"&CHAR(10)&"-","평균 정규화점수"&CHAR(10)&IFERROR(TEXT(ROUND(AVERAGE(FILTER(' + workerSheetRef + '!F2:F,' + workerSheetRef + '!B2:B=$B$4,IF($E$4="",'+ workerSheetRef + '!A2:A<>"",' + workerSheetRef + '!C2:C=TEXT($E$4,"yyyy-mm-dd")),' + workerSheetRef + '!F2:F<>"")),1),"0.0"),"-"))');
  sheet.getRange("J6:L7").merge().setFormula('=IF($B$4="","경고 건수"&CHAR(10)&"-","경고 건수"&CHAR(10)&IFERROR(COUNTIF(FILTER(' + workerSheetRef + '!K2:K,' + workerSheetRef + '!B2:B=$B$4,IF($E$4="",'+ workerSheetRef + '!A2:A<>"",' + workerSheetRef + '!C2:C=TEXT($E$4,"yyyy-mm-dd"))),"<>")&"건","0건"))');
  sheet.getRange("A10:J10").setValues([["검사일", "척도", "원점수", "정규화점수", "점수표시", "결과구간", "담당자", "비고", "경고여부", "기록고유값"]]);
  sheet.getRange("N10").setValue("점수 변화 그래프 데이터");
  sheet.getRange(columnToLetterScale_(dashboardConfig.namesHelperColumn) + "1").setValue("대상자 목록");

  sheet.getRange(dashboardConfig.clientNameCell).clearDataValidations();
  sheet.getRange(dashboardConfig.birthDateCell).setNumberFormat("yyyy-mm-dd");
  sheet.getRange(dashboardConfig.clientNameCell).setBackground("#ffffff");
  sheet.getRange(dashboardConfig.birthDateCell).setBackground("#ffffff");
  sheet.getRange("A11").setFormula(detailFormula);
  sheet.getRange(columnToLetterScale_(dashboardConfig.trendStartColumn) + String(dashboardConfig.trendStartRow)).setFormula(trendFormula);
  sheet.getRange(columnToLetterScale_(dashboardConfig.namesHelperColumn) + "2")
    .setFormula('=ARRAYFORMULA(SORT(UNIQUE(FILTER(' + workerSheetRef + '!B2:B,' + workerSheetRef + '!B2:B<>""))))');

  const nameValidation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sheet.getRange(columnToLetterScale_(dashboardConfig.namesHelperColumn) + "2:" + columnToLetterScale_(dashboardConfig.namesHelperColumn)), true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(dashboardConfig.clientNameCell).setDataValidation(nameValidation);

  sheet.setFrozenRows(10);
  sheet.hideColumns(dashboardConfig.namesHelperColumn, 35 - dashboardConfig.namesHelperColumn + 1);
  sheet.getRange("A1:L1").setFontSize(18).setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("left");
  sheet.getRange("A2:L2").setFontColor("#4f5b52");
  sheet.getRange("A4:F4").setFontWeight("bold");
  sheet.getRange("A4:L4").setBorder(true, true, true, true, true, true, "#d9e2dc", SpreadsheetApp.BorderStyle.SOLID);
  styleHeaderRow_(sheet, 10, 10);
  sheet.getRange("N10:Z10").setFontWeight("bold").setBackground("#fff2cc");
  styleScaleDashboardCards_(sheet);
  applyScaleBanding_(sheet.getRange(10, 1, Math.max(sheet.getMaxRows() - 9, 2), 10));
  applyScaleDashboardRules_(sheet);
  setScaleColumnWidths_(sheet, [95, 140, 85, 105, 120, 105, 105, 180, 140, 160, 90, 90, 90, 95, 95, 95, 95, 95, 95, 95, 95, 95, 95]);
  sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("D:D").setNumberFormat("0.0");

  const charts = sheet.getCharts();
  charts.forEach(function(chart) {
    sheet.removeChart(chart);
  });

  const chartRange = sheet.getRange("N11:Z320");
  const chart = sheet.newChart()
    .asLineChart()
    .addRange(chartRange)
    .setNumHeaders(1)
    .setOption("title", "검사일별 척도 변화")
    .setOption("legend", { position: "bottom", textStyle: { fontSize: 10 } })
    .setOption("hAxis", { title: "검사일", slantedText: true, slantedTextAngle: 30 })
    .setOption("vAxis", { title: "정규화 점수", viewWindow: { min: 0, max: 100 } })
    .setOption("curveType", "function")
    .setOption("lineWidth", 3)
    .setOption("pointSize", 6)
    .setOption("chartArea", { left: 70, top: 50, width: "72%", height: "62%" })
    .setOption("backgroundColor", "#ffffff")
    .setOption("series", {
      0: { color: "#3b82f6" },
      1: { color: "#f59e0b" },
      2: { color: "#10b981" },
      3: { color: "#ef4444" },
      4: { color: "#8b5cf6" },
      5: { color: "#14b8a6" },
      6: { color: "#ec4899" },
      7: { color: "#f97316" },
      8: { color: "#64748b" }
    })
    .setPosition(dashboardConfig.chartAnchorRow, dashboardConfig.chartAnchorColumn, 0, 0)
    .build();
  sheet.insertChart(chart);
}

function setScaleColumnWidths_(sheet, widths) {
  widths.forEach(function(width, index) {
    if (Number(width) > 0) {
      sheet.setColumnWidth(index + 1, Number(width));
    }
  });
}

function applyScaleBanding_(range) {
  const bandings = range.getSheet().getBandings();
  bandings.forEach(function(banding) {
    banding.remove();
  });
  range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
}

function styleScaleDashboardCards_(sheet) {
  [
    { range: "A6:C7", background: "#e8f3ed", fontColor: "#1b4332" },
    { range: "D6:F7", background: "#eef4ff", fontColor: "#1d4ed8" },
    { range: "G6:I7", background: "#fff6db", fontColor: "#9a6700" },
    { range: "J6:L7", background: "#fdecec", fontColor: "#b42318" }
  ].forEach(function(card) {
    sheet.getRange(card.range)
      .setBackground(card.background)
      .setFontColor(card.fontColor)
      .setFontWeight("bold")
      .setFontSize(12)
      .setWrap(true)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, false, false, "#d9e2dc", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
}

function applyScaleWorkerViewRules_(sheet) {
  const rules = [];
  const normalizedRange = sheet.getRange("F2:F");
  const warningRange = sheet.getRange("K2:K");

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(80)
    .setBackground("#fdecec")
    .setFontColor("#b42318")
    .setRanges([normalizedRange])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(60, 79.9999)
    .setBackground("#fff4e5")
    .setFontColor("#b54708")
    .setRanges([normalizedRange])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(30, 59.9999)
    .setBackground("#fff9db")
    .setFontColor("#9a6700")
    .setRanges([normalizedRange])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(30)
    .setBackground("#ecfdf3")
    .setFontColor("#027a48")
    .setRanges([normalizedRange])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=LEN($K2)>0')
    .setBackground("#fdecec")
    .setFontColor("#b42318")
    .setRanges([warningRange])
    .build());

  sheet.setConditionalFormatRules(rules);
}

function applyScaleRiskViewRules_(sheet) {
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LEN($H2)>0')
      .setBackground("#fdecec")
      .setFontColor("#b42318")
      .setRanges([sheet.getRange("A2:I")])
      .build()
  ];
  sheet.setConditionalFormatRules(rules);
}

function applyScaleDashboardRules_(sheet) {
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(80)
      .setBackground("#fdecec")
      .setFontColor("#b42318")
      .setRanges([sheet.getRange("D12:D")])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(60, 79.9999)
      .setBackground("#fff4e5")
      .setFontColor("#b54708")
      .setRanges([sheet.getRange("D12:D")])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(30, 59.9999)
      .setBackground("#fff9db")
      .setFontColor("#9a6700")
      .setRanges([sheet.getRange("D12:D")])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(30)
      .setBackground("#ecfdf3")
      .setFontColor("#027a48")
      .setRanges([sheet.getRange("D12:D")])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LEN($I12)>0')
      .setBackground("#fdecec")
      .setFontColor("#b42318")
      .setRanges([sheet.getRange("A12:J")])
      .build()
  ];
  sheet.setConditionalFormatRules(rules);
}

function escapeScaleSheetNameForFormula_(sheetName) {
  return "'" + String(sheetName || "").replace(/'/g, "''") + "'";
}

function columnToLetterScale_(columnNumber) {
  let number = Number(columnNumber) || 1;
  let result = "";
  while (number > 0) {
    const remainder = (number - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    number = Math.floor((number - 1) / 26);
  }
  return result;
}

function getExistingRowNumberMap_(sheet, headers, keyField) {
  const keyIndex = headers.indexOf(keyField);
  const result = {};

  if (keyIndex === -1 || sheet.getLastRow() < 2) {
    return result;
  }

  const values = sheet.getRange(2, keyIndex + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  values.forEach(function(row, index) {
    const key = normalizeText_(row[0]);
    if (key) {
      result[key] = index + 2;
    }
  });

  return result;
}

function ensureScaleScreeningSyncSheet_(sheetName, headers) {
  const sheet = getOrCreateScaleScreeningSheet_(sheetName);
  const normalizedHeaders = headers.map(function(header) {
    return normalizeText_(header);
  });

  if (sheet.getLastRow() === 0 || (sheet.getLastRow() === 1 && isHeaderRowBlank_(sheet, normalizedHeaders.length))) {
    ensureSheetSize_(sheet, 2, normalizedHeaders.length);
    sheet.getRange(1, 1, 1, normalizedHeaders.length).setValues([normalizedHeaders]);
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, normalizedHeaders.length).getDisplayValues()[0].map(function(value) {
      return normalizeText_(value);
    });
    if (currentHeaders.join("||") !== normalizedHeaders.join("||")) {
      console.warn("'" + sheetName + "' 시트 헤더를 최신 구조로 자동 보정합니다.");
      ensureSheetSize_(sheet, Math.max(sheet.getMaxRows(), 2), normalizedHeaders.length);
      sheet.getRange(1, 1, 1, normalizedHeaders.length).setValues([normalizedHeaders]);
    }
  }

  formatScaleScreeningSyncSheet_(sheet, normalizedHeaders.length);
  return sheet;
}

function getScaleScreeningTargetSpreadsheet_() {
  const activeSpreadsheet = tryGetActiveSpreadsheet_();
  const configuredId = normalizeText_(PropertiesService.getScriptProperties().getProperty(
    SCALE_SCREENING_SYNC_CONFIG.propertyKeys.spreadsheetId
  ));
  const defaultId = SCALE_SCREENING_SYNC_CONFIG.defaults.targetSpreadsheetId;

  if (configuredId) {
    try {
      return SpreadsheetApp.openById(configuredId);
    } catch (error) {
      if (activeSpreadsheet && activeSpreadsheet.getId() === configuredId) {
        return activeSpreadsheet;
      }
      throw new Error(
        "저장된 대상 스프레드시트에 접근할 수 없습니다. '현재 시트를 척도검사 DB로 지정' 메뉴를 다시 실행하거나, 실행 계정에 시트 권한을 부여해주세요."
      );
    }
  }

  if (activeSpreadsheet) {
    return activeSpreadsheet;
  }

  if (defaultId) {
    try {
      return SpreadsheetApp.openById(defaultId);
    } catch (error) {
      throw new Error(
        "기본 대상 스프레드시트에 접근할 수 없습니다. 대상 시트에 바운드된 Apps Script에서 '현재 시트를 척도검사 DB로 지정'을 먼저 실행해주세요."
      );
    }
  }

  throw new Error("대상 스프레드시트를 확인할 수 없습니다.");
}

function getScaleScreeningSheetIfExists_(sheetName) {
  return getScaleScreeningTargetSpreadsheet_().getSheetByName(sheetName);
}

function getOrCreateScaleScreeningSheet_(sheetName) {
  const spreadsheet = getScaleScreeningTargetSpreadsheet_();
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function tryGetActiveSpreadsheet_() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (error) {
    return null;
  }
}

function isHeaderRowBlank_(sheet, headerLength) {
  if (!sheet || sheet.getLastRow() < 1) {
    return true;
  }
  const values = sheet.getRange(1, 1, 1, headerLength).getDisplayValues()[0];
  return values.every(function(value) {
    return !normalizeText_(value);
  });
}

function formatScaleScreeningSyncSheet_(sheet, columnCount) {
  sheet.setFrozenRows(1);
  styleHeaderRow_(sheet, 1, columnCount);
  sheet.autoResizeColumns(1, Math.min(columnCount, sheet.getMaxColumns()));
}

function applyScaleScreeningDisplayFormats_() {
  const formatTargets = [
    { name: getScaleScreeningRecordSheetName_(), formats: { "H:H": "yyyy-mm-dd hh:mm", "I:I": "yyyy-mm-dd", "N:N": "0.0", "R:R": "yyyy-mm-dd" } },
    { name: getScaleScreeningAnswerSheetName_(), formats: { "C:C": "yyyy-mm-dd hh:mm", "D:D": "yyyy-mm-dd", "I:I": "yyyy-mm-dd" } },
    { name: getScaleScreeningWorkerViewSheetName_(), formats: { "A:A": "yyyy-mm-dd", "C:C": "yyyy-mm-dd", "F:F": "0.0" } },
    { name: getScaleScreeningRiskViewSheetName_(), formats: { "A:A": "yyyy-mm-dd", "C:C": "yyyy-mm-dd" } },
    { name: getScaleScreeningDashboardSheetName_(), formats: { "A:A": "yyyy-mm-dd", "D:D": "0.0", "E4": "yyyy-mm-dd" } }
  ];

  formatTargets.forEach(function(target) {
    const sheet = getScaleScreeningSheetIfExists_(target.name);
    if (!sheet) {
      return;
    }

    Object.keys(target.formats).forEach(function(rangeA1) {
      sheet.getRange(rangeA1).setNumberFormat(target.formats[rangeA1]);
    });
  });
}

function getScaleScreeningSyncToken_() {
  return normalizeText_(PropertiesService.getScriptProperties().getProperty(
    SCALE_SCREENING_SYNC_CONFIG.propertyKeys.token
  ));
}

function getScaleScreeningRecordSheetName_() {
  return normalizeText_(PropertiesService.getScriptProperties().getProperty(
    SCALE_SCREENING_SYNC_CONFIG.propertyKeys.recordSheetName
  )) || SCALE_SCREENING_SYNC_CONFIG.defaults.recordSheetName;
}

function getScaleScreeningAnswerSheetName_() {
  return normalizeText_(PropertiesService.getScriptProperties().getProperty(
    SCALE_SCREENING_SYNC_CONFIG.propertyKeys.answerSheetName
  )) || SCALE_SCREENING_SYNC_CONFIG.defaults.answerSheetName;
}

function getScaleScreeningQuestionnaireSheetName_() {
  return normalizeText_(PropertiesService.getScriptProperties().getProperty(
    SCALE_SCREENING_SYNC_CONFIG.propertyKeys.questionnaireSheetName
  )) || SCALE_SCREENING_SYNC_CONFIG.defaults.questionnaireSheetName;
}

function getScaleScreeningFieldSheetName_() {
  return normalizeText_(PropertiesService.getScriptProperties().getProperty(
    SCALE_SCREENING_SYNC_CONFIG.propertyKeys.fieldSheetName
  )) || SCALE_SCREENING_SYNC_CONFIG.defaults.fieldSheetName;
}

function getScaleScreeningOptionSheetName_() {
  return normalizeText_(PropertiesService.getScriptProperties().getProperty(
    SCALE_SCREENING_SYNC_CONFIG.propertyKeys.optionSheetName
  )) || SCALE_SCREENING_SYNC_CONFIG.defaults.optionSheetName;
}

function getScaleScreeningWorkerViewSheetName_() {
  return normalizeText_(PropertiesService.getScriptProperties().getProperty(
    SCALE_SCREENING_SYNC_CONFIG.propertyKeys.workerViewSheetName
  )) || SCALE_SCREENING_SYNC_CONFIG.defaults.workerViewSheetName;
}

function getScaleScreeningRiskViewSheetName_() {
  return normalizeText_(PropertiesService.getScriptProperties().getProperty(
    SCALE_SCREENING_SYNC_CONFIG.propertyKeys.riskViewSheetName
  )) || SCALE_SCREENING_SYNC_CONFIG.defaults.riskViewSheetName;
}

function getScaleScreeningDashboardSheetName_() {
  return normalizeText_(PropertiesService.getScriptProperties().getProperty(
    SCALE_SCREENING_SYNC_CONFIG.propertyKeys.dashboardSheetName
  )) || SCALE_SCREENING_SYNC_CONFIG.defaults.dashboardSheetName;
}

function getScaleScreeningSettingsSheetName_() {
  return normalizeText_(PropertiesService.getScriptProperties().getProperty(
    SCALE_SCREENING_SYNC_CONFIG.propertyKeys.settingsSheetName
  )) || SCALE_SCREENING_SYNC_CONFIG.defaults.settingsSheetName;
}

function createScaleScreeningJsonOutput_(data, callbackName) {
  const body = JSON.stringify(data);
  const normalizedCallback = normalizeScaleJsonpCallback_(callbackName);

  if (normalizedCallback) {
    return ContentService
      .createTextOutput(normalizedCallback + "(" + body + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeScaleJsonpCallback_(callbackName) {
  const text = normalizeText_(callbackName);
  if (!text) {
    return "";
  }

  return /^[A-Za-z_$][0-9A-Za-z_$]*(?:\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(text) ? text : "";
}

function escapeHtmlScale_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function findDisplayValueByLabel_(items, label) {
  const matched = (items || []).filter(function(item) {
    return item && item.label === label;
  })[0];
  return matched ? toCellText_(matched.value) : "";
}

function buildScaleScreeningProgressSummary_(progress) {
  if (!progress || progress.percent === null || progress.percent === undefined) {
    return "";
  }

  const answered = progress.answered === null || progress.answered === undefined ? "" : String(progress.answered);
  const total = progress.total === null || progress.total === undefined ? "" : String(progress.total);
  return String(progress.percent) + "% (" + answered + "/" + total + "항목)";
}

function buildRespondentSummary_(items) {
  return (items || []).map(function(item) {
    if (!item || !item.label || !item.value) {
      return "";
    }
    return item.label + ": " + item.value;
  }).filter(Boolean).join(" | ");
}

function buildBreakdownSummary_(items) {
  return (items || []).map(function(item) {
    if (!item) {
      return "";
    }

    const baseText = [
      toCellText_(item.number),
      toCellText_(item.text),
      "=>",
      toCellText_(item.answerLabel),
      item.score === null || item.score === undefined ? "" : "(" + item.score + "점)"
    ].filter(Boolean).join(" ");

    const subText = (item.subAnswers || []).map(function(subItem) {
      return [
        toCellText_(subItem.number),
        toCellText_(subItem.text),
        "=>",
        toCellText_(subItem.answerLabel),
        subItem.score === null || subItem.score === undefined ? "" : "(" + subItem.score + "점)"
      ].filter(Boolean).join(" ");
    }).filter(Boolean).join(" / ");

    return [baseText, subText].filter(Boolean).join(" || ");
  }).filter(Boolean).join(" ### ");
}

function joinScaleTextList_(items) {
  return (items || []).map(function(item) {
    return toCellText_(item);
  }).filter(Boolean).join(" | ");
}

function safeStringifyScaleValue_(value) {
  try {
    return JSON.stringify(value);
  } catch (error) {
    return "";
  }
}

function toCellText_(value) {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "string") {
    return value;
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  return safeStringifyScaleValue_(value);
}
