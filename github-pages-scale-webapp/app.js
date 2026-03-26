(() => {
  "use strict";

  const ORG_NAME = "다시서기종합지원센터";
  const TEAM_NAME = "정신건강팀";
  const GOOGLE_SHEET_URL =
    "https://docs.google.com/spreadsheets/d/11y5p7Cp_yN2vggMOlCwn4pKNBEmio-CmkK25Nyd2nIk/edit?gid=0#gid=0";
  const DEFAULT_GOOGLE_SYNC_URL =
    "https://script.google.com/macros/s/AKfycbywbENzL--pd_pLcmJPzWvAKOhzM7SwAkL38Zd7aNldafguQO85N_U2k0v5baUxhr4E/exec";
  const STORAGE_KEYS = {
    records: "mindmap_scale_records_v1",
    worker: "mindmap_scale_worker_v1",
    googleSync: "mindmap_scale_google_sync_v2"
  };
  const SCALE_COLORS = [
    "#126b57",
    "#f59e0b",
    "#3b82f6",
    "#ef4444",
    "#8b5cf6",
    "#14b8a6",
    "#f97316",
    "#0ea5e9",
    "#65a30d"
  ];

  const state = {
    manifest: [],
    questionnaires: new Map(),
    currentQuestionnaire: null,
    currentQuestionnaireId: null,
    activeView: "screening",
    lastResult: null,
    records: [],
    chart: null,
    remoteDashboardRecords: [],
    syncSettings: {
      webAppUrl: DEFAULT_GOOGLE_SYNC_URL,
      syncToken: "",
      syncEnabled: false
    }
  };

  const ui = {};

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    cacheUi();
    bindEvents();
    loadBundle();
    loadSavedRecords();
    loadSyncSettings();
    restoreWorkerName();
    if (!ui.sessionDate.value) {
      ui.sessionDate.value = new Date().toISOString().slice(0, 10);
    }
    renderScaleFilters();
    renderQuestionnaireNav();
    setCurrentQuestionnaire(state.manifest[0]?.id || null);
    renderRecordsTable();
    renderDashboard();
    updateSummary();
    setHeroStatus("준비가 끝났습니다. 결과는 현재 브라우저에 저장됩니다.");
  }

  function cacheUi() {
    ui.summaryRecordCount = document.getElementById("summaryRecordCount");
    ui.summarySubjectCount = document.getElementById("summarySubjectCount");
    ui.summaryScaleCount = document.getElementById("summaryScaleCount");
    ui.heroStatusText = document.getElementById("heroStatusText");
    ui.scaleSearchInput = document.getElementById("scaleSearchInput");
    ui.questionnaireCount = document.getElementById("questionnaireCount");
    ui.questionnaireNav = document.getElementById("questionnaireNav");
    ui.tabs = Array.from(document.querySelectorAll(".view-tab"));
    ui.panes = Array.from(document.querySelectorAll(".view-pane"));
    ui.scaleBadge = document.getElementById("scaleBadge");
    ui.scaleTitle = document.getElementById("scaleTitle");
    ui.scaleMeta = document.getElementById("scaleMeta");
    ui.scaleSource = document.getElementById("scaleSource");
    ui.questionPrompt = document.getElementById("questionPrompt");
    ui.questionPromptTitle = document.getElementById("questionPromptTitle");
    ui.progressLabel = document.getElementById("progressLabel");
    ui.progressFill = document.getElementById("progressFill");
    ui.screeningForm = document.getElementById("screeningForm");
    ui.sessionDate = document.getElementById("sessionDate");
    ui.workerName = document.getElementById("workerName");
    ui.clientLabel = document.getElementById("clientLabel");
    ui.birthDate = document.getElementById("birthDate");
    ui.sessionNote = document.getElementById("sessionNote");
    ui.subjectSuggestions = document.getElementById("subjectSuggestions");
    ui.respondentFields = document.getElementById("respondentFields");
    ui.questionsRoot = document.getElementById("questionsRoot");
    ui.resetBtn = document.getElementById("resetBtn");
    ui.resultEmpty = document.getElementById("resultEmpty");
    ui.resultPanel = document.getElementById("resultPanel");
    ui.printOrgLine = document.getElementById("printOrgLine");
    ui.printScaleLine = document.getElementById("printScaleLine");
    ui.printMetaLine = document.getElementById("printMetaLine");
    ui.resultTitle = document.getElementById("resultTitle");
    ui.resultScore = document.getElementById("resultScore");
    ui.resultBand = document.getElementById("resultBand");
    ui.resultFlags = document.getElementById("resultFlags");
    ui.resultHighlights = document.getElementById("resultHighlights");
    ui.resultMeta = document.getElementById("resultMeta");
    ui.resultBreakdown = document.getElementById("resultBreakdown");
    ui.resultNotes = document.getElementById("resultNotes");
    ui.saveResultBtn = document.getElementById("saveResultBtn");
    ui.printBtn = document.getElementById("printBtn");
    ui.exportCurrentBtn = document.getElementById("exportCurrentBtn");
    ui.recordsSearchInput = document.getElementById("recordsSearchInput");
    ui.recordsScaleFilter = document.getElementById("recordsScaleFilter");
    ui.recordsStatusText = document.getElementById("recordsStatusText");
    ui.recordsTableBody = document.getElementById("recordsTableBody");
    ui.recordsEmpty = document.getElementById("recordsEmpty");
    ui.exportAllJsonBtn = document.getElementById("exportAllJsonBtn");
    ui.exportAllCsvBtn = document.getElementById("exportAllCsvBtn");
    ui.syncAllRecordsBtn = document.getElementById("syncAllRecordsBtn");
    ui.importJsonBtn = document.getElementById("importJsonBtn");
    ui.importJsonInput = document.getElementById("importJsonInput");
    ui.clearAllBtn = document.getElementById("clearAllBtn");
    ui.dashboardNameInput = document.getElementById("dashboardNameInput");
    ui.dashboardBirthInput = document.getElementById("dashboardBirthInput");
    ui.dashboardScaleFilter = document.getElementById("dashboardScaleFilter");
    ui.dashboardSubjectSuggestions = document.getElementById("dashboardSubjectSuggestions");
    ui.dashboardRefreshBtn = document.getElementById("dashboardRefreshBtn");
    ui.dashboardStatusText = document.getElementById("dashboardStatusText");
    ui.dashboardSummary = document.getElementById("dashboardSummary");
    ui.dashboardChart = document.getElementById("dashboardChart");
    ui.dashboardTableBody = document.getElementById("dashboardTableBody");
    ui.dashboardEmpty = document.getElementById("dashboardEmpty");
    ui.googleSyncUrl = document.getElementById("googleSyncUrl");
    ui.googleSyncToken = document.getElementById("googleSyncToken");
    ui.googleSyncEnabled = document.getElementById("googleSyncEnabled");
    ui.openAuthorizeBtn = document.getElementById("openAuthorizeBtn");
    ui.openSheetBtn = document.getElementById("openSheetBtn");
    ui.syncCurrentBtn = document.getElementById("syncCurrentBtn");
    ui.syncHistoryBtn = document.getElementById("syncHistoryBtn");
    ui.syncQuestionnairesBtn = document.getElementById("syncQuestionnairesBtn");
    ui.checkSyncStatusBtn = document.getElementById("checkSyncStatusBtn");
    ui.syncStatusText = document.getElementById("syncStatusText");
  }

  function bindEvents() {
    const debouncedScaleSearch = debounce(() => renderQuestionnaireNav(ui.scaleSearchInput.value), 90);
    const debouncedRecordsRender = debounce(() => renderRecordsTable(), 120);
    const debouncedDashboardCriteria = debounce(() => onDashboardCriteriaChanged(), 180);

    ui.scaleSearchInput.addEventListener("input", debouncedScaleSearch);
    ui.tabs.forEach((tab) => {
      tab.addEventListener("click", () => setActiveView(tab.dataset.view));
    });
    ui.questionnaireNav.addEventListener("click", (event) => {
      const button = event.target.closest("[data-questionnaire-id]");
      if (button) {
        setCurrentQuestionnaire(button.dataset.questionnaireId);
      }
    });
    ui.screeningForm.addEventListener("submit", onCalculate);
    ui.screeningForm.addEventListener("change", onFormChanged);
    ui.screeningForm.addEventListener("input", onFormInput);
    ui.resetBtn.addEventListener("click", onResetCurrentForm);
    ui.saveResultBtn.addEventListener("click", onSaveResult);
    ui.printBtn.addEventListener("click", () => window.print());
    ui.exportCurrentBtn.addEventListener("click", onExportCurrentRecord);
    ui.workerName.addEventListener("input", () => {
      localStorage.setItem(STORAGE_KEYS.worker, ui.workerName.value.trim());
    });
    ui.recordsSearchInput.addEventListener("input", debouncedRecordsRender);
    ui.recordsScaleFilter.addEventListener("change", renderRecordsTable);
    ui.recordsTableBody.addEventListener("click", onRecordsTableClick);
    ui.exportAllJsonBtn.addEventListener("click", exportAllRecordsAsJson);
    ui.exportAllCsvBtn.addEventListener("click", exportAllRecordsAsCsv);
    ui.syncAllRecordsBtn.addEventListener("click", onSyncAllRecords);
    ui.importJsonBtn.addEventListener("click", () => ui.importJsonInput.click());
    ui.importJsonInput.addEventListener("change", onImportJson);
    ui.clearAllBtn.addEventListener("click", onClearAllRecords);
    ui.dashboardNameInput.addEventListener("input", debouncedDashboardCriteria);
    ui.dashboardBirthInput.addEventListener("change", onDashboardCriteriaChanged);
    ui.dashboardScaleFilter.addEventListener("change", onDashboardCriteriaChanged);
    ui.dashboardRefreshBtn.addEventListener("click", onDashboardRefresh);
    ui.googleSyncUrl.addEventListener("input", onGoogleSyncSettingsInput);
    ui.googleSyncToken.addEventListener("input", onGoogleSyncSettingsInput);
    ui.googleSyncEnabled.addEventListener("change", onGoogleSyncSettingsInput);
    ui.openAuthorizeBtn.addEventListener("click", onOpenAuthorizePage);
    ui.openSheetBtn.addEventListener("click", () => window.open(GOOGLE_SHEET_URL, "_blank", "noopener"));
    ui.syncCurrentBtn.addEventListener("click", onSyncCurrentResult);
    ui.syncHistoryBtn.addEventListener("click", onSyncAllRecords);
    ui.syncQuestionnairesBtn.addEventListener("click", onSyncQuestionnaires);
    ui.checkSyncStatusBtn.addEventListener("click", onCheckSyncStatus);
  }

  function loadBundle() {
    const bundle = window.__SCALE_SCREENING_BUNDLE__;
    if (!bundle?.manifest || !bundle?.questionnaires) {
      throw new Error("척도 데이터 번들을 불러오지 못했습니다.");
    }

    state.manifest = [...bundle.manifest].sort((a, b) => {
      const qa = bundle.questionnaires[a.id];
      const qb = bundle.questionnaires[b.id];
      return Number(qa?.selfSeq || 999) - Number(qb?.selfSeq || 999);
    });
    state.questionnaires = new Map(Object.entries(bundle.questionnaires));
    ui.summaryScaleCount.textContent = String(state.manifest.length);
    ui.questionnaireCount.textContent = `${state.manifest.length}개`;
  }

  function loadSavedRecords() {
    try {
      const parsed = JSON.parse(localStorage.getItem(STORAGE_KEYS.records) || "[]");
      state.records = Array.isArray(parsed) ? parsed.filter(Boolean) : [];
    } catch (error) {
      console.warn("저장 결과를 읽지 못했습니다.", error);
      state.records = [];
    }
  }

  function persistRecords() {
    localStorage.setItem(STORAGE_KEYS.records, JSON.stringify(state.records.slice(0, 1000)));
    updateSummary();
    populateSubjectSuggestions();
    renderRecordsTable();
    renderDashboard();
  }

  function restoreWorkerName() {
    ui.workerName.value = localStorage.getItem(STORAGE_KEYS.worker) || "";
  }

  function loadSyncSettings() {
    state.syncSettings = loadStoredSyncSettings();
    ui.googleSyncUrl.value = state.syncSettings.webAppUrl || "";
    ui.googleSyncToken.value = state.syncSettings.syncToken || "";
    ui.googleSyncEnabled.checked = Boolean(state.syncSettings.syncEnabled);
    syncSheetControls();
  }

  function loadStoredSyncSettings() {
    try {
      const parsed = JSON.parse(localStorage.getItem(STORAGE_KEYS.googleSync) || "{}");
      return {
        webAppUrl: normalizeGoogleSyncUrl(parsed.webAppUrl || DEFAULT_GOOGLE_SYNC_URL),
        syncToken: typeof parsed.syncToken === "string" ? parsed.syncToken : "",
        syncEnabled: Boolean(parsed.syncEnabled)
      };
    } catch (error) {
      console.warn("구글 시트 연동 설정을 읽지 못했습니다.", error);
      return {
        webAppUrl: DEFAULT_GOOGLE_SYNC_URL,
        syncToken: "",
        syncEnabled: false
      };
    }
  }

  function onGoogleSyncSettingsInput() {
    state.syncSettings = {
      webAppUrl: normalizeGoogleSyncUrl(ui.googleSyncUrl.value),
      syncToken: ui.googleSyncToken.value.trim(),
      syncEnabled: Boolean(ui.googleSyncEnabled.checked)
    };
    localStorage.setItem(STORAGE_KEYS.googleSync, JSON.stringify(state.syncSettings));
    syncSheetControls();

    if (!state.syncSettings.webAppUrl) {
      setSyncStatus("구글 시트 연동 주소를 입력하면 승인 사용자만 시트 저장과 조회를 사용할 수 있습니다.");
      return;
    }

    if (!isValidGoogleSyncUrlFormat(state.syncSettings.webAppUrl)) {
      setSyncStatus("구글 연동 주소 형식이 올바르지 않습니다.", "error");
      return;
    }

    if (!state.syncSettings.syncEnabled) {
      setSyncStatus("설정은 저장되었습니다. 시트 기능 사용을 켜면 승인 사용자 전용 기능이 활성화됩니다.");
      return;
    }

    setSyncStatus("연동 설정을 저장했습니다. 권한 승인 후 시트 저장과 조회를 사용할 수 있습니다.");
  }

  function normalizeGoogleSyncUrl(value) {
    const normalized = typeof value === "string" ? value.trim() : "";
    if (!normalized) {
      return "";
    }

    if (isValidGoogleSyncUrlFormat(normalized)) {
      return normalized;
    }

    if (
      /^https:\/\/docs\.google\.com\/spreadsheets\//i.test(normalized) ||
      /^https:\/\/script\.google\.com\/.*\/edit(?:[?#].*)?$/i.test(normalized) ||
      /^https:\/\/script\.google\.com\/u\/\d+\/home\/projects\//i.test(normalized)
    ) {
      return "";
    }

    return normalized;
  }

  function isValidGoogleSyncUrlFormat(url) {
    return /^https:\/\/script\.google\.com\/(?:macros\/s\/[A-Za-z0-9_-]+(?:\/(?:exec|dev))?|a\/macros\/[^/]+\/s\/[A-Za-z0-9_-]+(?:\/(?:exec|dev))?)(?:[/?#].*)?$/i.test(url || "");
  }

  function normalizeLookupText(value) {
    return String(value || "")
      .trim()
      .replace(/\s+/g, " ")
      .toLowerCase();
  }

  function syncSheetControls() {
    const enabled = Boolean(state.syncSettings.syncEnabled);
    [
      ui.syncCurrentBtn,
      ui.syncHistoryBtn,
      ui.syncQuestionnairesBtn,
      ui.checkSyncStatusBtn,
      ui.dashboardRefreshBtn
    ].forEach((button) => {
      if (button) {
        button.disabled = !enabled;
      }
    });
    if (!enabled) {
      setDashboardStatus("시트 기능 사용이 꺼져 있습니다. 비교 분석은 현재 브라우저 저장 결과만 표시합니다.");
      setSyncStatus("시트 기능 사용이 꺼져 있습니다. 일반 사용자는 검사와 기기 저장만 사용할 수 있습니다.");
    }
  }

  function onOpenAuthorizePage() {
    const url = buildAuthorizeUrl(state.syncSettings.webAppUrl || DEFAULT_GOOGLE_SYNC_URL);
    if (!url) {
      setSyncStatus("구글 연동 주소가 올바르지 않아 권한 승인 페이지를 열 수 없습니다.", "error");
      ui.googleSyncUrl.focus();
      return;
    }
    window.open(url, "_blank", "noopener");
  }

  function buildAuthorizeUrl(webAppUrl) {
    const normalized = normalizeGoogleSyncUrl(webAppUrl);
    if (!normalized || !isValidGoogleSyncUrlFormat(normalized)) {
      return "";
    }
    const url = new URL(normalized);
    url.searchParams.set("action", "authorize");
    return url.toString();
  }

  function compareRecordDate(record) {
    return String(record?.meta?.sessionDate || record?.createdAt || "");
  }

  function validateGoogleSyncSettings(settings, options = {}) {
    if (options.requireEnabled !== false && !settings?.syncEnabled) {
      return {
        ok: false,
        message: "시트 기능 사용을 먼저 켜고, 권한 승인 페이지에서 Google 승인을 완료해주세요.",
        focusTarget: ui.googleSyncEnabled
      };
    }

    if (!settings?.webAppUrl) {
      return {
        ok: false,
        message: "구글 연동 주소를 먼저 입력해주세요.",
        focusTarget: ui.googleSyncUrl
      };
    }

    if (!isValidGoogleSyncUrlFormat(settings.webAppUrl)) {
      return {
        ok: false,
        message: "구글 연동 주소 형식이 올바르지 않습니다. 배포된 Apps Script 웹앱 주소를 사용해주세요.",
        focusTarget: ui.googleSyncUrl
      };
    }

    return { ok: true };
  }

  function setSyncStatus(message, type = "") {
    ui.syncStatusText.textContent = message;
    ui.syncStatusText.classList.remove("success", "error");
    if (type) {
      ui.syncStatusText.classList.add(type);
    }
  }

  function setDashboardStatus(message, type = "") {
    ui.dashboardStatusText.textContent = message;
    ui.dashboardStatusText.classList.remove("success", "error");
    if (type) {
      ui.dashboardStatusText.classList.add(type);
    }
  }

  function setSyncBusy(isBusy, message = "") {
    [
      ui.saveResultBtn,
      ui.syncAllRecordsBtn,
      ui.syncCurrentBtn,
      ui.syncHistoryBtn,
      ui.syncQuestionnairesBtn,
      ui.checkSyncStatusBtn,
      ui.dashboardRefreshBtn
    ].forEach((button) => {
      if (button) {
        button.disabled = isBusy;
      }
    });

    if (message) {
      setSyncStatus(message, isBusy ? "" : "success");
    }
  }

  function renderScaleFilters() {
    [ui.recordsScaleFilter, ui.dashboardScaleFilter].forEach((select) => {
      select.innerHTML = '<option value="">전체 척도</option>';
      state.manifest.forEach((item) => {
        const option = document.createElement("option");
        option.value = item.id;
        option.textContent = item.title;
        select.appendChild(option);
      });
    });
  }

  function renderQuestionnaireNav(filterText = "") {
    const filter = String(filterText || "").trim().toLowerCase();
    ui.questionnaireNav.innerHTML = "";
    const fragment = document.createDocumentFragment();
    const items = state.manifest.filter((item) => {
      if (!filter) {
        return true;
      }
      const haystack = `${item.title} ${item.id} ${state.questionnaires.get(item.id)?.shortTitle || ""}`.toLowerCase();
      return haystack.includes(filter);
    });

    items.forEach((item) => {
      const data = state.questionnaires.get(item.id);
      const button = document.createElement("button");
      button.type = "button";
      button.className = `nav-btn${item.id === state.currentQuestionnaireId ? " is-active" : ""}`;
      button.dataset.questionnaireId = item.id;
      button.innerHTML = `
        <span class="nav-title">${escapeHtml(item.title)}</span>
        <span class="nav-meta">${escapeHtml(data?.shortTitle || item.id)} · ${escapeHtml(item.recommendedAge || "연령 제한 없음")}</span>
      `;
      fragment.appendChild(button);
    });
    ui.questionnaireNav.appendChild(fragment);
  }

  function setCurrentQuestionnaire(id) {
    if (!id) {
      return;
    }
    const questionnaire = state.questionnaires.get(id);
    if (!questionnaire) {
      return;
    }

    const metaSnapshot = readMetaFields();
    state.currentQuestionnaireId = id;
    state.currentQuestionnaire = questionnaire;
    state.lastResult = null;
    renderQuestionnaireNav(ui.scaleSearchInput.value);
    renderQuestionnaire(questionnaire);
    writeMetaFields(metaSnapshot);
    clearResult();
    setActiveView("screening");
    refreshQuestionUi();
  }

  function renderQuestionnaire(questionnaire) {
    ui.scaleBadge.textContent = `${questionnaire.shortTitle || questionnaire.title} · ${(questionnaire.questions || []).length}문항`;
    ui.scaleTitle.textContent = questionnaire.title;
    ui.scaleMeta.textContent = [
      questionnaire.shortTitle,
      questionnaire.recommendedAge ? `권장연령 ${questionnaire.recommendedAge}` : "",
      `${(questionnaire.questions || []).length}문항`
    ]
      .filter(Boolean)
      .join(" · ");
    ui.scaleSource.textContent = questionnaire.source?.citation || questionnaire.source?.institution || "별도 출처 정보 없음";
    ui.questionPrompt.textContent = questionnaire.questionPrompt || "";
    renderRespondentFields(questionnaire);
    renderQuestions(questionnaire);
  }

  function renderRespondentFields(questionnaire) {
    ui.respondentFields.innerHTML = "";
    const fields = questionnaire.respondentFields || [];

    if (!fields.length) {
      ui.respondentFields.innerHTML = '<div class="empty-state">이 척도는 별도 응답자 정보 문항이 없습니다.</div>';
      return;
    }

    fields.forEach((field) => {
      const article = document.createElement("article");
      article.className = "respondent-card";
      article.innerHTML = `<p class="respondent-title">${escapeHtml(field.label)}${field.required ? " *" : ""}</p>`;
      const choices = document.createElement("div");
      choices.className = "choice-grid";

      (field.options || []).forEach((option, index) => {
        const optionId = `respondent_${field.id}_${index}`;
        const label = document.createElement("label");
        label.className = "choice-card";
        label.innerHTML = `
          <input type="radio" name="respondent_${escapeHtml(field.id)}" id="${escapeHtml(optionId)}" value="${escapeHtml(option.value)}" />
          <span class="choice-label"><span>${escapeHtml(option.label)}</span></span>
        `;
        choices.appendChild(label);
      });

      article.appendChild(choices);
      ui.respondentFields.appendChild(article);
    });
  }

  function renderQuestions(questionnaire) {
    ui.questionsRoot.innerHTML = "";
    let previousSection = "";

    (questionnaire.questions || []).forEach((question, questionIndex) => {
      if (question.sectionTitle && question.sectionTitle !== previousSection) {
        const heading = document.createElement("h4");
        heading.className = "question-section-label";
        heading.textContent = question.sectionTitle;
        ui.questionsRoot.appendChild(heading);
        previousSection = question.sectionTitle;
      }

      const article = document.createElement("article");
      article.className = "question-card";
      article.dataset.questionId = question.id;
      article.innerHTML = `
        <div class="question-card-head">
          <p class="question-number">${escapeHtml(question.number || `${questionIndex + 1}.`)}</p>
          <p class="question-text">${escapeHtml(question.text)}</p>
        </div>
      `;
      article.appendChild(buildChoiceGrid(question.options || [], question.id, true));

      if (question.subQuestions?.length) {
        const subWrap = document.createElement("div");
        subWrap.className = "subquestions hidden";
        question.subQuestions.forEach((subQuestion) => {
          const subBlock = document.createElement("div");
          subBlock.className = "respondent-card";
          subBlock.innerHTML = `<p class="respondent-title">${escapeHtml(subQuestion.label || subQuestion.text || "")}</p>`;
          subBlock.appendChild(buildChoiceGrid(subQuestion.options || [], subQuestion.id, false));
          subWrap.appendChild(subBlock);
        });
        article.appendChild(subWrap);
      }

      ui.questionsRoot.appendChild(article);
    });
  }

  function buildChoiceGrid(options, name, showScore) {
    const wrapper = document.createElement("div");
    wrapper.className = "choice-grid";

    options.forEach((option, index) => {
      const optionId = `${name}_${index}`;
      const label = document.createElement("label");
      label.className = "choice-card";
      label.innerHTML = `
        <input
          type="radio"
          name="${escapeHtml(name)}"
          id="${escapeHtml(optionId)}"
          value="${escapeHtml(String(option.value))}"
          data-score="${escapeHtml(String(Number(option.score || 0)))}"
        />
        <span class="choice-label">
          <span>${escapeHtml(option.label)}</span>
          ${showScore && Number.isFinite(Number(option.score)) ? `<span class="choice-score">점수 ${escapeHtml(String(Number(option.score)))}</span>` : ""}
        </span>
      `;
      wrapper.appendChild(label);
    });

    return wrapper;
  }

  function onFormChanged() {
    refreshQuestionUi();
  }

  function onFormInput(event) {
    if (event.target === ui.workerName) {
      localStorage.setItem(STORAGE_KEYS.worker, ui.workerName.value.trim());
    }
    refreshQuestionUi();
  }

  function refreshQuestionUi() {
    updateSelectionStates();
    updateSubQuestionVisibility();
    updateProgress();
  }

  function updateSelectionStates(root = document) {
    root.querySelectorAll(".choice-card").forEach((card) => {
      const input = card.querySelector("input[type='radio']");
      card.classList.toggle("is-selected", Boolean(input?.checked));
    });
  }

  function updateSubQuestionVisibility() {
    const questionnaire = state.currentQuestionnaire;
    if (!questionnaire) {
      return;
    }

    const snapshot = collectAnswers(questionnaire);
    (questionnaire.questions || []).forEach((question) => {
      if (!question.subQuestions?.length) {
        return;
      }

      const block = ui.questionsRoot.querySelector(`[data-question-id="${cssEscape(question.id)}"]`);
      const subWrap = block?.querySelector(".subquestions");
      if (!subWrap) {
        return;
      }

      const visible = shouldShowSubQuestions(question, questionnaire, snapshot.responses);
      subWrap.classList.toggle("hidden", !visible);

      if (!visible) {
        subWrap.querySelectorAll("input[type='radio']").forEach((input) => {
          input.checked = false;
        });
      }
    });

    updateSelectionStates(ui.questionsRoot);
  }

  function shouldShowSubQuestions(question, questionnaire, responses) {
    if (!question.subQuestions?.length) {
      return false;
    }
    if (questionnaire.scoring?.type === "k_mdq" && question.id === "q14") {
      return (questionnaire.scoring.symptomQuestionIds || []).some((id) => Number(responses[id]?.score || 0) > 0);
    }
    return Number(responses[question.id]?.score || 0) > 0;
  }

  function updateProgress() {
    const questionnaire = state.currentQuestionnaire;
    if (!questionnaire) {
      return;
    }
    const payload = collectAnswers(questionnaire);
    const progress = buildProgressSnapshot(questionnaire, payload);
    ui.progressLabel.textContent = `${progress.percent}%`;
    ui.progressFill.style.width = `${progress.percent}%`;
  }

  function onCalculate(event) {
    event.preventDefault();
    const questionnaire = state.currentQuestionnaire;
    const validation = validateResponses(questionnaire);

    if (!validation.ok) {
      alert(validation.message);
      validation.focusTarget?.focus();
      return;
    }

    const payload = collectAnswers(questionnaire);
    const evaluation = evaluateQuestionnaire(questionnaire, payload);
    const record = buildResultRecord(questionnaire, payload, evaluation);
    state.lastResult = record;
    renderResultRecord(record);
    setHeroStatus(`${questionnaire.title} 결과를 계산했습니다.`);
  }

  function validateResponses(questionnaire) {
    if (!ui.sessionDate.value) {
      return { ok: false, message: "검사일을 입력해주세요.", focusTarget: ui.sessionDate };
    }
    if (!ui.clientLabel.value.trim()) {
      return { ok: false, message: "대상자를 입력해주세요.", focusTarget: ui.clientLabel };
    }

    for (const field of questionnaire.respondentFields || []) {
      if (!field.required) {
        continue;
      }
      const checked = document.querySelector(`input[name="respondent_${cssEscape(field.id)}"]:checked`);
      if (!checked) {
        return {
          ok: false,
          message: `응답자 정보의 '${field.label}'을 선택해주세요.`,
          focusTarget: document.querySelector(`input[name="respondent_${cssEscape(field.id)}"]`)
        };
      }
    }

    const payload = collectAnswers(questionnaire);
    for (const question of questionnaire.questions || []) {
      const checked = document.querySelector(`input[name="${cssEscape(question.id)}"]:checked`);
      if (!checked) {
        return {
          ok: false,
          message: `${question.number || question.text} 문항에 답변해주세요.`,
          focusTarget: document.querySelector(`input[name="${cssEscape(question.id)}"]`)
        };
      }

      if (question.subQuestions?.length && shouldShowSubQuestions(question, questionnaire, payload.responses)) {
        for (const subQuestion of question.subQuestions) {
          const subChecked = document.querySelector(`input[name="${cssEscape(subQuestion.id)}"]:checked`);
          if (!subChecked) {
            return {
              ok: false,
              message: `${question.number || question.text} 문항의 추가 질문에 답변해주세요.`,
              focusTarget: document.querySelector(`input[name="${cssEscape(subQuestion.id)}"]`)
            };
          }
        }
      }
    }

    return { ok: true };
  }

  function collectAnswers(questionnaire) {
    const meta = {
      sessionDate: ui.sessionDate.value,
      workerName: ui.workerName.value.trim(),
      clientLabel: ui.clientLabel.value.trim(),
      birthDate: ui.birthDate.value,
      sessionNote: ui.sessionNote.value.trim()
    };

    const respondent = {};
    (questionnaire.respondentFields || []).forEach((field) => {
      const checked = document.querySelector(`input[name="respondent_${cssEscape(field.id)}"]:checked`);
      respondent[field.id] = checked ? checked.value : null;
    });

    const responses = {};
    (questionnaire.questions || []).forEach((question) => {
      const checked = document.querySelector(`input[name="${cssEscape(question.id)}"]:checked`);
      responses[question.id] = checked
        ? { value: checked.value, score: Number(checked.dataset.score || 0) }
        : null;

      (question.subQuestions || []).forEach((subQuestion) => {
        const subChecked = document.querySelector(`input[name="${cssEscape(subQuestion.id)}"]:checked`);
        responses[subQuestion.id] = subChecked
          ? { value: subChecked.value, score: Number(subChecked.dataset.score || 0) }
          : null;
      });
    });

    return { meta, respondent, responses };
  }

  function buildProgressSnapshot(questionnaire, payload) {
    const respondent = payload?.respondent || {};
    const responses = payload?.responses || payload?.answers || {};

    const respondentTotal = (questionnaire.respondentFields || []).filter((field) => field.required !== false).length;
    const respondentAnswered = (questionnaire.respondentFields || []).filter((field) => {
      if (field.required === false) {
        return false;
      }
      const value = respondent[field.id];
      return value !== null && value !== undefined && String(value).trim() !== "";
    }).length;

    let questionTotal = 0;
    let questionAnswered = 0;

    (questionnaire.questions || []).forEach((question) => {
      questionTotal += 1;
      if (responses[question.id]) {
        questionAnswered += 1;
      }

      if (question.subQuestions?.length && shouldShowSubQuestions(question, questionnaire, responses)) {
        question.subQuestions.forEach((subQuestion) => {
          questionTotal += 1;
          if (responses[subQuestion.id]) {
            questionAnswered += 1;
          }
        });
      }
    });

    const total = respondentTotal + questionTotal;
    const answered = respondentAnswered + questionAnswered;
    const percent = total ? Math.round((answered / total) * 100) : 0;

    return { answered, total, percent };
  }

  function formatProgressSummary(progress) {
    return progress ? `${progress.percent}% (${progress.answered}/${progress.total}항목)` : "";
  }

  function buildResultRecord(questionnaire, payload, evaluation) {
    return {
      id: createLocalId(),
      appVersion: "github-pages-v1",
      questionnaireId: questionnaire.id,
      questionnaireTitle: questionnaire.title,
      shortTitle: questionnaire.shortTitle || questionnaire.title,
      createdAt: new Date().toISOString(),
      meta: payload.meta,
      respondent: payload.respondent,
      respondentDisplay: buildRespondentDisplay(questionnaire, payload.respondent),
      answers: payload.responses,
      breakdown: buildAnswerBreakdown(questionnaire, payload.responses),
      progress: buildProgressSnapshot(questionnaire, payload),
      evaluation
    };
  }

  function buildRespondentDisplay(questionnaire, respondent) {
    return (questionnaire.respondentFields || []).map((field) => ({
      label: field.label,
      value: optionLabel(field.options || [], respondent[field.id]) || "미응답"
    }));
  }

  function buildAnswerBreakdown(questionnaire, responses) {
    return (questionnaire.questions || []).map((question) => {
      const answer = responses[question.id];
      const showSubAnswers = question.subQuestions?.length && shouldShowSubQuestions(question, questionnaire, responses);
      return {
        id: question.id,
        number: question.number,
        text: question.text,
        answerLabel: optionLabel(question.options || [], answer?.value) || "미응답",
        score: answer?.score ?? null,
        subAnswers: showSubAnswers
          ? (question.subQuestions || []).map((subQuestion) => {
              const subAnswer = responses[subQuestion.id];
              return {
                number: subQuestion.number || "",
                text: subQuestion.label || subQuestion.text || "",
                answerLabel: optionLabel(subQuestion.options || [], subAnswer?.value) || "미응답",
                score: subAnswer?.score ?? null
              };
            })
          : []
      };
    });
  }

  function optionLabel(options, value) {
    const found = (options || []).find((option) => String(option.value) === String(value));
    return found ? found.label : "";
  }

  function evaluateQuestionnaire(questionnaire, payload) {
    switch (questionnaire.scoring?.type) {
      case "cri":
        return evaluateCri(questionnaire, payload);
      case "k_mdq":
        return evaluateKmDQ(questionnaire, payload);
      case "endorsement_with_distress":
        return evaluateEndorsement(questionnaire, payload);
      default:
        return evaluateSum(questionnaire, payload);
    }
  }

  function evaluateSum(questionnaire, payload) {
    const score = (questionnaire.questions || []).reduce((sum, question) => sum + Number(payload.responses[question.id]?.score || 0), 0);
    const maxScore = (questionnaire.questions || []).reduce((sum, question) => sum + maxOptionScore(question.options || []), 0);
    const band = findBand(questionnaire.scoring?.bands || [], score);
    const flags = [];

    if (questionnaire.scoring?.suicideFlagQuestionId) {
      const suicideScore = payload.responses[questionnaire.scoring.suicideFlagQuestionId]?.score || 0;
      if (suicideScore > 0) {
        flags.push({
          level: "warn",
          text: "자해·자살사고 관련 문항에 응답이 있어 즉시 추가 확인이 필요합니다."
        });
      }
    }

    return {
      score,
      maxScore,
      normalizedScore: maxScore ? Math.round((score / maxScore) * 100) : 0,
      title: `${questionnaire.shortTitle || questionnaire.title} 결과`,
      scoreText: maxScore ? `${score}점 / ${maxScore}점` : `${score}점`,
      bandText: band?.label || "참고 구간 없음",
      highlights: [
        maxScore ? `총점 ${score}점 / ${maxScore}점` : `총점 ${score}점`,
        ...(band?.description ? [band.description] : []),
        ...(questionnaire.scoring?.highlights || [])
      ],
      flags,
      notes: questionnaire.scoring?.notes || []
    };
  }

  function evaluateCri(questionnaire, payload) {
    const responses = payload.responses || {};
    const scoring = questionnaire.scoring || {};
    const groups = scoring.groupQuestionIds || {};
    const groupLabels = scoring.groupLabels || {};
    const grades = scoring.grades || {};
    const groupScore = (ids) => (ids || []).reduce((sum, id) => sum + Number(responses[id]?.score || 0), 0);

    const harmScore = groupScore(groups.harm);
    const mentalScore = groupScore(groups.mental);
    const functionScore = groupScore(groups.function);
    const supportScore = groupScore(groups.support);
    const score = harmScore + mentalScore + functionScore + supportScore;
    const maxScore = 23;
    const primaryRiskScore = Number(responses[scoring.primaryRiskQuestionId]?.score || 0);
    const psychosisRiskScore = Number(responses[scoring.psychosisRiskQuestionId]?.score || 0);
    const psychosisMentalSum = psychosisRiskScore + mentalScore;

    let gradeKey = "E";
    if (primaryRiskScore === 1) {
      if (harmScore >= 2) {
        gradeKey = psychosisMentalSum >= 1 ? "A" : "B";
      } else {
        gradeKey = psychosisMentalSum >= 1 ? "C" : "D";
      }
    } else if (harmScore >= 1) {
      gradeKey = psychosisMentalSum >= 1 ? "C" : "D";
    } else {
      gradeKey = psychosisMentalSum >= 1 ? "D" : "E";
    }

    const grade = grades[gradeKey] || null;
    const flags = grade?.flagLevel && grade?.flagText ? [{ level: grade.flagLevel, text: grade.flagText }] : [];

    return {
      score,
      maxScore,
      normalizedScore: Math.round((score / maxScore) * 100),
      title: `${questionnaire.shortTitle || questionnaire.title} 결과`,
      scoreText: `${score}점 / ${maxScore}점`,
      bandText: grade?.label || "참고 구간 없음",
      highlights: [
        `총점 ${score}점 / ${maxScore}점`,
        `${groupLabels.harm || "자타해 위험"} ${harmScore}점`,
        `${groupLabels.mental || "정신상태"} ${mentalScore}점`,
        `${groupLabels.function || "기능수준"} ${functionScore}점`,
        `${groupLabels.support || "지지체계"} ${supportScore}점`,
        ...(grade?.description ? [grade.description] : [])
      ],
      flags,
      notes: scoring.notes || []
    };
  }

  function evaluateEndorsement(questionnaire, payload) {
    const score = (questionnaire.questions || []).reduce((count, question) => {
      return count + (Number(payload.responses[question.id]?.score || 0) > 0 ? 1 : 0);
    }, 0);
    const maxScore = (questionnaire.questions || []).length;
    const distressScore = (questionnaire.questions || []).reduce((sum, question) => {
      return sum + (question.subQuestions || []).reduce((subSum, subQuestion) => subSum + Number(payload.responses[subQuestion.id]?.score || 0), 0);
    }, 0);

    return {
      score,
      maxScore,
      normalizedScore: maxScore ? Math.round((score / maxScore) * 100) : 0,
      title: `${questionnaire.shortTitle || questionnaire.title} 결과`,
      scoreText: `예 ${score}개 / 힘듦 ${distressScore}점`,
      bandText: questionnaire.scoring?.summaryLabel || "참고용 요약",
      highlights: [
        `예 응답 수 ${score}개`,
        `힘듦 보조합계 ${distressScore}점`,
        ...(questionnaire.scoring?.referenceCutoff ? [`참고 절단값 ${questionnaire.scoring.referenceCutoff}`] : [])
      ],
      flags: questionnaire.scoring?.referenceWarning
        ? [{ level: "info", text: questionnaire.scoring.referenceWarning }]
        : [],
      notes: questionnaire.scoring?.notes || []
    };
  }

  function evaluateKmDQ(questionnaire, payload) {
    const symptomIds = questionnaire.scoring?.symptomQuestionIds || [];
    const score = symptomIds.reduce((count, id) => count + (Number(payload.responses[id]?.score || 0) > 0 ? 1 : 0), 0);
    const maxScore = symptomIds.length;
    const samePeriodScore = Number(payload.responses[questionnaire.scoring?.samePeriodQuestionId]?.score || 0);
    const impairmentScore = Number(payload.responses[questionnaire.scoring?.impairmentQuestionId]?.score || 0);
    const positive =
      score >= (questionnaire.scoring?.positiveRule?.symptomYesAtLeast || 7) &&
      samePeriodScore >= (questionnaire.scoring?.positiveRule?.samePeriodAtLeast || 1) &&
      impairmentScore >= (questionnaire.scoring?.positiveRule?.impairmentAtLeast || 1);

    return {
      score,
      maxScore,
      normalizedScore: maxScore ? Math.round((score / maxScore) * 100) : 0,
      title: `${questionnaire.shortTitle || questionnaire.title} 결과`,
      scoreText: `예 ${score}개 / ${maxScore}개`,
      bandText: positive ? "양성 참고 기준 충족" : "양성 참고 기준 미충족",
      highlights: [
        `증상 문항 예 응답 ${score}개`,
        `같은 시기 발생 여부 ${samePeriodScore > 0 ? "예" : "아니오"}`,
        `기능 문제 수준 ${impairmentLabel(impairmentScore)}`
      ],
      flags: positive
        ? [{ level: "warn", text: "K-MDQ 참고 양성 기준에 해당해 추가 평가를 권합니다." }]
        : [],
      notes: questionnaire.scoring?.notes || []
    };
  }

  function renderResultRecord(record) {
    ui.resultEmpty.classList.add("hidden");
    ui.resultPanel.classList.remove("hidden");
    renderPrintHeader(record);
    ui.resultTitle.textContent = record.evaluation.title;
    ui.resultScore.textContent = record.evaluation.scoreText;
    ui.resultBand.textContent = record.evaluation.bandText;

    ui.resultHighlights.innerHTML = "";
    (record.evaluation.highlights || []).forEach((text) => {
      const item = document.createElement("li");
      item.textContent = text;
      ui.resultHighlights.appendChild(item);
    });

    ui.resultFlags.innerHTML = "";
    (record.evaluation.flags || []).forEach((flag) => {
      const div = document.createElement("div");
      div.className = `flag-item ${flag.level || "info"}`;
      div.textContent = flag.text;
      ui.resultFlags.appendChild(div);
    });

    renderResultMeta(record);
    renderResultBreakdown(record);
    renderResultNotes(record);
    setActiveView("result");
  }

  function renderPrintHeader(record) {
    ui.printOrgLine.textContent = `${ORG_NAME} · ${TEAM_NAME}`;
    ui.printScaleLine.textContent = `${record.questionnaireTitle} 결과지`;
    ui.printMetaLine.textContent = [
      record.meta?.sessionDate ? `검사일 ${record.meta.sessionDate}` : "",
      record.meta?.workerName ? `담당자 ${record.meta.workerName}` : "",
      record.meta?.clientLabel ? `대상자 ${record.meta.clientLabel}` : "",
      record.meta?.birthDate ? `생년월일 ${record.meta.birthDate}` : "",
      record.meta?.sessionNote ? `비고 ${record.meta.sessionNote}` : ""
    ]
      .filter(Boolean)
      .join(" / ");
  }

  function renderResultMeta(record) {
    const items = [
      { label: "척도", value: record.questionnaireTitle },
      { label: "검사일", value: record.meta?.sessionDate || "" },
      { label: "응답 진행률", value: formatProgressSummary(record.progress) },
      { label: "저장시각", value: formatDateTime(record.createdAt) },
      { label: "담당자", value: record.meta?.workerName || "" },
      { label: "대상자", value: record.meta?.clientLabel || "" },
      { label: "생년월일", value: record.meta?.birthDate || "" },
      ...record.respondentDisplay,
      { label: "비고", value: record.meta?.sessionNote || "" }
    ].filter((item) => item.value);

    ui.resultMeta.innerHTML = "";
    items.forEach((item) => {
      const dt = document.createElement("dt");
      dt.textContent = item.label;
      const dd = document.createElement("dd");
      dd.textContent = item.value;
      ui.resultMeta.append(dt, dd);
    });
  }

  function renderResultBreakdown(record) {
    ui.resultBreakdown.innerHTML = "";
    (record.breakdown || []).forEach((item) => {
      const article = document.createElement("article");
      article.className = "breakdown-item";
      article.innerHTML = `
        <div class="breakdown-head">
          <p class="breakdown-title">${escapeHtml(item.number || "")} ${escapeHtml(item.text)}</p>
          <p class="breakdown-score">${item.score === null ? "-" : `${item.score}점`}</p>
        </div>
        <p class="breakdown-answer">${escapeHtml(item.answerLabel)}</p>
      `;

      if (item.subAnswers?.length) {
        const subWrap = document.createElement("div");
        subWrap.className = "breakdown-sublist";
        item.subAnswers.forEach((subItem) => {
          const row = document.createElement("div");
          row.className = "breakdown-subitem";
          row.innerHTML = `
            <p><strong>${escapeHtml(subItem.number || "")}</strong> ${escapeHtml(subItem.text)}</p>
            <p>${escapeHtml(subItem.answerLabel)} · ${subItem.score === null ? "-" : `${subItem.score}점`}</p>
          `;
          subWrap.appendChild(row);
        });
        article.appendChild(subWrap);
      }

      ui.resultBreakdown.appendChild(article);
    });
  }

  function renderResultNotes(record) {
    ui.resultNotes.innerHTML = "";
    const notes = record.evaluation.notes?.length ? record.evaluation.notes : ["이 결과는 현장 참고용이며 전문적 진단을 대체하지 않습니다."];
    notes.forEach((text) => {
      const item = document.createElement("li");
      item.textContent = text;
      ui.resultNotes.appendChild(item);
    });
  }

  function clearResult() {
    ui.resultEmpty.classList.remove("hidden");
    ui.resultPanel.classList.add("hidden");
    ui.resultFlags.innerHTML = "";
    ui.resultHighlights.innerHTML = "";
    ui.resultMeta.innerHTML = "";
    ui.resultBreakdown.innerHTML = "";
    ui.resultNotes.innerHTML = "";
  }

  async function onSaveResult() {
    if (!state.lastResult) {
      alert("먼저 결과를 계산해주세요.");
      return;
    }

    saveRecordToDevice(state.lastResult);
    let remoteMessage = "";

    if (state.syncSettings.syncEnabled && state.syncSettings.webAppUrl) {
      const synced = await syncSingleRecordToGoogleSheets(state.lastResult, "current_result", false);
      remoteMessage = synced ? " 구글 시트 DB에도 전송했습니다." : " 구글 시트 DB 전송은 실패했습니다.";
    }

    setActiveView("records");
    setHeroStatus(`${state.lastResult.meta?.clientLabel || "대상자"} 결과를 저장했습니다.${remoteMessage}`);
  }

  function saveRecordToDevice(record) {
    if (!record) {
      return;
    }

    const existingIndex = state.records.findIndex((entry) => entry.id === record.id);
    if (existingIndex >= 0) {
      state.records.splice(existingIndex, 1);
    }
    state.records.unshift(structuredCloneSafe(record));
    persistRecords();
  }

  function onExportCurrentRecord() {
    if (!state.lastResult) {
      alert("먼저 결과를 계산해주세요.");
      return;
    }
    exportRecordAsJson(state.lastResult);
  }

  function renderRecordsTable() {
    ui.recordsTableBody.innerHTML = "";
    const records = getFilteredRecords();
    const fragment = document.createDocumentFragment();

    ui.recordsStatusText.textContent = records.length ? `총 ${records.length}건을 표시합니다.` : "저장된 결과가 없습니다.";
    ui.recordsEmpty.classList.toggle("hidden", records.length > 0);

    records.forEach((record) => {
      const row = document.createElement("tr");
      row.innerHTML = `
        <td>${escapeHtml(record.meta?.sessionDate || "")}</td>
        <td class="cell-strong">${escapeHtml(record.meta?.clientLabel || "")}</td>
        <td>${escapeHtml(record.shortTitle || record.questionnaireTitle || "")}</td>
        <td>${escapeHtml(record.evaluation?.scoreText || "")}</td>
        <td>${escapeHtml(record.evaluation?.bandText || "")}</td>
        <td>${escapeHtml(record.meta?.workerName || "")}</td>
        <td>${escapeHtml(formatDateTime(record.createdAt))}</td>
        <td>
          <div class="table-actions">
            <button class="table-action" type="button" data-action="view" data-id="${escapeHtml(record.id)}">보기</button>
            <button class="table-action" type="button" data-action="load" data-id="${escapeHtml(record.id)}">불러오기</button>
            <button class="table-action" type="button" data-action="compare" data-id="${escapeHtml(record.id)}">비교</button>
            <button class="table-action" type="button" data-action="sync" data-id="${escapeHtml(record.id)}">시트전송</button>
            <button class="table-action" type="button" data-action="export" data-id="${escapeHtml(record.id)}">파일</button>
            <button class="table-action danger" type="button" data-action="delete" data-id="${escapeHtml(record.id)}">삭제</button>
          </div>
        </td>
      `;
      fragment.appendChild(row);
    });
    ui.recordsTableBody.appendChild(fragment);
  }

  function getFilteredRecords() {
    const search = ui.recordsSearchInput.value.trim().toLowerCase();
    const scaleId = ui.recordsScaleFilter.value;
    return [...state.records].filter((record) => {
      if (scaleId && record.questionnaireId !== scaleId) {
        return false;
      }
      if (!search) {
        return true;
      }
      const haystack = [
        record.meta?.clientLabel,
        record.meta?.workerName,
        record.questionnaireTitle,
        record.shortTitle,
        record.evaluation?.bandText
      ]
        .filter(Boolean)
        .join(" ")
        .toLowerCase();
      return haystack.includes(search);
    });
  }

  async function onRecordsTableClick(event) {
    const button = event.target.closest("[data-action][data-id]");
    if (!button) {
      return;
    }

    const record = state.records.find((entry) => entry.id === button.dataset.id);
    if (!record) {
      return;
    }

    switch (button.dataset.action) {
      case "view":
        state.lastResult = structuredCloneSafe(record);
        renderResultRecord(record);
        break;
      case "load":
        loadRecordIntoForm(record);
        break;
      case "compare":
        ui.dashboardNameInput.value = record.meta?.clientLabel || "";
        ui.dashboardBirthInput.value = record.meta?.birthDate || "";
        ui.dashboardScaleFilter.value = record.questionnaireId || "";
        renderDashboard();
        setActiveView("dashboard");
        await refreshDashboardRemoteRecords(true);
        break;
      case "export":
        exportRecordAsJson(record);
        break;
      case "sync":
        await syncSingleRecordToGoogleSheets(record, "saved_record", true);
        break;
      case "delete":
        if (!confirm("이 저장 결과를 삭제할까요?")) {
          return;
        }
        state.records = state.records.filter((entry) => entry.id !== record.id);
        persistRecords();
        break;
      default:
        break;
    }
  }

  function loadRecordIntoForm(record) {
    const questionnaire = state.questionnaires.get(record.questionnaireId);
    if (!questionnaire) {
      alert("해당 척도를 찾을 수 없습니다.");
      return;
    }

    state.currentQuestionnaireId = questionnaire.id;
    state.currentQuestionnaire = questionnaire;
    renderQuestionnaireNav(ui.scaleSearchInput.value);
    renderQuestionnaire(questionnaire);
    writeMetaFields(record.meta || {});

    (questionnaire.respondentFields || []).forEach((field) => {
      const value = record.respondent?.[field.id];
      if (!value) {
        return;
      }
      const input = document.querySelector(`input[name="respondent_${cssEscape(field.id)}"][value="${cssEscape(value)}"]`);
      if (input) {
        input.checked = true;
      }
    });

    (questionnaire.questions || []).forEach((question) => {
      const answer = record.answers?.[question.id];
      if (answer) {
        const input = document.querySelector(`input[name="${cssEscape(question.id)}"][value="${cssEscape(String(answer.value))}"]`);
        if (input) {
          input.checked = true;
        }
      }

      (question.subQuestions || []).forEach((subQuestion) => {
        const subAnswer = record.answers?.[subQuestion.id];
        if (!subAnswer) {
          return;
        }
        const input = document.querySelector(`input[name="${cssEscape(subQuestion.id)}"][value="${cssEscape(String(subAnswer.value))}"]`);
        if (input) {
          input.checked = true;
        }
      });
    });

    refreshQuestionUi();
    state.lastResult = structuredCloneSafe(record);
    setActiveView("screening");
    setHeroStatus("저장된 결과를 입력 화면으로 불러왔습니다.");
  }

  function exportAllRecordsAsJson() {
    if (!state.records.length) {
      alert("내보낼 저장 결과가 없습니다.");
      return;
    }
    downloadFile(
      `mindmap-scale-records-${new Date().toISOString().slice(0, 10)}.json`,
      JSON.stringify(state.records, null, 2),
      "application/json;charset=utf-8"
    );
  }

  function exportAllRecordsAsCsv() {
    if (!state.records.length) {
      alert("내보낼 저장 결과가 없습니다.");
      return;
    }
    downloadFile(
      `mindmap-scale-records-${new Date().toISOString().slice(0, 10)}.csv`,
      buildRecordsCsv(state.records),
      "text/csv;charset=utf-8"
    );
  }

  async function onImportJson(event) {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      const incoming = Array.isArray(parsed) ? parsed : Array.isArray(parsed.records) ? parsed.records : [];
      if (!incoming.length) {
        alert("가져올 기록이 없습니다.");
        return;
      }

      const map = new Map(state.records.map((record) => [record.id, record]));
      incoming.forEach((record) => {
        if (record?.id && record?.questionnaireId) {
          map.set(record.id, record);
        }
      });
      state.records = [...map.values()].sort((a, b) => String(b.createdAt || "").localeCompare(String(a.createdAt || "")));
      persistRecords();
      setHeroStatus(`${incoming.length}건의 기록을 가져왔습니다.`);
    } catch (error) {
      alert(`기록 불러오기에 실패했습니다: ${error.message}`);
    } finally {
      event.target.value = "";
    }
  }

  function onClearAllRecords() {
    if (!state.records.length) {
      alert("삭제할 저장 결과가 없습니다.");
      return;
    }
    if (!confirm("현재 브라우저에 저장된 결과를 모두 삭제할까요?")) {
      return;
    }
    state.records = [];
    state.lastResult = null;
    persistRecords();
    clearResult();
    setHeroStatus("현재 브라우저 저장 결과를 모두 삭제했습니다.");
  }

  async function onSyncCurrentResult() {
    if (!state.lastResult) {
      alert("먼저 결과를 계산해주세요.");
      return;
    }
    await syncSingleRecordToGoogleSheets(state.lastResult, "current_result", true);
  }

  async function onSyncAllRecords() {
    if (!state.records.length) {
      alert("전송할 저장 결과가 없습니다.");
      return;
    }

    await sendRecordsToGoogleSheets(state.records, "history_all", true);
  }

  async function onSyncQuestionnaires() {
    const settings = { ...state.syncSettings };
    const validation = validateGoogleSyncSettings(settings);
    if (!validation.ok) {
      setSyncStatus(validation.message, "error");
      alert(validation.message);
      validation.focusTarget?.focus();
      return;
    }

    setSyncBusy(true, "척도 마스터를 시트에 전송하고 있습니다...");
    try {
      await postGoogleSyncPayload(buildQuestionnaireMasterPayload(settings.syncToken));
      setSyncStatus(`척도 마스터 ${state.manifest.length}개 전송 요청을 보냈습니다.`, "success");
    } catch (error) {
      console.error("척도 마스터 전송 실패", error);
      setSyncStatus(`척도 마스터 전송 실패: ${error.message}`, "error");
      alert("척도 마스터 전송에 실패했습니다.");
    } finally {
      setSyncBusy(false);
    }
  }

  async function onCheckSyncStatus() {
    const settings = { ...state.syncSettings };
    const validation = validateGoogleSyncSettings(settings);
    if (!validation.ok) {
      setSyncStatus(validation.message, "error");
      validation.focusTarget?.focus();
      return;
    }

    setSyncBusy(true, "구글 시트 연동 상태를 확인하고 있습니다...");
    try {
      const response = await requestGoogleSyncJsonp(settings.webAppUrl, {
        action: "status",
        token: settings.syncToken
      });
      if (!response?.ok) {
        throw new Error(response?.error || "연동 상태를 확인하지 못했습니다.");
      }

      const statusMessage = `${response.spreadsheetName || "시트"} · 기록 ${response.recordRowCount || 0}건 · 문항 ${response.answerRowCount || 0}건`;
      setSyncStatus(`연동 확인 완료: ${statusMessage}`, "success");
    } catch (error) {
      console.error("구글 시트 연동 상태 확인 실패", error);
      setSyncStatus(`연동 상태 확인 실패: ${error.message}`, "error");
    } finally {
      setSyncBusy(false);
    }
  }

  async function syncSingleRecordToGoogleSheets(record, syncScope, showAlert) {
    const ok = await sendRecordsToGoogleSheets([record], syncScope, showAlert);
    return ok;
  }

  async function sendRecordsToGoogleSheets(records, syncScope, showAlert) {
    const settings = { ...state.syncSettings };
    const validation = validateGoogleSyncSettings(settings);
    if (!validation.ok) {
      setSyncStatus(validation.message, "error");
      if (showAlert) {
        alert(validation.message);
      }
      validation.focusTarget?.focus();
      return false;
    }

    const scopeLabel = syncScope === "current_result" ? "현재 결과" : syncScope === "history_all" ? "저장 결과 전체" : "선택 결과";
    setSyncBusy(true, `${scopeLabel}를 구글 시트 DB로 전송하고 있습니다...`);

    try {
      const payload = buildGoogleSyncPayload(records, syncScope, settings.syncToken);
      await postGoogleSyncPayload(payload);
      setSyncStatus(`${scopeLabel} ${records.length}건 전송 요청을 보냈습니다.`, "success");
      if (showAlert) {
        alert(`${scopeLabel} ${records.length}건을 구글 시트로 전송했습니다.`);
      }
      return true;
    } catch (error) {
      console.error("구글 시트 전송 실패", error);
      setSyncStatus(`구글 시트 전송 실패: ${error.message}`, "error");
      if (showAlert) {
        alert("구글 시트 전송에 실패했습니다.");
      }
      return false;
    } finally {
      setSyncBusy(false);
    }
  }

  function buildGoogleSyncPayload(records, syncScope, syncToken) {
    return {
      version: 1,
      source: "github-pages-scale-webapp",
      syncScope,
      sentAt: new Date().toISOString(),
      token: syncToken,
      appSettings: {
        organizationName: ORG_NAME,
        teamName: TEAM_NAME,
        contactNote: ""
      },
      records: records.map((record) => structuredCloneSafe(record))
    };
  }

  function buildQuestionnaireMasterPayload(syncToken) {
    return {
      version: 1,
      source: "github-pages-scale-webapp",
      syncScope: "questionnaire_master",
      sentAt: new Date().toISOString(),
      token: syncToken,
      appSettings: {
        organizationName: ORG_NAME,
        teamName: TEAM_NAME,
        contactNote: ""
      },
      questionnaires: [...state.questionnaires.values()].map((questionnaire) => structuredCloneSafe(questionnaire))
    };
  }

  async function postGoogleSyncPayload(payload) {
    const webAppUrl = state.syncSettings.webAppUrl;
    await fetch(webAppUrl, {
      method: "POST",
      mode: "no-cors",
      credentials: "include",
      cache: "no-store",
      headers: {
        "Content-Type": "text/plain;charset=utf-8"
      },
      body: JSON.stringify(payload)
    });
  }

  function onDashboardCriteriaChanged() {
    state.remoteDashboardRecords = [];
    setDashboardStatus("현재 브라우저 저장 결과를 먼저 표시합니다. 구글 시트 DB 조회 버튼을 누르면 누적 DB 기록까지 함께 불러옵니다.");
    renderDashboard();
  }

  async function onDashboardRefresh() {
    await refreshDashboardRemoteRecords(false);
  }

  async function refreshDashboardRemoteRecords(silent) {
    const settings = { ...state.syncSettings };
    const validation = validateGoogleSyncSettings(settings);
    if (!validation.ok) {
      setDashboardStatus(validation.message, "error");
      if (!silent) {
        alert(validation.message);
      }
      return;
    }

    const name = ui.dashboardNameInput.value.trim();
    const birthDate = ui.dashboardBirthInput.value;
    const questionnaireId = ui.dashboardScaleFilter.value;
    if (!name && !birthDate) {
      const message = "대상자 이름 또는 생년월일을 입력한 뒤 구글 시트 DB 조회를 실행해주세요.";
      setDashboardStatus(message, "error");
      if (!silent) {
        alert(message);
      }
      return;
    }

    setDashboardStatus("구글 시트 DB에서 대상자 기록을 조회하고 있습니다...");
    setSyncBusy(true);

    try {
      const response = await requestGoogleSyncJsonp(settings.webAppUrl, {
        action: "searchRecords",
        name,
        birthDate,
        questionnaireId,
        limit: "300",
        token: settings.syncToken
      });

      if (!response?.ok) {
        throw new Error(response?.error || "대상자 조회에 실패했습니다.");
      }

      state.remoteDashboardRecords = Array.isArray(response.records) ? response.records.map((record) => structuredCloneSafe(record)) : [];
      setDashboardStatus(`구글 시트 DB에서 ${state.remoteDashboardRecords.length}건을 불러왔습니다.`, "success");
      renderDashboard();
    } catch (error) {
      console.error("대시보드 DB 조회 실패", error);
      state.remoteDashboardRecords = [];
      setDashboardStatus(`구글 시트 DB 조회 실패: ${error.message}`, "error");
      if (!silent) {
        alert("구글 시트 DB 조회에 실패했습니다.");
      }
      renderDashboard();
    } finally {
      setSyncBusy(false);
    }
  }

  async function requestGoogleSyncJsonp(webAppUrl, params) {
    const callbackName = `mindmapScaleSync_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
    const searchParams = new URLSearchParams({
      ...params,
      callback: callbackName,
      _ts: String(Date.now())
    });
    const requestUrl = `${webAppUrl}${webAppUrl.includes("?") ? "&" : "?"}${searchParams.toString()}`;

    return new Promise((resolve, reject) => {
      const script = document.createElement("script");
      const cleanup = () => {
        delete window[callbackName];
        if (script.parentNode) {
          script.parentNode.removeChild(script);
        }
      };
      const timeoutId = window.setTimeout(() => {
        cleanup();
        reject(new Error("구글 시트 응답이 지연되고 있습니다. 권한 승인 페이지를 먼저 열었는지 확인해주세요."));
      }, 12000);

      window[callbackName] = (payload) => {
        window.clearTimeout(timeoutId);
        cleanup();
        resolve(payload);
      };

      script.onerror = () => {
        window.clearTimeout(timeoutId);
        cleanup();
        reject(new Error("구글 시트 웹앱을 불러오지 못했습니다. 승인 페이지를 먼저 열고, 공유 권한이 있는 계정인지 확인해주세요."));
      };
      script.src = requestUrl;
      document.body.appendChild(script);
    });
  }

  function renderDashboard() {
    populateSubjectSuggestions();
    const localRecords = getLocalDashboardRecords();
    const remoteRecords = getRemoteDashboardRecords();
    const records = mergeDashboardRecords(localRecords, remoteRecords);
    ui.dashboardSummary.innerHTML = "";
    ui.dashboardTableBody.innerHTML = "";

    if (!records.length) {
      ui.dashboardEmpty.classList.remove("hidden");
      destroyChart();
      ui.dashboardSummary.innerHTML = `
        <div class="summary-card">
          <span class="section-label">안내</span>
          <strong>대상자 검색 필요</strong>
          <p class="muted">대상자 이름과 생년월일을 입력하면 현재 브라우저 저장 결과를 먼저 확인할 수 있고, 구글 시트 DB 조회 버튼으로 누적 기록까지 함께 불러올 수 있습니다.</p>
        </div>
      `;
      return;
    }

    ui.dashboardEmpty.classList.add("hidden");
    renderDashboardSummary(records, localRecords.length, remoteRecords.length);
    renderDashboardChart(records);
    renderDashboardTable(records);
  }

  function getLocalDashboardRecords() {
    const name = normalizeLookupText(ui.dashboardNameInput.value);
    const birthDate = ui.dashboardBirthInput.value;
    const scaleId = ui.dashboardScaleFilter.value;

    if (!name && !birthDate) {
      return [];
    }

    return [...state.records]
      .filter((record) => {
        if (scaleId && record.questionnaireId !== scaleId) {
          return false;
        }
        const matchesName = name ? normalizeLookupText(record.meta?.clientLabel) === name : true;
        const matchesBirth = birthDate ? String(record.meta?.birthDate || "") === birthDate : true;
        return matchesName && matchesBirth;
      })
      .sort((a, b) => compareRecordDate(a).localeCompare(compareRecordDate(b)));
  }

  function getRemoteDashboardRecords() {
    const scaleId = ui.dashboardScaleFilter.value;
    return [...state.remoteDashboardRecords]
      .filter((record) => {
        if (scaleId && record.questionnaireId !== scaleId) {
          return false;
        }
        return true;
      })
      .sort((a, b) => compareRecordDate(a).localeCompare(compareRecordDate(b)));
  }

  function mergeDashboardRecords(localRecords, remoteRecords) {
    const merged = new Map();
    [...remoteRecords, ...localRecords].forEach((record) => {
      if (!record?.id) {
        return;
      }
      if (!merged.has(record.id)) {
        merged.set(record.id, record);
      }
    });
    return [...merged.values()].sort((a, b) => compareRecordDate(a).localeCompare(compareRecordDate(b)));
  }

  function renderDashboardSummary(records, localCount, remoteCount) {
    const uniqueScales = new Set(records.map((record) => record.questionnaireId)).size;
    const uniqueDates = new Set(records.map((record) => record.meta?.sessionDate || record.createdAt?.slice(0, 10))).size;
    const latest = records.at(-1);
    const riskCount = records.filter((record) => (record.evaluation?.flags || []).some((flag) => flag.level === "warn")).length;

    ui.dashboardSummary.innerHTML = `
      <div class="summary-card">
        <span class="section-label">검사 건수</span>
        <strong>${records.length}건</strong>
        <p class="muted">선택된 대상자의 누적 검사 수</p>
      </div>
      <div class="summary-card">
        <span class="section-label">검사일 수</span>
        <strong>${uniqueDates}일</strong>
        <p class="muted">검사가 기록된 날짜 수</p>
      </div>
      <div class="summary-card">
        <span class="section-label">사용 척도</span>
        <strong>${uniqueScales}종</strong>
        <p class="muted">대상자에게 적용된 척도 종류</p>
      </div>
      <div class="summary-card">
        <span class="section-label">위험 알림</span>
        <strong>${riskCount}건</strong>
        <p class="muted">주의 플래그가 포함된 검사 수</p>
      </div>
      <div class="summary-card">
        <span class="section-label">최근 검사</span>
        <strong>${escapeHtml(latest?.meta?.sessionDate || latest?.createdAt?.slice(0, 10) || "-")}</strong>
        <p class="muted">${escapeHtml(latest?.questionnaireTitle || "-")}</p>
      </div>
      <div class="summary-card">
        <span class="section-label">데이터 출처</span>
        <strong>기기 ${localCount}건 / 시트 ${remoteCount}건</strong>
        <p class="muted">브라우저 저장과 구글 시트 DB에서 합쳐진 기록 수</p>
      </div>
    `;
  }

  function renderDashboardChart(records) {
    if (!window.Chart) {
      return;
    }

    const labels = [...new Set(records.map((record) => record.meta?.sessionDate || record.createdAt?.slice(0, 10) || ""))];
    const byScale = new Map();

    records.forEach((record) => {
      const dateKey = record.meta?.sessionDate || record.createdAt?.slice(0, 10) || "";
      if (!byScale.has(record.questionnaireId)) {
        byScale.set(record.questionnaireId, {
          title: record.shortTitle || record.questionnaireTitle,
          values: new Map()
        });
      }
      byScale.get(record.questionnaireId).values.set(dateKey, Number(record.evaluation?.normalizedScore || 0));
    });

    const datasets = [...byScale.entries()].map(([, info], index) => ({
      label: info.title,
      data: labels.map((label) => info.values.get(label) ?? null),
      borderColor: SCALE_COLORS[index % SCALE_COLORS.length],
      backgroundColor: SCALE_COLORS[index % SCALE_COLORS.length],
      borderWidth: 2,
      pointRadius: 3,
      spanGaps: true,
      tension: 0.25
    }));

    const chartConfig = {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "bottom"
        },
        tooltip: {
          callbacks: {
            label(context) {
              return `${context.dataset.label}: ${context.parsed.y}%`;
            }
          }
        }
      },
      scales: {
        y: {
          min: 0,
          max: 100,
          ticks: {
            callback(value) {
              return `${value}%`;
            }
          },
          title: {
            display: true,
            text: "정규화 점수"
          }
        },
        x: {
          title: {
            display: true,
            text: "검사일"
          }
        }
      }
    };

    if (state.chart) {
      state.chart.data.labels = labels;
      state.chart.data.datasets = datasets;
      state.chart.options = chartConfig;
      state.chart.update();
      return;
    }

    state.chart = new Chart(ui.dashboardChart, {
      type: "line",
      data: { labels, datasets },
      options: chartConfig
    });
  }

  function destroyChart() {
    if (state.chart) {
      state.chart.destroy();
      state.chart = null;
    }
  }

  function renderDashboardTable(records) {
    ui.dashboardTableBody.innerHTML = "";
    const fragment = document.createDocumentFragment();
    records.forEach((record) => {
      const row = document.createElement("tr");
      row.innerHTML = `
        <td>${escapeHtml(record.meta?.sessionDate || "")}</td>
        <td>${escapeHtml(record.shortTitle || record.questionnaireTitle || "")}</td>
        <td>${escapeHtml(record.evaluation?.scoreText || "")}</td>
        <td>${escapeHtml(String(record.evaluation?.normalizedScore ?? 0))}%</td>
        <td>${escapeHtml(record.evaluation?.bandText || "")}</td>
        <td>${escapeHtml(record.meta?.workerName || "")}</td>
      `;
      fragment.appendChild(row);
    });
    ui.dashboardTableBody.appendChild(fragment);
  }

  function populateSubjectSuggestions() {
    const subjectMap = new Map();
    [...state.records, ...state.remoteDashboardRecords].forEach((record) => {
      const name = record.meta?.clientLabel?.trim();
      if (!name) {
        return;
      }
      const birthDate = record.meta?.birthDate || "";
      const key = `${name}__${birthDate}`;
      if (!subjectMap.has(key)) {
        subjectMap.set(key, { name, birthDate });
      }
    });

    const renderDatalist = (element) => {
      element.innerHTML = "";
      [...subjectMap.values()]
        .sort((a, b) => a.name.localeCompare(b.name, "ko"))
        .forEach((subject) => {
          const option = document.createElement("option");
          option.value = subject.name;
          option.label = subject.birthDate ? `${subject.name} (${subject.birthDate})` : subject.name;
          element.appendChild(option);
        });
    };

    renderDatalist(ui.subjectSuggestions);
    renderDatalist(ui.dashboardSubjectSuggestions);
  }

  function setActiveView(view) {
    state.activeView = view;
    ui.tabs.forEach((tab) => tab.classList.toggle("is-active", tab.dataset.view === view));
    ui.panes.forEach((pane) => pane.classList.toggle("hidden", pane.dataset.pane !== view));
  }

  function updateSummary() {
    const subjectKeys = new Set(
      state.records
        .map((record) => {
          const name = record.meta?.clientLabel?.trim();
          if (!name) {
            return "";
          }
          return `${name}__${record.meta?.birthDate || ""}`;
        })
        .filter(Boolean)
    );

    ui.summaryRecordCount.textContent = String(state.records.length);
    ui.summarySubjectCount.textContent = String(subjectKeys.size);
    ui.summaryScaleCount.textContent = String(state.manifest.length);
  }

  function setHeroStatus(message) {
    ui.heroStatusText.textContent = message;
  }

  function debounce(callback, delay = 150) {
    let timeoutId = 0;
    return (...args) => {
      window.clearTimeout(timeoutId);
      timeoutId = window.setTimeout(() => callback(...args), delay);
    };
  }

  function readMetaFields() {
    return {
      sessionDate: ui.sessionDate.value,
      workerName: ui.workerName.value,
      clientLabel: ui.clientLabel.value,
      birthDate: ui.birthDate.value,
      sessionNote: ui.sessionNote.value
    };
  }

  function writeMetaFields(meta = {}) {
    ui.sessionDate.value = meta.sessionDate || "";
    ui.workerName.value = meta.workerName || localStorage.getItem(STORAGE_KEYS.worker) || "";
    ui.clientLabel.value = meta.clientLabel || "";
    ui.birthDate.value = meta.birthDate || "";
    ui.sessionNote.value = meta.sessionNote || "";
  }

  function onResetCurrentForm() {
    const worker = ui.workerName.value;
    ui.screeningForm.reset();
    ui.workerName.value = worker;
    ui.sessionDate.value = new Date().toISOString().slice(0, 10);
    clearResult();
    refreshQuestionUi();
    setHeroStatus("현재 검사를 초기화했습니다.");
  }

  function findBand(bands, score) {
    return (bands || []).find((band) => score >= band.min && score <= band.max) || null;
  }

  function impairmentLabel(score) {
    if (score === 0) return "문제 없음";
    if (score === 1) return "경미";
    if (score === 2) return "중등도";
    if (score === 3) return "심각";
    return "미응답";
  }

  function maxOptionScore(options) {
    return (options || []).reduce((max, option) => Math.max(max, Number(option.score || 0)), 0);
  }

  function structuredCloneSafe(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function createLocalId() {
    if (window.crypto?.randomUUID) {
      return window.crypto.randomUUID();
    }
    return `record_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
  }

  function exportRecordAsJson(record) {
    const filename = `${sanitizeFilePart(record.meta?.clientLabel || "대상자")}_${sanitizeFilePart(record.shortTitle || record.questionnaireId || "척도")}_${record.meta?.sessionDate || new Date().toISOString().slice(0, 10)}.json`;
    downloadFile(filename, JSON.stringify(record, null, 2), "application/json;charset=utf-8");
  }

  function buildRecordsCsv(records) {
    const header = [
      "record_id",
      "검사일",
      "저장시각",
      "대상자",
      "생년월일",
      "담당자",
      "척도ID",
      "척도명",
      "원점수",
      "최대점수",
      "정규화점수",
      "결과구간",
      "응답진행률",
      "비고"
    ];

    const rows = records.map((record) => [
      record.id,
      record.meta?.sessionDate || "",
      formatDateTime(record.createdAt),
      record.meta?.clientLabel || "",
      record.meta?.birthDate || "",
      record.meta?.workerName || "",
      record.questionnaireId || "",
      record.questionnaireTitle || "",
      record.evaluation?.score ?? "",
      record.evaluation?.maxScore ?? "",
      record.evaluation?.normalizedScore ?? "",
      record.evaluation?.bandText || "",
      formatProgressSummary(record.progress),
      record.meta?.sessionNote || ""
    ]);

    return [header, ...rows].map((row) => row.map(csvCell).join(",")).join("\r\n");
  }

  function csvCell(value) {
    const text = String(value ?? "");
    if (/[",\r\n]/.test(text)) {
      return `"${text.replace(/"/g, '""')}"`;
    }
    return text;
  }

  function downloadFile(filename, content, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);
  }

  function formatDateTime(iso) {
    if (!iso) {
      return "";
    }
    const date = new Date(iso);
    if (Number.isNaN(date.getTime())) {
      return iso;
    }
    return `${date.toLocaleDateString("ko-KR")} ${date.toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit" })}`;
  }

  function sanitizeFilePart(value) {
    return String(value || "")
      .replace(/[\\/:*?"<>|]+/g, "-")
      .replace(/\s+/g, "_")
      .slice(0, 40);
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function cssEscape(value) {
    if (window.CSS?.escape) {
      return window.CSS.escape(String(value));
    }
    return String(value).replace(/[^a-zA-Z0-9_-]/g, "\\$&");
  }
})();
