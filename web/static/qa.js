const statusBadge = document.querySelector("#statusBadge");
const runtimeMode = document.querySelector("#runtimeMode");
const loadingOverlay = document.querySelector("#loadingOverlay");
const loadingTitle = document.querySelector("#loadingTitle");
const loadingMessage = document.querySelector("#loadingMessage");

const resultTitle = document.querySelector("#resultTitle");
const resultMeta = document.querySelector("#resultMeta");
const errorPanel = document.querySelector("#errorPanel");
const downloadPanel = document.querySelector("#downloadPanel");
const emptyState = document.querySelector("#emptyState");

const tcCount = document.querySelector("#tcCount");
const tsCount = document.querySelector("#tsCount");
const processStatus = document.querySelector("#processStatus");
const generatedFileCount = document.querySelector("#generatedFileCount");

const tcForm = document.querySelector("#tcForm");
const tsForm = document.querySelector("#tsForm");
const taskTabs = [...document.querySelectorAll("[data-task-tab]")];
const taskPanels = [...document.querySelectorAll("[data-task-panel]")];
const fileInputs = [...document.querySelectorAll("input[type='file']")];
const ENABLE_TS_ANALYSIS_MOCK = true;
const initialFileLabels = new Map(
  fileInputs.map((input) => {
    const label = document.querySelector(`[data-file-label="${input.id}"]`);
    return [input.id, label?.textContent || "파일 선택"];
  })
);

let currentRuntimeMode = null;

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function setBadge(text, mode) {
  if (!statusBadge) return;
  statusBadge.textContent = text;
  statusBadge.className = `status-pill ${mode || ""}`.trim();
}

function setProcessStatus(text) {
  if (!processStatus) return;
  processStatus.textContent = text;
}

function setRuntimeMode(data) {
  currentRuntimeMode = data || currentRuntimeMode;
  if (!runtimeMode || !currentRuntimeMode) return;

  const mode = currentRuntimeMode.label || currentRuntimeMode.mode || "-";
  const modelText = currentRuntimeMode.model ? ` · ${currentRuntimeMode.model}` : "";
  runtimeMode.textContent = `${mode}${modelText}`;
}

async function loadRuntimeMode() {
  try {
    const response = await fetch("/api/runtime-mode", { cache: "no-store" });
    if (!response.ok) throw new Error("runtime mode request failed");
    setRuntimeMode(await response.json());
  } catch {
    setRuntimeMode({ mode: "unknown", label: "실행 모드 확인 실패" });
  }
}

function setLoading(isLoading, title = "처리 중", message = "요청을 처리하고 있습니다.") {
  loadingOverlay.hidden = !isLoading;
  loadingTitle.textContent = title;
  loadingMessage.textContent = message;
  document.querySelectorAll("button").forEach((button) => {
    button.disabled = isLoading;
    button.classList.toggle("is-loading", isLoading);
  });
}

function setActiveTask(taskName) {
  taskTabs.forEach((tab) => tab.classList.toggle("active", tab.dataset.taskTab === taskName));
  taskPanels.forEach((panel) => panel.classList.toggle("active", panel.dataset.taskPanel === taskName));
  clearDownloads();
  clearError();
  setProcessStatus("대기");
  resultTitle.textContent = taskName === "tc" ? "단위시험 케이스 생성" : "통합시험 시나리오 생성";
  resultMeta.textContent = "파일을 업로드한 뒤 생성 버튼을 누르세요.";
}

function updateFileLabel(input) {
  const label = document.querySelector(`[data-file-label="${input.id}"]`);
  if (!label) return;

  if (!input.files.length) {
    label.textContent = initialFileLabels.get(input.id) || "파일 선택";
    return;
  }

  if (input.files.length === 1) {
    label.textContent = input.files[0].name;
    return;
  }

  label.textContent = `${input.files.length}개 선택됨`;
}

function buildFormData(form) {
  const body = new FormData();
  for (const element of form.elements) {
    if (!element.name) continue;
    if (element.type === "file") {
      for (const file of element.files) {
        body.append(element.name, file, file.name);
      }
      continue;
    }
    body.append(element.name, element.value);
  }
  return body;
}

function clearDownloads() {
  downloadPanel.hidden = true;
  downloadPanel.replaceChildren();
  emptyState.hidden = false;
  if (generatedFileCount) generatedFileCount.textContent = "-";
}

function clearError() {
  if (!errorPanel) return;
  errorPanel.hidden = true;
  errorPanel.replaceChildren();
}

function renderDownloads(files, title, options = {}) {
  const downloadFiles = (files || []).filter((file) => file && file.download_url);
  if (generatedFileCount) generatedFileCount.textContent = downloadFiles.length || "-";

  if (!downloadFiles.length) {
    clearDownloads();
    return;
  }

  emptyState.hidden = true;
  clearError();
  downloadPanel.hidden = false;
  downloadPanel.className = `download-panel ${options.plainPanel ? "plain" : ""}`.trim();
  downloadPanel.innerHTML = `
    <div class="download-output-card">
      <div class="download-panel-head">
        <strong>${escapeHtml(title)}</strong>
        <span>${downloadFiles.length}개</span>
      </div>
      <div class="download-list">
        ${downloadFiles.map((file) => `
          <a class="download-button" href="${escapeHtml(file.download_url)}" download="${escapeHtml(file.download_name || file.name || "")}">
            <span>${escapeHtml((file.kind || "file").toUpperCase())}</span>
            <strong>${escapeHtml(file.download_name || file.name || "download")}</strong>
          </a>
        `).join("")}
      </div>
    </div>
  `;
}

function renderSourceResults(sourceResults, options = {}) {
  if (!sourceResults?.length || downloadPanel.hidden) return;

  const title = options.title || "사전 분석 요약";
  const itemLabel = options.itemLabel || "항목";
  const countLabel = options.countLabel || "생성 행";
  const fileLabel = options.fileLabel || "파일";
  const primaryTotalLabel = options.primaryTotalLabel || "화면";
  const showCountMetric = options.showCountMetric !== false;
  const showPrimaryTotal = options.showPrimaryTotal !== false;
  const countedResults = sourceResults.filter((item) => !item.is_summary);

  const okCount = countedResults.filter((item) => item.ok).length;
  const warningCount = countedResults.length - okCount;
  const totalScreens = countedResults.reduce((sum, item) => {
    const analysis = item.analysis || {};
    return sum + Number(analysis.screen_count || 0);
  }, 0);
  const totalRows = countedResults.reduce((sum, item) => sum + Number(item.count || 0), 0);
  const totalFiles = countedResults.reduce((sum, item) => sum + Number(item.file_count || 0), 0);

  const rows = sourceResults.map((item) => {
    const analysis = item.analysis || {};
    const risks = analysis.risks || [];
    const recommendations = analysis.recommendations || [];
    const screens = analysis.screens || [];
    const quality = analysis.quality || (item.ok ? "good" : "warning");
    const statusText = item.ok ? "완료" : "확인 필요";
    const statusClass = item.ok ? quality : "poor";
    const showStatus = !item.is_summary;
    const primaryMetricLabel = analysis.metric_label || "화면";
    const primaryMetricCount = analysis.metric_count ?? analysis.screen_count ?? screens.length ?? 0;
    const noteLabel = item.is_summary ? "매칭된 세트" : "권장 조치";
    const screenIds = screens
      .map((screen) => typeof screen === "string" ? screen : screen?.screen_id)
      .filter(Boolean);
    const visibleScreenIds = screenIds.slice(0, 12);
    const hiddenScreenCount = Math.max(0, screenIds.length - visibleScreenIds.length);
    const isMatchedSet = Boolean(item.source_tc || item.source_ui);
    const cardTitle = item.source_tc || item.source_ui
      ? "매칭 완료 세트"
      : (item.source_pdf || "-");
    const metricItems = showCountMetric
      ? [
        `<span>${escapeHtml(primaryMetricLabel)} <strong>${escapeHtml(primaryMetricCount)}</strong></span>`,
        `<span>${escapeHtml(countLabel)} <strong>${escapeHtml(item.count ?? 0)}</strong></span>`,
        `<span>${escapeHtml(fileLabel)} <strong>${escapeHtml(item.file_count ?? 0)}</strong></span>`,
      ].join("")
      : "";

    return `
      <article class="source-result-card ${escapeHtml(statusClass)} ${item.is_summary ? "summary" : ""} ${isMatchedSet ? "matched" : ""}">
        <div class="source-result-head">
          <div>
            <strong>${escapeHtml(cardTitle)}</strong>
            ${isMatchedSet ? "" : `<small>${escapeHtml(analysis.summary || "분석 요약이 없습니다.")}</small>`}
          </div>
          ${showStatus ? `<span class="${escapeHtml(statusClass)}">${statusText}</span>` : ""}
        </div>
        ${isMatchedSet ? `
          <div class="source-set-summary">
            <span><b>${escapeHtml(primaryMetricLabel)}</b><strong>${escapeHtml(primaryMetricCount)}</strong></span>
            <span><b>화면ID</b><strong>${escapeHtml(analysis.summary || `${primaryMetricCount}개 일치`)}</strong></span>
            <span><b>${escapeHtml(fileLabel)}</b><strong>${escapeHtml(item.file_count ?? 0)}</strong></span>
          </div>
        ` : ""}
        ${item.source_tc || item.source_ui ? `
          <div class="source-section-label">매칭 파일</div>
          <div class="source-match-files">
            <div><b>TC 파일</b><span>${escapeHtml(item.source_tc || "-")}</span></div>
            <div><b>UI 설계서</b><span>${escapeHtml(item.source_ui || "-")}</span></div>
          </div>
        ` : ""}
        ${metricItems ? `<div class="source-result-metrics">
          ${metricItems}
        </div>` : ""}
        ${visibleScreenIds.length ? `
          <div class="source-section-label">포함 화면ID</div>
          <div class="source-screen-list" aria-label="화면ID 목록">
            ${visibleScreenIds.map((screenId) => `<span>${escapeHtml(screenId)}</span>`).join("")}
            ${hiddenScreenCount ? `<span>+${hiddenScreenCount}</span>` : ""}
          </div>
        ` : ""}
        ${risks.length ? `
          <div class="source-result-notes">
            <b>확인할 점</b>
            <ul>
            ${risks.map((risk) => `<li>${escapeHtml(risk)}</li>`).join("")}
            </ul>
          </div>
        ` : ""}
        ${recommendations.length ? `
          <div class="source-result-notes">
            <b>${escapeHtml(noteLabel)}</b>
            <ul>
            ${recommendations.map((recommendation) => `<li>${escapeHtml(recommendation)}</li>`).join("")}
            </ul>
          </div>
        ` : ""}
      </article>
    `;
  }).join("");

  downloadPanel.insertAdjacentHTML("beforeend", `
    <div class="source-result-list">
      <div class="download-panel-head">
        <strong>${escapeHtml(title)}</strong>
      </div>
      <div class="source-result-overview">
        <span>처리 ${escapeHtml(itemLabel)} <strong>${countedResults.length}</strong></span>
        <span>확인 필요 <strong>${warningCount}</strong></span>
        ${showPrimaryTotal ? `<span>${escapeHtml(primaryTotalLabel)} <strong>${totalScreens}</strong></span>` : ""}
        <span>${escapeHtml(fileLabel)} <strong>${totalFiles}</strong></span>
      </div>
      ${rows}
    </div>
  `);
}

function validateFiles(form, rules) {
  for (const [selector, message] of rules) {
    const input = form.querySelector(selector);
    if (!input || !input.files.length) {
      showGenerationError(new Error(message), form === tsForm ? "ts" : "tc");
      return false;
    }
  }
  return true;
}

function selectedFileName(inputId) {
  const input = document.querySelector(`#${inputId}`);
  return input?.files?.[0]?.name || "";
}

function friendlyErrorInfo(message, taskName) {
  const rawMessage = String(message || "").trim();
  const fallback = taskName === "ts"
    ? "통합시험 시나리오 생성 중 문제가 발생했습니다. 업로드한 파일을 다시 확인하세요."
    : "단위시험 케이스 생성 중 문제가 발생했습니다. 업로드한 파일을 다시 확인하세요.";

  if (!rawMessage) {
    return {
      summary: fallback,
      checks: ["업로드한 파일 형식과 위치가 올바른지 확인하세요."],
      detail: "",
    };
  }

  const selectedTcMatch = rawMessage.match(/선택된 단위시험 케이스 파일:\s*(.+)$/m);

  if (taskName === "tc") {
    const selectedTemplate = selectedFileName("tcTemplateHwpx");
    const selectedPdf = selectedFileName("tcUiPdf");
    const detail = [
      selectedTemplate ? `현재 선택된 기존 단위시험 케이스 파일: ${selectedTemplate}` : "",
      selectedPdf ? `현재 선택된 사용자인터페이스설계서 파일: ${selectedPdf}` : "",
    ].filter(Boolean).join("\n");

    if (rawMessage.includes("기존 단위시험 케이스 HWPX를 선택하세요") || rawMessage.includes("template_hwpx")) {
      return {
        summary: "'기존 단위시험 케이스' 파일을 확인하세요.",
        checks: [
          "1단계의 '기존 단위시험 케이스' 칸에는 HWPX 파일을 넣어야 합니다.",
          "단위시험 ID, 단위시험 명, 사전조건, 화면 ID, 수행 결과가 있는 기존 단위시험 케이스 양식인지 확인하세요.",
          "PDF나 XLSX 파일을 넣으면 단위시험 케이스 HWPX를 만들 수 없습니다.",
        ],
        detail,
      };
    }

    if (rawMessage.includes("사용자인터페이스 설계서 PDF를 선택하세요") || rawMessage.includes("ui_pdf") || rawMessage.includes("PDF 파일을 찾을 수 없습니다")) {
      return {
        summary: "'사용자인터페이스설계서' 파일을 확인하세요.",
        checks: [
          "1단계의 '사용자인터페이스설계서' 칸에는 PDF 파일을 넣어야 합니다.",
          "PDF 안에 화면 ID, 화면명, 처리흐름 정보가 포함되어 있어야 단위시험 케이스를 만들 수 있습니다.",
          "파일이 비어 있거나 잘못된 PDF라면 다른 파일로 다시 선택하세요.",
        ],
        detail,
      };
    }

    if (rawMessage.includes("HWPX 양식") || rawMessage.includes("단위시험 ID") || rawMessage.includes("수행 결과") || rawMessage.includes("채울 단위시험 표")) {
      return {
        summary: "기존 단위시험 케이스 HWPX 양식을 확인하세요.",
        checks: [
          "업로드한 HWPX 안에 단위시험 케이스 표 양식이 있는지 확인하세요.",
          "양식에는 '단위시험 ID'와 테스트 스텝 표의 '수행 결과' 항목이 있어야 합니다.",
          "일반 HWPX 문서나 표 구조가 다른 파일이면 생성 중에 실패할 수 있습니다.",
        ],
        detail,
      };
    }

    if (rawMessage.includes("단위시험 케이스 생성 결과가 없습니다")) {
      return {
        summary: "사용자인터페이스설계서에서 생성할 테스트 케이스를 찾지 못했습니다.",
        checks: [
          "PDF에 화면 ID와 처리흐름이 포함되어 있는지 확인하세요.",
          "스캔 이미지 위주의 PDF라면 텍스트 추출이 되지 않아 생성 결과가 없을 수 있습니다.",
          "Ollama 서버와 모델이 정상 동작 중인지 확인한 뒤 다시 시도하세요.",
        ],
        detail,
      };
    }
  }

  if (rawMessage.includes("단위시험 케이스 엑셀에서 데이터를 찾지 못했습니다")) {
    return {
      summary: "'단위시험 케이스' 파일을 확인하세요.",
      checks: [
        "'단위시험 케이스' 칸에는 1단계에서 생성한 단위시험 케이스 XLSX를 넣어야 합니다.",
        "'기존 통합시험 시나리오' 칸에는 기존 통합시험 시나리오 양식 XLSX를 넣어야 합니다.",
        "두 XLSX 파일이 서로 바뀌었거나, 단위시험 케이스 칸에 통합시험 시나리오 파일이 들어가면 데이터를 찾을 수 없습니다.",
      ],
      detail: selectedTcMatch ? `현재 선택된 단위시험 케이스 파일: ${selectedTcMatch[1].trim()}` : "",
    };
  }

  if (taskName === "ts" && rawMessage.includes("list index out of range")) {
    return {
      summary: "'기존 통합시험 시나리오' 파일을 읽지 못했습니다.",
      checks: [
        "'기존 통합시험 시나리오' 칸에 정상적인 XLSX 파일을 넣었는지 확인하세요.",
        "엑셀에서 해당 파일을 열어 '다른 이름으로 저장'으로 새 XLSX 파일을 만든 뒤 다시 업로드해보세요.",
        "한셀, 구버전 엑셀, 외부 도구에서 만든 파일은 셀 스타일 정보 때문에 읽지 못할 수 있습니다.",
      ],
      detail: selectedFileName("tsTemplateXlsx")
        ? `현재 선택된 기존 통합시험 시나리오 파일: ${selectedFileName("tsTemplateXlsx")}`
        : "",
    };
  }

  const looksLikeInternalError =
    /^name '.+' is not defined$/i.test(rawMessage) ||
    rawMessage.includes("Traceback") ||
    rawMessage.includes("pywintypes.com_error");

  if (looksLikeInternalError) {
    return {
      summary: fallback,
      checks: ["업로드한 파일을 다시 확인한 뒤 한 번 더 시도하세요.", "같은 문제가 반복되면 서버 로그를 확인하세요."],
      detail: "",
    };
  }

  return {
    summary: rawMessage.split("\n")[0],
    checks: rawMessage.split("\n").slice(1).filter(Boolean),
    detail: "",
  };
}

function renderErrorPanel(info) {
  if (!errorPanel) return;

  const checks = info.checks || [];
  errorPanel.hidden = false;
  errorPanel.innerHTML = `
    <div class="error-panel-head">
      <strong>${escapeHtml(info.summary)}</strong>
      <span>확인 필요</span>
    </div>
    ${checks.length ? `
      <ul>
        ${checks.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}
      </ul>
    ` : ""}
    ${info.detail ? `<p>${escapeHtml(info.detail)}</p>` : ""}
  `;
}

function showGenerationError(error, taskName) {
  const title = taskName === "ts" ? "통합시험 시나리오 생성 실패" : "단위시험 케이스 생성 실패";
  const info = friendlyErrorInfo(error.message, taskName);
  resultTitle.textContent = title;
  resultMeta.textContent = "아래 안내를 확인한 뒤 파일을 다시 선택하세요.";
  renderErrorPanel(info);
  setBadge("확인 필요", "error");
  setProcessStatus("확인 필요");
}

async function postForm(endpoint, body) {
  const response = await fetch(endpoint, { method: "POST", body });
  const data = await response.json().catch(() => ({ error: "서버 응답을 읽지 못했습니다." }));
  if (!response.ok) {
    throw new Error(data.error || "처리 실패");
  }
  return data;
}

async function runTcGeneration(event) {
  event.preventDefault();
  if (!validateFiles(tcForm, [
    ["#tcTemplateHwpx", "기존 단위시험 케이스 HWPX를 선택하세요."],
    ["#tcUiPdf", "사용자인터페이스 설계서 PDF를 선택하세요."],
  ])) return;

  setBadge("처리중", "busy");
  setProcessStatus("생성 중");
  setLoading(true, "단위시험 케이스 생성 중", "AI가 화면별 테스트 케이스를 만들고 있습니다.");
  clearDownloads();
  clearError();
  resultTitle.textContent = "단위시험 케이스 생성";
  resultMeta.textContent = "요청 처리 중입니다.";

  try {
    const data = await postForm("/api/generate-tc", buildFormData(tcForm));
    setBadge("완료", "done");
    setProcessStatus("완료");
    if (tcCount) tcCount.textContent = data.count ?? 0;
    const sourceCount = data.source_count ?? data.source_results?.length ?? 1;
    const failedCount = data.failed_count ?? 0;
    resultMeta.textContent = `처리 설계서 ${sourceCount}개 · 생성 행 수 ${data.count ?? 0}개 · 다운로드 ${data.download_files?.length || 0}개${failedCount ? ` · 실패 ${failedCount}개` : ""}`;
    renderDownloads(data.download_files || data.files, "단위시험 케이스 파일");
    renderSourceResults(data.source_results || [], {
      title: "단위시험 케이스 사전 분석",
      itemLabel: "설계서",
      countLabel: "생성 행",
      fileLabel: "생성 파일",
    });
  } catch (error) {
    showGenerationError(error, "tc");
  } finally {
    setLoading(false);
  }
}

async function runTsGeneration(event) {
  event.preventDefault();
  if (!validateFiles(tsForm, [
    ["#tsTemplateXlsx", "기존 통합시험 시나리오 XLSX를 선택하세요."],
    ["#tsTcXlsx", "단위시험 케이스 XLSX를 선택하세요."],
    ["#tsUiPdf", "사용자인터페이스설계서 PDF를 선택하세요."],
  ])) return;

  setBadge("처리중", "busy");
  setProcessStatus("생성 중");
  setLoading(true, "통합시험 시나리오 생성 중", "단위시험 케이스를 통합시험 시나리오 양식으로 변환하고 있습니다.");
  clearDownloads();
  clearError();
  resultTitle.textContent = "통합시험 시나리오 생성";
  resultMeta.textContent = "요청 처리 중입니다.";

  try {
    const data = await postForm("/api/generate-ts", buildFormData(tsForm));
    setBadge("완료", "done");
    setProcessStatus("완료");
    if (tsCount) tsCount.textContent = data.count ?? 0;
    const setCount = data.set_count ?? data.source_results?.length ?? 1;
    const failedCount = data.failed_count ?? 0;
    resultMeta.textContent = `처리 세트 ${setCount}개 · 다운로드 ${data.download_files?.length || 0}개${failedCount ? ` · 확인 필요 ${failedCount}개` : ""}`;
    renderDownloads(data.download_files || data.files, "통합시험 시나리오 파일", { plainPanel: true });
    renderSourceResults(data.source_results || [], {
      title: "통합시험 시나리오 세트 분석",
      itemLabel: "세트",
      fileLabel: "생성 파일",
      primaryTotalLabel: "화면",
      showPrimaryTotal: false,
      showCountMetric: false,
    });
  } catch (error) {
    showGenerationError(error, "ts");
  } finally {
    setLoading(false);
  }
}

function renderTemporaryTsAnalysisMock() {
  if (!ENABLE_TS_ANALYSIS_MOCK) return;

  const mockFiles = [
    { name: "ts_tc_MFDS-AD3-A0301-01-사용자인터페이스설계서_V1.3_SFR-ECP-007.xlsx", download_url: "#" },
    { name: "ts_tc_MFDS-ADT-A0101-01-사용자인터페이스설계서_V1.0_SFR-SFD-001.xlsx", download_url: "#" },
    { name: "ts_tc_MFDS-ADM-B0201-01-사용자인터페이스설계서_V1.2_SFR-MGT-003.xlsx", download_url: "#" },
  ];
  const mockSourceResults = [
    {
      ok: true,
      count: 24,
      file_count: 1,
      source_tc: "tc_MFDS-AD3-A0301-01-사용자인터페이스설계서_V1.3_SFR-ECP-007.xlsx",
      source_ui: "MFDS-AD3-A0301-01-사용자인터페이스설계서_V1.3_SFR-ECP-007.pdf",
      analysis: {
        summary: "7개 일치",
        screen_count: 7,
        screens: [
          "UI-ECP-007-01-001",
          "UI-ECP-007-01-002",
          "UI-ECP-007-01-003",
          "UI-ECP-007-02-001",
          "UI-ECP-007-03-001",
          "UI-ECP-007-03-002",
          "UI-ECP-007-03-003",
        ],
      },
    },
    {
      ok: true,
      count: 10,
      file_count: 1,
      source_tc: "tc_MFDS-ADT-A0101-01-사용자인터페이스설계서_V1.0_SFR-SFD-001.xlsx",
      source_ui: "MFDS-ADT-A0101-01-사용자인터페이스설계서_V1.0_SFR-SFD-001.pdf",
      analysis: {
        summary: "2개 일치",
        screen_count: 2,
        screens: ["UI-SFD-001-01-01", "UI-SFD-001-01-02"],
      },
    },
    {
      ok: false,
      count: 0,
      file_count: 0,
      source_tc: "tc_MFDS-ADM-B0201-01-사용자인터페이스설계서_V1.2_SFR-MGT-003.xlsx",
      source_ui: "MFDS-ADM-B0201-01-사용자인터페이스설계서_V1.2_SFR-MGT-003.pdf",
      analysis: {
        summary: "4개 후보",
        screen_count: 4,
        quality: "warning",
        risks: ["SFR 키는 일치하지만 일부 화면ID가 단위시험 케이스에 없습니다."],
        recommendations: ["누락 화면ID가 실제 시나리오 대상인지 확인하세요."],
        screens: [
          "UI-MGT-003-01-001",
          "UI-MGT-003-01-002",
          "UI-MGT-003-02-001",
          "UI-MGT-003-02-002",
        ],
      },
    },
  ];

  setActiveTask("ts");
  setBadge("디자인 확인", "done");
  setProcessStatus("mock");
  resultTitle.textContent = "통합시험 시나리오 생성";
  resultMeta.textContent = "임시 mock 데이터 3세트로 디자인을 확인 중입니다.";
  renderDownloads(mockFiles, "통합시험 시나리오 파일", { plainPanel: true });
  renderSourceResults(mockSourceResults, {
    title: "통합시험 시나리오 세트 분석",
    itemLabel: "세트",
    fileLabel: "생성 파일",
    primaryTotalLabel: "화면",
    showPrimaryTotal: false,
    showCountMetric: false,
  });
}

taskTabs.forEach((tab) => tab.addEventListener("click", () => setActiveTask(tab.dataset.taskTab)));
fileInputs.forEach((input) => input.addEventListener("change", () => updateFileLabel(input)));
tcForm.addEventListener("submit", runTcGeneration);
tsForm.addEventListener("submit", runTsGeneration);

setBadge("대기", "");
setProcessStatus("대기");
loadRuntimeMode();
renderTemporaryTsAnalysisMock();
