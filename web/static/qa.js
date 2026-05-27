const statusBadge = document.querySelector("#statusBadge");
const runtimeMode = document.querySelector("#runtimeMode");
const loadingOverlay = document.querySelector("#loadingOverlay");
const loadingTitle = document.querySelector("#loadingTitle");
const loadingMessage = document.querySelector("#loadingMessage");
const loadingSteps = document.querySelector("#loadingSteps");
const loadingCancelButton = document.querySelector("#loadingCancelButton");

const resultTitle = document.querySelector("#resultTitle");
const resultMeta = document.querySelector("#resultMeta");
const errorPanel = document.querySelector("#errorPanel");
const downloadPanel = document.querySelector("#downloadPanel");
const emptyState = document.querySelector("#emptyState");

const tcForm = document.querySelector("#tcForm");
const tsForm = document.querySelector("#tsForm");
const folderQaForm = document.querySelector("#folderQaForm");
const qaDumpRoot = document.querySelector("#qaDumpRoot");
const tcSourceRoot = document.querySelector("#tcSourceRoot");
const unitResultRoot = document.querySelector("#unitResultRoot");
const tsSourceRoot = document.querySelector("#tsSourceRoot");
const integrationResultRoot = document.querySelector("#integrationResultRoot");
const uiDesignRoot = document.querySelector("#uiDesignRoot");
const qaSourcePreview = document.querySelector("#qaSourcePreview");
const taskTabs = [...document.querySelectorAll("[data-task-tab]")];
const taskPanels = [...document.querySelectorAll("[data-task-panel]")];
const fileInputs = [...document.querySelectorAll("input[type='file']")];
const LAST_DUMP_ROOT_KEY = "atoz:lastDumpRoot";
const QA_SOURCE_EXTENSIONS = new Set([".hwp", ".hwpx", ".pdf", ".xlsx"]);
const QA_LOADING_STEPS = [
  "입력 파일을 확인하는 중",
  "설계서와 산출물을 분석하는 중",
  "QA 결과를 생성하는 중",
];
const initialFileLabels = new Map(
  fileInputs.map((input) => {
    const label = document.querySelector(`[data-file-label="${input.id}"]`);
    return [input.id, label?.textContent || "파일 선택"];
  })
);

let currentRuntimeMode = null;
let loadingStepTimer = null;
let loadingStepIndex = 0;
let activeRequest = null;

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

function renderLoadingSteps(steps = [], activeIndex = 0) {
  if (!loadingSteps) return;
  if (!steps.length) {
    loadingSteps.hidden = true;
    loadingSteps.replaceChildren();
    return;
  }

  loadingSteps.hidden = false;
  loadingSteps.replaceChildren(
    ...steps.map((step, index) => {
      const item = document.createElement("li");
      item.textContent = step;
      if (index < activeIndex) item.classList.add("done");
      if (index === activeIndex) item.classList.add("active");
      return item;
    })
  );
}

function stopLoadingSteps() {
  if (loadingStepTimer) {
    clearInterval(loadingStepTimer);
    loadingStepTimer = null;
  }
}

function startLoadingSteps(steps = []) {
  stopLoadingSteps();
  loadingStepIndex = 0;
  renderLoadingSteps(steps, loadingStepIndex);

  if (steps.length <= 1) return;
  loadingStepTimer = setInterval(() => {
    loadingStepIndex = Math.min(loadingStepIndex + 1, steps.length - 1);
    renderLoadingSteps(steps, loadingStepIndex);
  }, 2600);
}

function setLoading(isLoading, title = "처리 중", message = "요청을 처리하고 있습니다.", steps = []) {
  if (!isLoading) {
    stopLoadingSteps();
    renderLoadingSteps([]);
  }
  loadingOverlay.hidden = !isLoading;
  loadingTitle.textContent = title;
  loadingMessage.textContent = message;
  loadingMessage.classList.toggle("is-hidden", isLoading && steps.length > 0);
  if (isLoading) startLoadingSteps(steps.length ? steps : [message]);
  document.querySelectorAll("button").forEach((button) => {
    button.disabled = isLoading;
    button.classList.toggle("is-loading", isLoading);
  });
  if (loadingCancelButton) {
    loadingCancelButton.disabled = !isLoading;
    loadingCancelButton.classList.toggle("is-loading", false);
  }
}

function beginCancelableRequest() {
  if (activeRequest) activeRequest.controller.abort();
  activeRequest = {
    controller: new AbortController(),
    requestId: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`,
  };
  return activeRequest;
}

function endCancelableRequest(request) {
  if (activeRequest === request) activeRequest = null;
}

function isAbortError(error) {
  return error?.name === "AbortError";
}

function markRequestCanceled() {
  setBadge("취소됨", "");
  resultMeta.textContent = "작업을 취소했습니다. 파일을 다시 선택한 뒤 실행할 수 있습니다.";
}

function sendCancelRequest(request) {
  if (!request?.requestId) return;
  fetch("/api/cancel-request", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ request_id: request.requestId }),
    keepalive: true,
  }).catch(() => {});
}

function setActiveTask(taskName) {
  taskTabs.forEach((tab) => tab.classList.toggle("active", tab.dataset.taskTab === taskName));
  taskPanels.forEach((panel) => panel.classList.toggle("active", panel.dataset.taskPanel === taskName));
  clearDownloads();
  clearError();
  resultTitle.textContent = taskName === "tc" ? "단위시험 케이스 생성" : "통합시험 시나리오 생성";
  resultMeta.textContent = "파일을 업로드한 뒤 생성 버튼을 누르세요.";
}

function fileDisplayPath(file) {
  return file.webkitRelativePath || file.name;
}

function fileExtension(file) {
  const name = file.name || "";
  const dotIndex = name.lastIndexOf(".");
  return dotIndex >= 0 ? name.slice(dotIndex).toLowerCase() : "";
}

function fileSizeLabel(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 KB";
  const kb = bytes / 1024;
  if (kb < 1024) return `${Math.max(1, Math.round(kb))} KB`;
  return `${(kb / 1024).toFixed(1)} MB`;
}

function renderFolderPreview(input, preview, allowedExtensions, emptyText) {
  if (!preview) return;

  const files = [...(input?.files || [])];
  if (!files.length) {
    preview.hidden = true;
    preview.replaceChildren();
    return;
  }

  const rootName = fileDisplayPath(files[0]).includes("/")
    ? fileDisplayPath(files[0]).split("/")[0]
    : "";
  const targetFiles = files
    .filter((file) => allowedExtensions.has(fileExtension(file)))
    .sort((left, right) => fileDisplayPath(left).localeCompare(fileDisplayPath(right), "ko"));
  const skippedCount = files.length - targetFiles.length;

  preview.hidden = false;
  preview.innerHTML = `
    <div class="selected-folder-preview-head">
      <strong>${escapeHtml(rootName || "선택된 폴더")}</strong>
      <span>${targetFiles.length}개 대상${skippedCount ? ` · ${skippedCount}개 제외` : ""}</span>
    </div>
    ${targetFiles.length ? `
      <div class="selected-folder-file-list">
        ${targetFiles.map((file) => `
          <article class="selected-folder-file">
            <span>${escapeHtml(fileDisplayPath(file))}</span>
            <small>${escapeHtml(fileSizeLabel(file.size))}</small>
          </article>
        `).join("")}
      </div>
    ` : `
      <p class="selected-folder-empty">${escapeHtml(emptyText)}</p>
    `}
  `;
}

function renderQaSourcePreview(input) {
  renderFolderPreview(input, qaSourcePreview, QA_SOURCE_EXTENSIONS, "HWP, HWPX, PDF, XLSX 파일이 없습니다.");
}

function updateFileLabel(input) {
  const label = document.querySelector(`[data-file-label="${input.id}"]`);
  if (!label) return;

  if (!input.files.length) {
    label.textContent = initialFileLabels.get(input.id) || "파일 선택";
    if (input.id === "qaSourceFiles") renderQaSourcePreview(input);
    return;
  }

  const firstPath = input.files[0].webkitRelativePath || input.files[0].name;
  const folderName = input.hasAttribute("webkitdirectory") && firstPath.includes("/")
    ? firstPath.split("/")[0]
    : "";

  if (folderName) {
    label.textContent = `${folderName} · ${input.files.length}개`;
    if (input.id === "qaSourceFiles") renderQaSourcePreview(input);
    return;
  }

  label.textContent = input.files.length > 1
    ? `${input.files[0].name} 외 ${input.files.length - 1}개`
    : input.files[0].name;
  if (input.id === "qaSourceFiles") renderQaSourcePreview(input);
}

function buildFormData(form) {
  const body = new FormData();
  for (const element of form.elements) {
    if (!element.name) continue;
    if (element.type === "file") {
      for (const file of element.files) {
        body.append(element.name, file, file.webkitRelativePath || file.name);
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
}

function clearError() {
  if (!errorPanel) return;
  errorPanel.hidden = true;
  errorPanel.replaceChildren();
}

function renderDownloads(files, title, options = {}) {
  const downloadFiles = (files || []).filter((file) => file && file.download_url);

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
  const fallback = taskName === "folder"
    ? "산출물 QA 생성 중 문제가 발생했습니다. QA 대상 산출물 폴더 안의 대상 문서를 확인하세요."
    : taskName === "ts"
    ? "통합시험 시나리오 생성 중 문제가 발생했습니다. 업로드한 파일을 다시 확인하세요."
    : "단위시험 케이스 생성 중 문제가 발생했습니다. 업로드한 파일을 다시 확인하세요.";

  if (!rawMessage) {
    return {
      summary: fallback,
      checks: ["업로드한 파일 형식과 위치가 올바른지 확인하세요."],
      detail: "",
    };
  }

  if (taskName === "folder") {
    if (rawMessage.includes("찾지 못했습니다") || rawMessage.includes("요구사항 ID 기준")) {
      return {
        summary: rawMessage.split("\n")[0],
        checks: [
          "check.html에서 문서 반영이 끝난 산출물 폴더 경로인지 확인하세요.",
          "다섯 문서가 함께 있다면 '추가 문서 폴더'를 선택하세요.",
          "화면/사용자인터페이스설계서는 HWP, HWPX, PDF 중 하나를 직접 업로드하세요.",
          "여러 건이면 '화면설계서 폴더'에 전체 폴더 경로를 입력하면 하위 파일을 일괄 탐색합니다.",
          "업로드한 설계서와 QA 대상 산출물 폴더의 단위시험케이스/단위시험결과서/통합시험시나리오/통합시험결과서 파일명이 같은 SFR 요구사항 ID를 포함하는지 확인하세요.",
          "QA 대상 산출물 폴더 안에 단위시험케이스 양식이 없으면 TC HWPX를 기존 위치에 배치할 수 없습니다.",
          "예: SFR-ESS-001 설계서는 SFR-ESS-001 단위시험케이스, SFR-ESS-001 단위시험결과서, SFR-ESS-001 통합시험시나리오, SFR-ESS-001 통합시험결과서와 매칭됩니다.",
        ],
      detail: [
        qaDumpRoot?.value ? `현재 QA 대상 산출물 폴더: ${qaDumpRoot.value}` : "",
        uiDesignRoot?.value ? `현재 화면설계서 폴더: ${uiDesignRoot.value}` : "",
        tcSourceRoot?.value ? `현재 단위시험 폴더: ${tcSourceRoot.value}` : "",
        unitResultRoot?.value ? `현재 단위시험결과서 폴더: ${unitResultRoot.value}` : "",
        tsSourceRoot?.value ? `현재 통합시험 폴더: ${tsSourceRoot.value}` : "",
        integrationResultRoot?.value ? `현재 통합시험결과서 폴더: ${integrationResultRoot.value}` : "",
      ].filter(Boolean).join("\n"),
      };
    }
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

    if (rawMessage.includes("사용자인터페이스 설계서 문서를 선택하세요") || rawMessage.includes("사용자인터페이스 설계서 PDF를 선택하세요") || rawMessage.includes("ui_pdf") || rawMessage.includes("PDF 파일을 찾을 수 없습니다")) {
      return {
        summary: "'사용자인터페이스설계서' 파일을 확인하세요.",
        checks: [
          "'사용자인터페이스설계서' 칸에는 HWP, HWPX, PDF 중 하나를 넣어야 합니다.",
          "문서 안에 화면 ID, 화면명, 처리흐름 정보가 포함되어 있어야 단위시험 케이스를 만들 수 있습니다.",
          "파일이 비어 있거나 텍스트를 추출할 수 없는 문서라면 다른 파일로 다시 선택하세요.",
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
          "문서에 화면 ID와 처리흐름이 포함되어 있는지 확인하세요.",
          "스캔 이미지 위주의 PDF나 이미지형 문서라면 텍스트 추출이 되지 않아 생성 결과가 없을 수 있습니다.",
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
        "'단위시험 케이스' 칸에는 단위시험 케이스에서 생성한 XLSX를 넣어야 합니다.",
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
  const title = taskName === "folder"
    ? "산출물 QA 생성 실패"
    : taskName === "ts" ? "통합시험 시나리오 생성 실패" : "단위시험 케이스 생성 실패";
  const info = friendlyErrorInfo(error.message, taskName);
  resultTitle.textContent = title;
  resultMeta.textContent = "아래 안내를 확인한 뒤 파일을 다시 선택하세요.";
  renderErrorPanel(info);
  setBadge("확인 필요", "error");
}

async function postForm(endpoint, body, request) {
  if (request?.requestId && body instanceof FormData) {
    body.append("request_id", request.requestId);
  }
  const response = await fetch(endpoint, { method: "POST", body, signal: request?.controller.signal });
  const data = await response.json().catch(() => ({ error: "서버 응답을 읽지 못했습니다." }));
  if (!response.ok) {
    const error = new Error(data.error || "처리 실패");
    error.data = data;
    throw error;
  }
  return data;
}

async function postJson(endpoint, payload) {
  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });
  const data = await response.json().catch(() => ({ error: "서버 응답을 읽지 못했습니다." }));
  if (!response.ok) {
    throw new Error(data.error || "처리 실패");
  }
  return data;
}

function initializeDumpRoot() {
  if (!qaDumpRoot) return;

  const params = new URLSearchParams(window.location.search);
  const queryDumpRoot = params.get("dump_root") || "";
  let storedDumpRoot = "";
  try {
    storedDumpRoot = localStorage.getItem(LAST_DUMP_ROOT_KEY) || "";
  } catch {
    storedDumpRoot = "";
  }

  qaDumpRoot.value = queryDumpRoot || storedDumpRoot;
}

function renderFolderQaResult(data) {
  const placedFiles = Array.isArray(data.placed_files) ? data.placed_files : [];
  const sourceFiles = Array.isArray(data.source_files) ? data.source_files : [];
  const requirementItems = Array.isArray(data.requirement_items) ? data.requirement_items : [];
  const missingRequirements = Array.isArray(data.missing_requirements) ? data.missing_requirements : [];
  const roleCounts = data.role_counts || {};
  emptyState.hidden = Boolean(placedFiles.length || sourceFiles.length || requirementItems.length || missingRequirements.length);
  clearError();

  const roleLabels = {
    ui_design: "사용자인터페이스설계서",
    tc_template: "단위시험케이스 양식",
    unit_result_template: "단위시험결과서 양식",
    ts_template: "통합시험시나리오 양식",
    integration_result_template: "통합시험결과서 양식",
  };
  const roleOrder = ["ui_design", "tc_template", "unit_result_template", "ts_template", "integration_result_template"];
  const resultLabels = {
    tc_hwpx: "단위시험케이스 생성본",
    unit_result_hwpx: "단위시험결과서 생성본",
    ts_xlsx: "통합시험시나리오 생성본",
    integration_result_xlsx: "통합시험결과서 생성본",
  };
  const sourceByRequirement = new Map();
  const placedByRequirement = new Map();
  const requirementById = new Map();

  for (const item of requirementItems) {
    const requirementId = item.requirement_id || "-";
    requirementById.set(requirementId, item);
  }
  for (const file of sourceFiles) {
    const requirementId = file.requirement_id || "-";
    if (!sourceByRequirement.has(requirementId)) sourceByRequirement.set(requirementId, []);
    sourceByRequirement.get(requirementId).push(file);
  }
  for (const file of placedFiles) {
    const requirementId = file.requirement_id || "-";
    if (!placedByRequirement.has(requirementId)) placedByRequirement.set(requirementId, []);
    placedByRequirement.get(requirementId).push(file);
  }

  const requirementIds = [...new Set([
    ...requirementItems.map((item) => item.requirement_id || "-"),
    ...sourceFiles.map((file) => file.requirement_id || "-"),
    ...placedFiles.map((file) => file.requirement_id || "-"),
  ])].sort((left, right) => String(left).localeCompare(String(right), "ko"));

  function basename(path) {
    const value = String(path || "");
    return value.split(/[\\/]/).filter(Boolean).pop() || value || "-";
  }

  function dirname(path) {
    const value = String(path || "");
    const parts = value.split(/[\\/]/).filter(Boolean);
    if (parts.length <= 1) return "";
    return parts.slice(0, -1).join("\\");
  }

  function renderSourceRows(files) {
    const sortedFiles = [...files].sort((left, right) => {
      const leftRole = roleOrder.indexOf(left.role);
      const rightRole = roleOrder.indexOf(right.role);
      return (leftRole < 0 ? 99 : leftRole) - (rightRole < 0 ? 99 : rightRole);
    });

    return sortedFiles.map((file) => `
      <article class="folder-qa-file input">
        <span>${escapeHtml(roleLabels[file.role] || file.label || file.role || "입력 파일")}</span>
        <strong>${escapeHtml(basename(file.path))}</strong>
        <code>${escapeHtml(file.path || "-")}</code>
      </article>
    `).join("") || `<p class="folder-qa-empty">매칭된 입력 파일이 없습니다.</p>`;
  }

  function renderPlacedRows(files) {
    return files.map((file) => `
      <article class="folder-qa-file output">
        <span>${escapeHtml(resultLabels[file.kind] || file.label || file.kind || "배치 파일")}</span>
        <strong>${escapeHtml(basename(file.path))}</strong>
        <code>${escapeHtml(file.path || "-")}</code>
        ${file.backup_path ? `<small>기존 파일 백업: ${escapeHtml(basename(file.backup_path))}</small>` : `<small>기존 파일 백업 없음</small>`}
      </article>
    `).join("") || `<p class="folder-qa-empty">생성되어 배치된 파일이 없습니다.</p>`;
  }

  function renderPathDetails(sources, outputs) {
    const sourceRows = sources.map((file) => renderPathItem(
      roleLabels[file.role] || file.label || file.role || "입력 파일",
      file.path,
    )).join("");
    const outputRows = outputs.map((file) => renderPathItem(
      resultLabels[file.kind] || file.label || file.kind || "배치 파일",
      file.path,
    )).join("");
    const backupRows = outputs
      .filter((file) => file.backup_path)
      .map((file) => renderPathItem("기존 파일 백업", file.backup_path))
      .join("");

    if (!sourceRows && !outputRows && !backupRows) return "";

    function renderPathItem(label, path) {
      const folder = dirname(path);
      return `
        <article class="folder-qa-path-item">
          <span>${escapeHtml(label)}</span>
          <strong>${escapeHtml(basename(path))}</strong>
          ${folder ? `<code title="${escapeHtml(path || "")}">${escapeHtml(folder)}</code>` : ""}
        </article>
      `;
    }

    function renderPathGroup(title, rows) {
      if (!rows) return "";
      return `
        <section class="folder-qa-path-group">
          <h4>${escapeHtml(title)}</h4>
          <div>${rows}</div>
        </section>
      `;
    }

    return `
      <details class="folder-qa-path-details">
        <summary>
          <span>상세 경로 보기</span>
          <small>입력 ${sources.length}개 · 배치 ${outputs.length}개${backupRows ? " · 백업 있음" : ""}</small>
        </summary>
        <div class="folder-qa-path-groups">
          ${renderPathGroup("선택된 입력 파일", sourceRows)}
          ${renderPathGroup("생성 후 배치된 파일", outputRows)}
          ${renderPathGroup("백업된 기존 파일", backupRows)}
        </div>
      </details>
    `;
  }

  function renderRequirementCards() {
    if (!requirementIds.length) {
      return `
        <article class="folder-qa-requirement has-errors">
          <div class="folder-qa-requirement-head">
            <strong>매칭된 요구사항 없음</strong>
            <span>확인 필요</span>
          </div>
          <p class="folder-qa-empty">아래 누락 항목을 확인하세요.</p>
        </article>
      `;
    }

    return requirementIds.map((requirementId) => {
      const item = requirementById.get(requirementId) || {};
      const isError = item.status === "error";
      const sources = sourceByRequirement.get(requirementId) || [];
      const outputs = placedByRequirement.get(requirementId) || [];
      const statusText = isError ? "실패" : outputs.length ? "배치 완료" : "대기";
      const metricText = isError
        ? item.error || "처리 중 오류가 발생했습니다."
        : `TC ${item.tc_count ?? 0}행 · TS ${item.ts_count ?? 0}행`;

      return `
        <article class="folder-qa-requirement ${isError ? "has-errors" : ""}">
          <div class="folder-qa-requirement-head">
            <div>
              <strong>${escapeHtml(requirementId)}</strong>
              <small>${escapeHtml(metricText)}</small>
            </div>
            <span>${escapeHtml(statusText)}</span>
          </div>
          <div class="folder-qa-columns">
            <section>
              <h3>선택된 입력 파일</h3>
              <div class="folder-qa-file-list">${renderSourceRows(sources)}</div>
            </section>
            <section>
              <h3>생성 후 배치된 파일</h3>
              <div class="folder-qa-file-list">${renderPlacedRows(outputs)}</div>
            </section>
          </div>
          ${renderPathDetails(sources, outputs)}
        </article>
      `;
    }).join("");
  }

  downloadPanel.hidden = false;
  downloadPanel.innerHTML = `
    <div class="download-panel-head">
      <strong>${data.ok === false ? "QA 생성 현황" : "QA 배치 결과"}</strong>
      <span>${data.processed_requirement_count ?? 0}/${data.requirement_count ?? 0}개 요구사항</span>
    </div>
    <div class="role-count-list">
      <article>
        <span>입력 설계서</span>
        <strong>${escapeHtml(roleCounts.ui_design ?? "-")}</strong>
      </article>
      <article>
        <span>TC 양식</span>
        <strong>${escapeHtml(roleCounts.tc_template ?? "-")}</strong>
      </article>
      <article>
        <span>단위시험 결과서</span>
        <strong>${escapeHtml(roleCounts.unit_result_template ?? "-")}</strong>
      </article>
      <article>
        <span>TS 양식</span>
        <strong>${escapeHtml(roleCounts.ts_template ?? "-")}</strong>
      </article>
      <article>
        <span>통합시험 결과서</span>
        <strong>${escapeHtml(roleCounts.integration_result_template ?? "-")}</strong>
      </article>
    </div>
    <div class="folder-qa-result-list">
      ${renderRequirementCards()}
    </div>
    ${missingRequirements.length ? `
      <div class="missing-list">
        <div class="folder-qa-section-head">
          <strong>누락된 매칭 항목</strong>
        </div>
        ${missingRequirements.map((item) => `
          <article>
            <span>${escapeHtml(item.requirement_id || "-")}</span>
            <strong>${escapeHtml((item.missing || []).join(", "))}</strong>
          </article>
        `).join("")}
      </div>
    ` : ""}
  `;
}

async function runFolderQa(event) {
  event.preventDefault();
  const dumpRoot = qaDumpRoot?.value?.trim() || "";
  if (!dumpRoot) {
    showGenerationError(new Error("QA 대상 산출물 폴더 경로를 입력하세요."), "folder");
    return;
  }

  try {
    localStorage.setItem(LAST_DUMP_ROOT_KEY, dumpRoot);
  } catch {
    // 저장 실패는 실행을 막지 않는다.
  }

  setBadge("처리중", "busy");
  setLoading(true, "QA 생성 중", "QA 산출물을 생성하고 있습니다.", QA_LOADING_STEPS);
  clearDownloads();
  clearError();
  resultTitle.textContent = "QA 생성";
  resultMeta.textContent = "요청 처리 중입니다.";

  const request = beginCancelableRequest();

  try {
    const data = await postForm("/api/run-qa-folder", buildFormData(folderQaForm), request);
    setBadge("완료", "done");
    resultMeta.textContent = `QA 대상 산출물 폴더: ${data.dump_root || dumpRoot}`;
    renderFolderQaResult(data);
  } catch (error) {
    if (isAbortError(error)) {
      markRequestCanceled();
      return;
    }
    showGenerationError(error, "folder");
    if (error.data?.source_files || error.data?.missing_requirements) {
      renderFolderQaResult(error.data);
    }
  } finally {
    endCancelableRequest(request);
    setLoading(false);
  }
}

async function runTcGeneration(event) {
  event.preventDefault();
  if (!validateFiles(tcForm, [
    ["#tcTemplateHwpx", "기존 단위시험 케이스 HWPX를 선택하세요."],
    ["#tcUiPdf", "사용자인터페이스 설계서 문서를 선택하세요."],
  ])) return;

  setBadge("처리중", "busy");
  setLoading(true, "단위시험 케이스 생성 중", "QA 산출물을 생성하고 있습니다.", QA_LOADING_STEPS);
  clearDownloads();
  clearError();
  resultTitle.textContent = "단위시험 케이스 생성";
  resultMeta.textContent = "요청 처리 중입니다.";

  const request = beginCancelableRequest();

  try {
    const data = await postForm("/api/generate-tc", buildFormData(tcForm), request);
    setBadge("완료", "done");
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
    if (isAbortError(error)) {
      markRequestCanceled();
      return;
    }
    showGenerationError(error, "tc");
  } finally {
    endCancelableRequest(request);
    setLoading(false);
  }
}

async function runTsGeneration(event) {
  event.preventDefault();
  if (!validateFiles(tsForm, [
    ["#tsTemplateXlsx", "기존 통합시험 시나리오 XLSX를 선택하세요."],
    ["#tsTcXlsx", "단위시험 케이스 XLSX를 선택하세요."],
    ["#tsUiPdf", "사용자인터페이스설계서 문서를 선택하세요."],
  ])) return;

  setBadge("처리중", "busy");
  setLoading(true, "통합시험 시나리오 생성 중", "QA 산출물을 생성하고 있습니다.", QA_LOADING_STEPS);
  clearDownloads();
  clearError();
  resultTitle.textContent = "통합시험 시나리오 생성";
  resultMeta.textContent = "요청 처리 중입니다.";

  const request = beginCancelableRequest();

  try {
    const data = await postForm("/api/generate-ts", buildFormData(tsForm), request);
    setBadge("완료", "done");
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
    if (isAbortError(error)) {
      markRequestCanceled();
      return;
    }
    showGenerationError(error, "ts");
  } finally {
    endCancelableRequest(request);
    setLoading(false);
  }
}

taskTabs.forEach((tab) => tab.addEventListener("click", () => setActiveTask(tab.dataset.taskTab)));
fileInputs.forEach((input) => input.addEventListener("change", () => updateFileLabel(input)));
folderQaForm?.addEventListener("submit", runFolderQa);
tcForm.addEventListener("submit", runTcGeneration);
tsForm.addEventListener("submit", runTsGeneration);
loadingCancelButton?.addEventListener("click", () => {
  sendCancelRequest(activeRequest);
  activeRequest?.controller.abort();
});

initializeDumpRoot();
setBadge("대기", "");
loadRuntimeMode();
