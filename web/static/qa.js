const statusBadge = document.querySelector("#statusBadge");
const runtimeMode = document.querySelector("#runtimeMode");

const resultTitle = document.querySelector("#resultTitle");
const resultMeta = document.querySelector("#resultMeta");
const errorPanel = document.querySelector("#errorPanel");
const downloadPanel = document.querySelector("#downloadPanel");
const emptyState = document.querySelector("#emptyState");

const tcForm = document.querySelector("#tcForm");
const tsForm = document.querySelector("#tsForm");
const folderQaForm = document.querySelector("#folderQaForm");
const folderPreviewButton = document.querySelector("#folderPreviewButton");
const qaDumpRoot = document.querySelector("#qaDumpRoot");
const qaSourceRoot = document.querySelector("#qaSourceRoot");
const qaSourceFiles = document.querySelector("#qaSourceFiles");
const tcSourceRoot = document.querySelector("#tcSourceRoot");
const unitResultRoot = document.querySelector("#unitResultRoot");
const tsSourceRoot = document.querySelector("#tsSourceRoot");
const integrationResultRoot = document.querySelector("#integrationResultRoot");
const uiDesignRoot = document.querySelector("#uiDesignRoot");
const taskTabs = [...document.querySelectorAll("[data-task-tab]")];
const taskPanels = [...document.querySelectorAll("[data-task-panel]")];
const fileInputs = [...document.querySelectorAll("input[type='file']")];
const LAST_DUMP_ROOT_KEY = "atoz:lastDumpRoot";
const LAST_QA_FOLDER_JOB_KEY = "atoz:lastQaFolderJob";
const initialFileLabels = new Map(
  fileInputs.map((input) => {
    const label = document.querySelector(`[data-file-label="${input.id}"]`);
    return [input.id, label?.textContent || "파일 선택"];
  })
);

let currentRuntimeMode = null;
let activeRequest = null;
const cancelRequestedJobIds = new Set();

function rememberQaFolderJob(jobId) {
  if (!jobId) return;
  try {
    localStorage.setItem(LAST_QA_FOLDER_JOB_KEY, jobId);
  } catch {
    // localStorage를 쓸 수 없는 환경이면 복구 기능만 건너뛴다.
  }
}

function forgetQaFolderJob(jobId = "") {
  try {
    const stored = localStorage.getItem(LAST_QA_FOLDER_JOB_KEY) || "";
    if (!jobId || stored === jobId) {
      localStorage.removeItem(LAST_QA_FOLDER_JOB_KEY);
    }
  } catch {
    // 저장소 접근 실패는 화면 동작을 막지 않는다.
  }
}

function getRememberedQaFolderJob() {
  try {
    return localStorage.getItem(LAST_QA_FOLDER_JOB_KEY) || "";
  } catch {
    return "";
  }
}

function isTerminalQaFolderStatus(status) {
  return ["done", "error", "cancelled"].includes(status || "");
}

function isActiveQaFolderStatus(status) {
  return ["queued", "running"].includes(status || "");
}

function setQaFolderFinalStatus(data, fallbackRoot = "") {
  if (data?.status === "cancelled") {
    setBadge("취소됨", "");
    resultMeta.textContent = "QA 생성 작업을 취소했습니다. 필요한 경우 다시 생성 및 배치를 실행하세요.";
    return;
  }

  const isSuccess = data?.status === "done" && data?.ok;
  setBadge(isSuccess ? "완료" : "확인 필요", isSuccess ? "done" : "error");
  resultMeta.textContent = `QA 대상 산출물 폴더: ${data?.dump_root || fallbackRoot || "-"}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function formatUserErrorMessage(message) {
  const text = String(message || "").trim();
  if (!text) return "";
  return text.split("|")[0].trim() || text;
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

function setLoading(isLoading) {
  document.querySelectorAll("button").forEach((button) => {
    button.disabled = isLoading;
    button.classList.toggle("is-loading", isLoading);
  });
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

function updateFileLabel(input) {
  const label = document.querySelector(`[data-file-label="${input.id}"]`);
  if (!label) return;

  if (!input.files.length) {
    label.textContent = initialFileLabels.get(input.id) || "파일 선택";
    return;
  }

  const firstPath = input.files[0].webkitRelativePath || input.files[0].name;
  const folderName = input.hasAttribute("webkitdirectory") && firstPath.includes("/")
    ? firstPath.split("/")[0]
    : "";

  if (folderName) {
    label.textContent = `${folderName} · ${input.files.length}개`;
    return;
  }

  label.textContent = input.files.length > 1
    ? `${input.files[0].name} 외 ${input.files.length - 1}개`
    : input.files[0].name;
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

function renderIndividualGenerationResult(data, options = {}) {
  const files = (data.download_files || data.files || []).filter((file) => file && file.download_url);
  const sourceResults = Array.isArray(data.source_results) ? data.source_results : [];
  const title = options.title || "생성 결과";
  const itemLabel = options.itemLabel || "문서";
  const countLabel = options.countLabel || "생성";
  const fileLabel = options.fileLabel || "파일";
  const totalItems = data.source_count ?? data.set_count ?? (sourceResults.filter((item) => !item.is_summary).length || 1);
  const failedCount = data.failed_count ?? sourceResults.filter((item) => !item.ok && !item.is_summary).length;
  const generatedCount = data.count ?? sourceResults.reduce((sum, item) => sum + Number(item.count || 0), 0);

  function sourceTitle(item, index) {
    if (item.source_tc || item.source_ui) return item.source_pdf || `${item.source_tc || "-"} + ${item.source_ui || "-"}`;
    return item.source_pdf || `${itemLabel} ${index + 1}`;
  }

  function filesForSource(item) {
    if (!sourceResults.length) return files;
    const matched = files.filter((file) => {
      if (item.source_pdf && file.source_pdf === item.source_pdf) return true;
      if (item.source_tc && file.source_tc === item.source_tc) return true;
      if (item.source_ui && file.source_ui === item.source_ui) return true;
      return false;
    });
    if (matched.length) return matched;
    return sourceResults.length === 1 ? files : [];
  }

  function renderDownloadButtons(downloadFiles) {
    if (!downloadFiles.length) {
      return `<p class="folder-qa-empty">생성된 다운로드 파일이 없습니다.</p>`;
    }
    return `
      <div class="download-list individual-result-files">
        ${downloadFiles.map((file) => `
          <a class="download-button" href="${escapeHtml(file.download_url)}" download="${escapeHtml(file.download_name || file.name || "")}">
            <span>${escapeHtml((file.kind || "file").toUpperCase())}</span>
            <strong>${escapeHtml(file.download_name || file.name || "download")}</strong>
          </a>
        `).join("")}
      </div>
    `;
  }

  const resultItems = sourceResults.length ? sourceResults : [{
    source_pdf: title,
    ok: Boolean(data.ok),
    count: generatedCount,
    file_count: files.length,
    error: data.error || "",
    analysis: {},
  }];

  emptyState.hidden = true;
  clearError();
  downloadPanel.hidden = false;
  downloadPanel.className = "download-panel";
  downloadPanel.innerHTML = `
    <div class="download-panel-head">
      <strong>${escapeHtml(title)}</strong>
      <span>${escapeHtml(itemLabel)} ${escapeHtml(totalItems)}개 · ${escapeHtml(fileLabel)} ${escapeHtml(files.length)}개${failedCount ? ` · 확인 필요 ${escapeHtml(failedCount)}개` : ""}</span>
    </div>
    <div class="source-result-overview">
      <span>처리 ${escapeHtml(itemLabel)} <strong>${escapeHtml(totalItems)}</strong></span>
      <span>${escapeHtml(countLabel)} <strong>${escapeHtml(generatedCount)}</strong></span>
      <span>${escapeHtml(fileLabel)} <strong>${escapeHtml(files.length)}</strong></span>
      <span>확인 필요 <strong>${escapeHtml(failedCount)}</strong></span>
    </div>
    <div class="folder-qa-result-list">
      ${resultItems.map((item, index) => {
        const analysis = item.analysis || {};
        const isOk = Boolean(item.ok);
        const isSummary = Boolean(item.is_summary);
        const cardStatus = isSummary ? "status-queued" : isOk ? "status-done" : "status-error";
        const statusText = isSummary ? "요약" : isOk ? "완료" : "확인 필요";
        const downloadFiles = filesForSource(item);
        const risks = analysis.risks || [];
        const recommendations = analysis.recommendations || [];
        const screens = analysis.screens || [];
        const screenIds = screens
          .map((screen) => typeof screen === "string" ? screen : screen?.screen_id)
          .filter(Boolean)
          .slice(0, 12);
        return `
          <article class="folder-qa-requirement ${isOk || isSummary ? "" : "has-errors"} ${cardStatus}">
            <div class="folder-qa-requirement-head">
              <div>
                <strong>${escapeHtml(sourceTitle(item, index))}</strong>
                <small>${escapeHtml(analysis.summary || (isOk ? "생성이 완료되었습니다." : item.error || "확인이 필요합니다."))}</small>
              </div>
              <div class="folder-qa-requirement-actions">
                <span>${escapeHtml(statusText)}</span>
              </div>
            </div>
            <div class="source-result-metrics">
              <span>${escapeHtml(countLabel)} <strong>${escapeHtml(item.count ?? 0)}</strong></span>
              <span>${escapeHtml(fileLabel)} <strong>${escapeHtml(downloadFiles.length || item.file_count || 0)}</strong></span>
              ${screenIds.length ? `<span>화면 <strong>${escapeHtml(screenIds.length)}</strong></span>` : ""}
            </div>
            ${screenIds.length ? `
              <div class="source-screen-list" aria-label="화면ID 목록">
                ${screenIds.map((screenId) => `<span>${escapeHtml(screenId)}</span>`).join("")}
              </div>
            ` : ""}
            ${item.error && !isOk ? `<p class="folder-qa-empty">${escapeHtml(formatUserErrorMessage(item.error))}</p>` : ""}
            ${risks.length ? `
              <div class="source-result-notes">
                <b>확인할 내용</b>
                <ul>${risks.map((risk) => `<li>${escapeHtml(risk)}</li>`).join("")}</ul>
              </div>
            ` : ""}
            ${recommendations.length ? `
              <div class="source-result-notes">
                <b>${isSummary ? "매칭된 세트" : "권장 조치"}</b>
                <ul>${recommendations.map((recommendation) => `<li>${escapeHtml(recommendation)}</li>`).join("")}</ul>
              </div>
            ` : ""}
            ${isSummary ? "" : renderDownloadButtons(downloadFiles)}
          </article>
        `;
      }).join("")}
    </div>
  `;
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

function selectedFileNames(inputId) {
  const input = document.querySelector(`#${inputId}`);
  return Array.from(input?.files || []).map((file) => fileDisplayPath(file));
}

function renderIndividualGenerationProgress(options = {}) {
  const title = options.title || "생성 진행 중";
  const itemLabel = options.itemLabel || "문서";
  const items = options.items?.length ? options.items : [{ name: itemLabel }];

  emptyState.hidden = true;
  clearError();
  downloadPanel.hidden = false;
  downloadPanel.className = "download-panel";
  downloadPanel.innerHTML = `
    <div class="download-panel-head">
      <strong>${escapeHtml(title)}</strong>
      <span>${escapeHtml(itemLabel)} ${escapeHtml(items.length)}개 · 처리 중</span>
    </div>
    <div class="folder-qa-result-list">
      ${items.map((item, index) => `
        <article class="folder-qa-requirement status-running">
          <div class="folder-qa-requirement-head">
            <div>
              <strong>${escapeHtml(item.name || `${itemLabel} ${index + 1}`)}</strong>
              <small>${escapeHtml(item.detail || "생성 요청을 처리하고 있습니다.")}</small>
            </div>
            <div class="folder-qa-requirement-actions">
              <span>진행 중</span>
            </div>
          </div>
        </article>
      `).join("")}
    </div>
  `;
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
          "화면/사용자인터페이스설계서는 PDF, HWP 또는 HWPX 파일을 직접 업로드하세요.",
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

    if (rawMessage.includes("사용자인터페이스 설계서 문서를 선택하세요") || rawMessage.includes("사용자인터페이스 설계서 PDF를 선택하세요") || rawMessage.includes("ui_pdf") || rawMessage.includes("사용자인터페이스설계서 파일을 찾을 수 없습니다") || rawMessage.includes("PDF 파일을 찾을 수 없습니다")) {
      return {
        summary: "'사용자인터페이스설계서' 파일을 확인하세요.",
        checks: [
          "'사용자인터페이스설계서' 칸에는 PDF, HWP 또는 HWPX 파일을 넣어야 합니다.",
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
  const info = friendlyErrorInfo(formatUserErrorMessage(error.message), taskName);
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
    const error = new Error(data.error || "처리 실패");
    error.data = data;
    throw error;
  }
  return data;
}

async function cancelQaFolderJob(jobId) {
  if (!jobId || cancelRequestedJobIds.has(jobId)) return;
  cancelRequestedJobIds.add(jobId);
  setBadge("취소중", "busy");
  resultMeta.textContent = "QA 생성 취소를 요청했습니다. 현재 처리 중인 단계가 정리되면 중단됩니다.";
  await postJson("/api/cancel-request", { request_id: jobId });
}

async function retryQaFolderRequirement(jobId, requirementId) {
  if (!jobId || !requirementId) return null;
  setBadge("재생성중", "busy");
  resultMeta.textContent = `${requirementId} 실패 화면 재생성을 시작했습니다.`;
  return postJson("/api/retry-qa-folder-block", {
    job_id: jobId,
    requirement_id: requirementId,
  });
}

async function fetchQaFolderJob(jobId, request) {
  const response = await fetch(`/api/qa-folder-job?id=${encodeURIComponent(jobId)}`, {
    cache: "no-store",
    signal: request?.controller.signal,
  });
  const data = await response.json().catch(() => ({ error: "서버 응답을 읽지 못했습니다." }));
  if (!response.ok) {
    const error = new Error(data.error || "QA job 상태 조회 실패");
    error.data = data;
    throw error;
  }
  return data;
}

function wait(ms, request) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(resolve, ms);
    const signal = request?.controller?.signal;
    signal?.addEventListener("abort", () => {
      clearTimeout(timer);
      reject(new DOMException("Aborted", "AbortError"));
    }, { once: true });
  });
}

async function pollQaFolderJob(jobId, request) {
  let lastData = null;
  rememberQaFolderJob(jobId);
  while (true) {
    const data = await fetchQaFolderJob(jobId, request);
    lastData = data;
    rememberQaFolderJob(data.job_id || data.request_id || jobId);
    renderFolderQaResult(data);

    const total = data.requirement_count ?? 0;
    const done = data.processed_requirement_count ?? 0;
    const failed = data.failed_requirement_count ?? 0;
    resultMeta.textContent = `QA 생성 진행 중: 완료 ${done}/${total}건 · 실패 ${failed}건`;

    if (["done", "error", "cancelled"].includes(data.status)) {
      return data;
    }
    await wait(1500, request);
  }
}

async function restoreLastQaFolderJob() {
  const jobId = getRememberedQaFolderJob();
  if (!jobId) return false;

  const request = beginCancelableRequest();
  try {
    const data = await fetchQaFolderJob(jobId, request);
    rememberQaFolderJob(data.job_id || data.request_id || jobId);
    resultTitle.textContent = "QA 생성";

    if (["queued", "running"].includes(data.status || "")) {
      setBadge("처리중", "busy");
      setLoading(true);
      resultMeta.textContent = "이전 QA 생성 작업을 이어서 확인하고 있습니다.";
      renderFolderQaResult(data);

      const finalData = await pollQaFolderJob(data.job_id || data.request_id || jobId, request);
      setQaFolderFinalStatus(finalData);
      renderFolderQaResult(finalData);
      return true;
    }

    if (isTerminalQaFolderStatus(data.status)) {
      setQaFolderFinalStatus(data);
      renderFolderQaResult(data);
      return true;
    }

    renderFolderQaResult(data);
    resultMeta.textContent = "최근 QA 생성 작업 상태를 확인했습니다.";
    return true;
  } catch (error) {
    if (!isAbortError(error)) {
      forgetQaFolderJob(jobId);
    }
    return false;
  } finally {
    endCancelableRequest(request);
    setLoading(false);
  }
}

async function fetchPathStatus(path) {
  try {
    const response = await fetch(`/api/path-status?path=${encodeURIComponent(path)}`);
    const data = await response.json().catch(() => ({ ok: false }));
    if (!response.ok) {
      return { ok: false, ...data };
    }
    return data;
  } catch (error) {
    return { ok: false, error: String(error) };
  }
}

function forgetDumpRoot(path) {
  try {
    const stored = localStorage.getItem(LAST_DUMP_ROOT_KEY) || "";
    if (!path || stored === path) {
      localStorage.removeItem(LAST_DUMP_ROOT_KEY);
    }
  } catch {
    // localStorage를 쓸 수 없는 환경이면 입력값만 정리한다.
  }
}

async function initializeDumpRoot() {
  if (!qaDumpRoot) return;

  const params = new URLSearchParams(window.location.search);
  const queryDumpRoot = params.get("dump_root") || "";
  let storedDumpRoot = "";
  try {
    storedDumpRoot = localStorage.getItem(LAST_DUMP_ROOT_KEY) || "";
  } catch {
    storedDumpRoot = "";
  }

  const initialDumpRoot = queryDumpRoot || storedDumpRoot;
  qaDumpRoot.value = initialDumpRoot;
  if (!initialDumpRoot) return;

  const status = await fetchPathStatus(initialDumpRoot);
  if (status.ok) return;

  qaDumpRoot.value = "";
  forgetDumpRoot(initialDumpRoot);
  showGenerationError(
    new Error("이전에 저장된 산출물 매핑 결과 폴더를 찾을 수 없습니다. 다시 선택하거나 산출물 폴더를 업로드하세요."),
    "folder",
  );
}

function renderFolderQaResult(data) {
  const placedFiles = Array.isArray(data.placed_files) ? data.placed_files : [];
  const sourceFiles = Array.isArray(data.source_files) ? data.source_files : [];
  const requirementItems = Array.isArray(data.requirement_items) ? data.requirement_items : [];
  const missingRequirements = Array.isArray(data.missing_requirements) ? data.missing_requirements : [];
    const roleCounts = data.role_counts || {};
    const isPreview = Boolean(data.match_preview);
    const jobId = data.job_id || data.request_id || "";
    const canCancelJob = !isPreview && jobId && ["queued", "running"].includes(data.status || "");
    const cancelRequested = jobId && cancelRequestedJobIds.has(jobId);
    const showFileDetails = isPreview || isTerminalQaFolderStatus(data.status);
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

  function sortSourceFiles(files) {
    return [...files].sort((left, right) => {
      const leftRole = roleOrder.indexOf(left.role);
      const rightRole = roleOrder.indexOf(right.role);
      return (leftRole < 0 ? 99 : leftRole) - (rightRole < 0 ? 99 : rightRole);
    });
  }

  function renderFileDetails(sources, outputs) {
    const sortedSources = sortSourceFiles(sources);
    const backupFiles = outputs.filter((file) => file.backup_path);
    const summaryText = [
      `입력 ${sources.length}개`,
      isPreview ? "" : `생성 ${outputs.length}개`,
      !isPreview && backupFiles.length ? "백업 있음" : "",
    ].filter(Boolean).join(" · ");

    const sourceRows = sortedSources.map((file) => renderPathItem(file.path, {
      pathLabel: "파일 위치",
    })).join("");
    const outputRows = outputs.map((file) => renderPathItem(file.path, {
      pathLabel: "배치 위치",
      backupPath: file.backup_path || "",
    })).join("");

    if (!sourceRows && !outputRows) {
      return `
        <div class="folder-qa-file-summary">
          <span>파일 요약</span>
          <strong>매칭된 파일 없음</strong>
        </div>
      `;
    }

    function renderPathItem(path, options = {}) {
      const folder = dirname(path);
      const backupFolder = dirname(options.backupPath || "");
      const hasPathDetails = Boolean(folder || backupFolder);
      return `
        <article class="folder-qa-path-item">
          <strong>${escapeHtml(basename(path))}</strong>
          ${hasPathDetails ? `
            <details class="folder-qa-path-inline">
              <summary>
                ${options.backupPath ? `<em>백업 있음</em>` : ""}
                <span>경로 보기</span>
              </summary>
            </details>
            <div class="folder-qa-path-panel">
              ${folder ? `<code title="${escapeHtml(path || "")}"><span>${escapeHtml(options.pathLabel || "파일 위치")}</span>${escapeHtml(folder)}</code>` : ""}
              ${backupFolder ? `<code title="${escapeHtml(options.backupPath || "")}"><span>백업 위치</span>${escapeHtml(backupFolder)}</code>` : ""}
            </div>
          ` : ""}
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
          <span>파일 상세 보기</span>
          <small>${escapeHtml(summaryText)}</small>
        </summary>
        <div class="folder-qa-path-groups">
          ${renderPathGroup(isPreview ? "매칭된 입력 파일" : "입력 문서", sourceRows)}
          ${isPreview ? "" : renderPathGroup("생성/배치 결과", outputRows)}
        </div>
      </details>
    `;
  }

  function renderBlockTree(item) {
    const blocks = Array.isArray(item.blocks) ? item.blocks : [];
    if (!blocks.length || isPreview) return "";
    const statusLabels = {
      queued: "대기",
      running: "진행 중",
      updated: "완료",
      error: "실패",
      interrupted: "중단",
    };

    return `
      <div class="folder-qa-block-tree">
        <div class="folder-qa-block-list">
          ${blocks.map((block) => {
            const status = block.status || "queued";
            const statusClass = ["queued", "running", "updated", "error", "interrupted"].includes(status) ? status : "queued";
            const blockLabel = block.screen_id || block.unit_test_id || `블록 ${Number(block.display_index || 0) || "-"}`;
            const blockMeta = [
              block.expected_steps !== undefined && block.expected_steps !== null ? `처리흐름 ${block.expected_steps}개` : "",
              block.num_predict !== undefined && block.num_predict !== null ? `num_predict ${block.num_predict}` : "",
              block.timeout !== undefined && block.timeout !== null ? `timeout ${block.timeout}초` : "",
              block.generated_count !== undefined && block.generated_count !== null ? `생성 ${block.generated_count}건` : "",
            ].filter(Boolean).join(" · ");

            const blockError = formatUserErrorMessage(block.error);

            return `
              <article class="folder-qa-block ${statusClass}">
                <div class="folder-qa-block-main">
                  <span class="folder-qa-block-title">
                    <strong>${escapeHtml(blockLabel)}</strong>
                    ${blockMeta ? `<small>${escapeHtml(blockMeta)}</small>` : ""}
                  </span>
                  ${blockError ? `<p>${escapeHtml(blockError)}</p>` : ""}
                </div>
                <div class="folder-qa-block-actions">
                  <span class="folder-qa-block-status ${statusClass}">${escapeHtml(statusLabels[statusClass] || "대기")}</span>
                </div>
              </article>
            `;
          }).join("")}
        </div>
      </div>
    `;
  }

  function getBlockSummary(item) {
    const blocks = Array.isArray(item.blocks) ? item.blocks : [];
    if (!blocks.length || isPreview) return "";
    const statusCounts = blocks.reduce((counts, block) => {
      const status = block.status || "queued";
      counts[status] = (counts[status] || 0) + 1;
      return counts;
    }, {});
    return [
      `${blocks.length}개 화면`,
      statusCounts.updated ? `완료 ${statusCounts.updated}` : "",
      statusCounts.running ? `진행 ${statusCounts.running}` : "",
      statusCounts.error ? `실패 ${statusCounts.error}` : "",
      statusCounts.interrupted ? `중단 ${statusCounts.interrupted}` : "",
      statusCounts.queued ? `대기 ${statusCounts.queued}` : "",
    ].filter(Boolean).join(" · ");
  }

  function getFailedBlocks(item) {
    const blocks = Array.isArray(item.blocks) ? item.blocks : [];
    return blocks.filter((block) => block.status === "error");
  }

  function renderRequirementCards() {
    if (!requirementIds.length) {
      if (isActiveQaFolderStatus(data.status)) {
        return `
          <article class="folder-qa-requirement status-running">
            <div class="folder-qa-requirement-head">
              <strong>입력 파일 확인 중</strong>
              <span>진행 중</span>
            </div>
            <p class="folder-qa-empty">매칭된 요구사항을 불러오고 있습니다.</p>
          </article>
        `;
      }
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
      const isRunning = item.status === "running";
      const isUpdated = item.status === "updated";
      const sources = sourceByRequirement.get(requirementId) || [];
      const outputs = placedByRequirement.get(requirementId) || [];
      const blockSummary = getBlockSummary(item);
      const failedBlocks = getFailedBlocks(item);
      const canRetryRequirement = jobId && ["error", "done"].includes(data.status || "") && failedBlocks.length > 0;
      const retryRequirementPayload = encodeURIComponent(JSON.stringify({
        job_id: jobId,
        requirement_id: requirementId,
      }));
      const statusText = isError
        ? "실패"
        : isRunning
        ? "진행 중"
        : isUpdated || outputs.length
        ? "배치 완료"
        : isPreview
        ? "매칭 완료"
        : "대기";
      const statusClass = isError
        ? "status-error"
        : isRunning
        ? "status-running"
        : isUpdated || outputs.length
        ? "status-done"
        : isPreview
        ? "status-done"
        : "status-queued";
      return `
        <article class="folder-qa-requirement ${isError ? "has-errors" : ""} ${statusClass}">
          <div class="folder-qa-requirement-head">
            <div>
              <strong>${escapeHtml(requirementId)}</strong>
              ${blockSummary ? `<small>${escapeHtml(blockSummary)}</small>` : ""}
            </div>
            <div class="folder-qa-requirement-actions">
              ${canRetryRequirement ? `
                <button class="folder-qa-retry-button" type="button" data-retry-qa-requirement="${escapeHtml(retryRequirementPayload)}">재생성</button>
              ` : ""}
              <span>${escapeHtml(statusText)}</span>
            </div>
          </div>
          ${renderBlockTree(item)}
          ${showFileDetails ? renderFileDetails(sources, outputs) : ""}
        </article>
      `;
    }).join("");
  }

  function renderJobSummary() {
    if (isPreview) {
      return [
        `${data.requirement_count ?? 0}개 요구사항`,
        `입력 ${sourceFiles.length}개`,
      ].join(" · ");
    }
    return [
      `${data.processed_requirement_count ?? 0}/${data.requirement_count ?? 0}개 요구사항`,
      `입력 ${sourceFiles.length}개`,
      `생성 ${placedFiles.length}개`,
    ].join(" · ");
  }

  function renderRoleCounts() {
    if (!missingRequirements.length) return "";
    return `
      <div class="role-count-list">
        <article>
          <span>사용자인터페이스설계서</span>
          <strong>${escapeHtml(roleCounts.ui_design ?? "-")}</strong>
        </article>
        <article>
          <span>단위시험 케이스</span>
          <strong>${escapeHtml(roleCounts.tc_template ?? "-")}</strong>
        </article>
        <article>
          <span>단위시험 결과서</span>
          <strong>${escapeHtml(roleCounts.unit_result_template ?? "-")}</strong>
        </article>
        <article>
          <span>통합시험 시나리오</span>
          <strong>${escapeHtml(roleCounts.ts_template ?? "-")}</strong>
        </article>
        <article>
          <span>통합시험 결과서</span>
          <strong>${escapeHtml(roleCounts.integration_result_template ?? "-")}</strong>
        </article>
      </div>
    `;
  }

  downloadPanel.hidden = false;
  downloadPanel.innerHTML = `
    <div class="download-panel-head">
      <strong>${isPreview ? "QA 매칭 확인" : data.ok === false ? "QA 생성 현황" : "QA 배치 결과"}</strong>
      <div class="folder-qa-job-actions">
        <span>${escapeHtml(renderJobSummary())}</span>
        ${canCancelJob ? `
          <button class="folder-qa-cancel-button" type="button" data-cancel-qa-job="${escapeHtml(jobId)}" ${cancelRequested ? "disabled" : ""}>
            ${cancelRequested ? "취소 요청됨" : "취소"}
          </button>
        ` : ""}
      </div>
    </div>
    ${renderRoleCounts()}
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

async function validateQaTargetSelection() {
  const dumpRoot = qaDumpRoot?.value?.trim() || "";
  const sourceRoot = qaSourceRoot?.value?.trim() || "";
  const hasSourceUpload = Boolean(qaSourceFiles?.files?.length);
  if (!dumpRoot && !sourceRoot && !hasSourceUpload) {
    showGenerationError(new Error("QA 대상 산출물 폴더 경로를 입력하거나 산출물 폴더를 업로드하세요."), "folder");
    return null;
  }

  const selectedFolder = sourceRoot
    ? { path: sourceRoot, label: "다른 산출물 폴더", input: qaSourceRoot }
    : hasSourceUpload
    ? null
    : { path: dumpRoot, label: "산출물 매핑 결과 폴더", input: qaDumpRoot };
  if (selectedFolder?.path) {
    const status = await fetchPathStatus(selectedFolder.path);
    if (!status.ok) {
      if (selectedFolder.input === qaDumpRoot) {
        qaDumpRoot.value = "";
        forgetDumpRoot(selectedFolder.path);
      }
      if (selectedFolder.input === qaSourceRoot) {
        qaSourceRoot.value = "";
      }
      showGenerationError(new Error(`${selectedFolder.label}를 찾을 수 없습니다: ${selectedFolder.path}`), "folder");
      return null;
    }
  }

  if (dumpRoot && !sourceRoot && !hasSourceUpload) {
    try {
      localStorage.setItem(LAST_DUMP_ROOT_KEY, dumpRoot);
    } catch {
      // 저장 실패는 실행을 막지 않는다.
    }
  }

  return { dumpRoot, sourceRoot, hasSourceUpload };
}

async function previewFolderQa() {
  const selection = await validateQaTargetSelection();
  if (!selection) return;

  setBadge("확인중", "busy");
  setLoading(true);
  clearDownloads();
  clearError();
  resultTitle.textContent = "QA 매칭 확인";
  resultMeta.textContent = "요청 처리 중입니다.";

  const request = beginCancelableRequest();
  try {
    const data = await postForm("/api/preview-qa-folder", buildFormData(folderQaForm), request);
    setBadge(data.ok ? "확인 완료" : "확인 필요", data.ok ? "done" : "error");
    resultTitle.textContent = "QA 매칭 확인";
    resultMeta.textContent = `매칭된 요구사항 ${data.requirement_count ?? 0}개 · 생성 전 파일만 확인했습니다.`;
    renderFolderQaResult(data);
  } catch (error) {
    if (isAbortError(error)) {
      markRequestCanceled();
      return;
    }
    showGenerationError(error, "folder");
    if (error.data?.source_files || error.data?.missing_requirements) {
      resultTitle.textContent = "QA 매칭 확인";
      resultMeta.textContent = "매칭된 파일과 누락 항목을 확인하세요.";
      renderFolderQaResult(error.data);
    }
  } finally {
    endCancelableRequest(request);
    setLoading(false);
  }
}

async function runFolderQa(event) {
  event.preventDefault();
  const selection = await validateQaTargetSelection();
  if (!selection) return;
  const { dumpRoot, sourceRoot } = selection;

  setBadge("처리중", "busy");
  setLoading(true);
  clearDownloads();
  clearError();
  resultTitle.textContent = "QA 생성";
  resultMeta.textContent = "요청 처리 중입니다.";

  const request = beginCancelableRequest();

  try {
    const job = await postForm("/api/run-qa-folder-job", buildFormData(folderQaForm), request);
    const jobId = job.job_id || job.request_id || "";
    rememberQaFolderJob(jobId);
    resultMeta.textContent = `QA 생성 작업을 시작했습니다: ${jobId || "-"}`;
    const data = await pollQaFolderJob(jobId, request);
    setQaFolderFinalStatus(data, dumpRoot || sourceRoot);
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
  setLoading(true);
  clearDownloads();
  clearError();
  resultTitle.textContent = "단위시험 케이스 생성";
  resultMeta.textContent = "요청 처리 중입니다.";
  renderIndividualGenerationProgress({
    title: "단위시험 케이스 생성 진행",
    itemLabel: "설계서",
    items: selectedFileNames("tcUiPdf").map((name) => ({
      name,
      detail: "단위시험 케이스를 생성하고 있습니다.",
    })),
  });

  const request = beginCancelableRequest();

  try {
    const data = await postForm("/api/generate-tc", buildFormData(tcForm), request);
    setBadge("완료", "done");
    const sourceCount = data.source_count ?? data.source_results?.length ?? 1;
    const failedCount = data.failed_count ?? 0;
    resultMeta.textContent = `처리 설계서 ${sourceCount}개 · 생성 행 수 ${data.count ?? 0}개 · 다운로드 ${data.download_files?.length || 0}개${failedCount ? ` · 실패 ${failedCount}개` : ""}`;
    renderIndividualGenerationResult(data, {
      title: "단위시험 케이스 생성 결과",
      itemLabel: "설계서",
      countLabel: "생성 행",
      fileLabel: "다운로드",
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
  setLoading(true);
  clearDownloads();
  clearError();
  resultTitle.textContent = "통합시험 시나리오 생성";
  resultMeta.textContent = "요청 처리 중입니다.";
  const tcNames = selectedFileNames("tsTcXlsx").map((name) => ({
    name,
    detail: "통합시험 시나리오 생성을 준비하고 있습니다.",
  }));
  const uiNames = selectedFileNames("tsUiPdf").map((name) => ({
    name,
    detail: "화면설계서 매칭을 확인하고 있습니다.",
  }));
  renderIndividualGenerationProgress({
    title: "통합시험 시나리오 생성 진행",
    itemLabel: "파일",
    items: [...tcNames, ...uiNames],
  });

  const request = beginCancelableRequest();

  try {
    const data = await postForm("/api/generate-ts", buildFormData(tsForm), request);
    setBadge("완료", "done");
    const setCount = data.set_count ?? data.source_results?.length ?? 1;
    const failedCount = data.failed_count ?? 0;
    resultMeta.textContent = `처리 세트 ${setCount}개 · 다운로드 ${data.download_files?.length || 0}개${failedCount ? ` · 확인 필요 ${failedCount}개` : ""}`;
    renderIndividualGenerationResult(data, {
      title: "통합시험 시나리오 생성 결과",
      itemLabel: "세트",
      countLabel: "생성 행",
      fileLabel: "다운로드",
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
folderPreviewButton?.addEventListener("click", previewFolderQa);
folderQaForm?.addEventListener("submit", runFolderQa);
tcForm.addEventListener("submit", runTcGeneration);
tsForm.addEventListener("submit", runTsGeneration);
downloadPanel?.addEventListener("click", async (event) => {
  const button = event.target instanceof Element
    ? event.target.closest("[data-cancel-qa-job]")
    : null;
  if (!button) return;
  const jobId = button.dataset.cancelQaJob || "";
  button.disabled = true;
  button.textContent = "취소 요청됨";
  try {
    await cancelQaFolderJob(jobId);
  } catch (error) {
    button.disabled = false;
    button.textContent = "취소";
    showGenerationError(error, "folder");
  }
});
downloadPanel?.addEventListener("click", async (event) => {
  const button = event.target instanceof Element
    ? event.target.closest("[data-retry-qa-requirement]")
    : null;
  if (!button) return;

  let payload = {};
  try {
    payload = JSON.parse(decodeURIComponent(button.dataset.retryQaRequirement || "{}"));
  } catch {
    payload = {};
  }

  const jobId = payload.job_id || "";
  const requirementId = payload.requirement_id || "";
  button.disabled = true;
  button.textContent = "재생성 중";

  const request = beginCancelableRequest();
  try {
    const job = await retryQaFolderRequirement(jobId, requirementId);
    const nextJobId = job?.job_id || job?.request_id || jobId;
    rememberQaFolderJob(nextJobId);
    const data = await pollQaFolderJob(nextJobId, request);
    setQaFolderFinalStatus(data);
    renderFolderQaResult(data);
  } catch (error) {
    if (isAbortError(error)) {
      markRequestCanceled();
      return;
    }
    button.disabled = false;
    button.textContent = "재생성";
    showGenerationError(error, "folder");
  } finally {
    endCancelableRequest(request);
  }
});

async function initializeQaPage() {
  setBadge("대기", "");
  loadRuntimeMode();
  await initializeDumpRoot();
  await restoreLastQaFolderJob();
}

initializeQaPage();
