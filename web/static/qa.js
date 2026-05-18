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
  generatedFileCount.textContent = "-";
}

function clearError() {
  if (!errorPanel) return;
  errorPanel.hidden = true;
  errorPanel.replaceChildren();
}

function renderDownloads(files, title) {
  const downloadFiles = (files || []).filter((file) => file && file.download_url);
  generatedFileCount.textContent = downloadFiles.length || "-";

  if (!downloadFiles.length) {
    clearDownloads();
    return;
  }

  emptyState.hidden = true;
  clearError();
  downloadPanel.hidden = false;
  downloadPanel.innerHTML = `
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
  `;
}

function renderSourceResults(sourceResults) {
  if (!sourceResults?.length || downloadPanel.hidden) return;

  const rows = sourceResults.map((item) => {
    const analysis = item.analysis || {};
    const risks = analysis.risks || [];
    const recommendations = analysis.recommendations || [];
    const screens = analysis.screens || [];
    const quality = analysis.quality || (item.ok ? "good" : "warning");

    return `
      <article class="source-result-card">
        <div class="source-result-head">
          <strong>${escapeHtml(item.source_pdf || "-")}</strong>
          <span class="${escapeHtml(quality)}">${item.ok ? "생성 완료" : "확인 필요"}</span>
        </div>
        <p>${escapeHtml(analysis.summary || `${analysis.screen_count ?? 0}개 화면 분석 · ${item.count ?? 0}개 생성`)}</p>
        <small>화면 ${escapeHtml(analysis.screen_count ?? screens.length ?? 0)}개 · 생성 행 ${escapeHtml(item.count ?? 0)}개</small>
        ${risks.length ? `
          <ul>
            ${risks.map((risk) => `<li>${escapeHtml(risk)}</li>`).join("")}
          </ul>
        ` : ""}
        ${recommendations.length ? `
          <ul>
            ${recommendations.map((recommendation) => `<li>${escapeHtml(recommendation)}</li>`).join("")}
          </ul>
        ` : ""}
      </article>
    `;
  }).join("");

  downloadPanel.insertAdjacentHTML("beforeend", `
    <div class="source-result-list">
      <div class="download-panel-head">
        <strong>사전 분석 요약</strong>
        <span>${sourceResults.length}개</span>
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
    tcCount.textContent = data.count ?? 0;
    const sourceCount = data.source_count ?? data.source_results?.length ?? 1;
    const failedCount = data.failed_count ?? 0;
    resultMeta.textContent = `처리 설계서 ${sourceCount}개 · 생성 행 수 ${data.count ?? 0}개 · 다운로드 ${data.download_files?.length || 0}개${failedCount ? ` · 실패 ${failedCount}개` : ""}`;
    renderDownloads(data.download_files || data.files, "단위시험 케이스 파일");
    renderSourceResults(data.source_results || []);
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
    tsCount.textContent = data.count ?? 0;
    resultMeta.textContent = `생성 행 수 ${data.count ?? 0}개 · 다운로드 ${data.download_files?.length || 0}개`;
    renderDownloads(data.download_files || data.files, "통합시험 시나리오 파일");
  } catch (error) {
    showGenerationError(error, "ts");
  } finally {
    setLoading(false);
  }
}

taskTabs.forEach((tab) => tab.addEventListener("click", () => setActiveTask(tab.dataset.taskTab)));
fileInputs.forEach((input) => input.addEventListener("change", () => updateFileLabel(input)));
tcForm.addEventListener("submit", runTcGeneration);
tsForm.addEventListener("submit", runTsGeneration);

setBadge("대기", "");
setProcessStatus("대기");
loadRuntimeMode();
