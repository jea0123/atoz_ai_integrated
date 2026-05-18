const statusBadge = document.querySelector("#statusBadge");
const runtimeMode = document.querySelector("#runtimeMode");
const loadingOverlay = document.querySelector("#loadingOverlay");
const loadingTitle = document.querySelector("#loadingTitle");
const loadingMessage = document.querySelector("#loadingMessage");

const resultTitle = document.querySelector("#resultTitle");
const resultMeta = document.querySelector("#resultMeta");
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

  label.textContent = input.files[0].name;
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

function renderDownloads(files, title) {
  const downloadFiles = (files || []).filter((file) => file && file.download_url);
  generatedFileCount.textContent = downloadFiles.length || "-";

  if (!downloadFiles.length) {
    clearDownloads();
    return;
  }

  emptyState.hidden = true;
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

function validateFiles(form, rules) {
  for (const [selector, message] of rules) {
    const input = form.querySelector(selector);
    if (!input || !input.files.length) {
      setBadge("확인", "error");
      setProcessStatus("확인 필요");
      resultMeta.textContent = message;
      return false;
    }
  }
  return true;
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
  resultTitle.textContent = "단위시험 케이스 생성";
  resultMeta.textContent = "요청 처리 중입니다.";

  try {
    const data = await postForm("/api/generate-tc", buildFormData(tcForm));
    setBadge("완료", "done");
    setProcessStatus("완료");
    tcCount.textContent = data.count ?? 0;
    resultMeta.textContent = `생성 행 수 ${data.count ?? 0}개 · 다운로드 ${data.download_files?.length || 0}개`;
    renderDownloads(data.download_files || data.files, "단위시험 케이스 파일");
  } catch (error) {
    setBadge("오류", "error");
    setProcessStatus("오류");
    resultMeta.textContent = error.message;
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
    setBadge("오류", "error");
    setProcessStatus("오류");
    resultMeta.textContent = error.message;
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
