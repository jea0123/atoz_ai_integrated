const templateForm = document.querySelector("#templateForm");
const standardFile = document.querySelector("#standardFile");
const artifactFiles = document.querySelector("#artifactFiles") || document.querySelector("#templateFiles");
const requirementFiles = document.querySelector("#requirementFiles");
const proposalFiles = document.querySelector("#proposalFiles");
const standardName = document.querySelector("#standardName");
const templateName = document.querySelector("#templateName");
const requirementName = document.querySelector("#requirementName");
const proposalName = document.querySelector("#proposalName");
const statusBadge = document.querySelector("#statusBadge");
const runButton = document.querySelector("#runButton");
const applyButton = document.querySelector("#applyButton");
const loadingOverlay = document.querySelector("#loadingOverlay");
const templateCount = document.querySelector("#templateCount");
const outputCount = document.querySelector("#outputCount");
const requirementCount = document.querySelector("#requirementCount");
const buildState = document.querySelector("#buildState");
const resultMeta = document.querySelector("#resultMeta");
const templateRows = document.querySelector("#templateRows");
const emptyState = document.querySelector("#emptyState");
const downloadPanel = document.querySelector("#downloadPanel");

class ValidationError extends Error {
  constructor(message) {
    super(message);
    this.name = "ValidationError";
  }
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function setBadge(text, mode = "") {
  if (!statusBadge) return;
  statusBadge.textContent = text;
  statusBadge.className = `status-pill ${mode}`.trim();
}

function setLoading(isLoading, activeButton = null) {
  loadingOverlay.hidden = !isLoading;
  runButton.disabled = isLoading;
  if (applyButton) applyButton.disabled = isLoading;
  runButton.classList.toggle("is-loading", isLoading && activeButton === runButton);
  applyButton?.classList.toggle("is-loading", isLoading && activeButton === applyButton);
}

function fileSummary(files, fallback) {
  if (!files?.length) return fallback;
  return files.length > 1 ? `${files[0].name} 외 ${files.length - 1}개` : files[0].name;
}

function folderSummary(files, fallback) {
  if (!files?.length) return fallback;
  const firstPath = files[0].webkitRelativePath || files[0].name;
  const rootName = firstPath.includes("/") ? firstPath.split("/")[0] : firstPath;
  return `${rootName} · ${files.length}개`;
}

function syncFileLabels() {
  if (standardName) {
    standardName.textContent = fileSummary(standardFile?.files, "문서관리표준을 선택하세요");
  }
  if (templateName) {
    templateName.textContent = folderSummary(
      artifactFiles?.files,
      templateForm.dataset.category === "management"
        ? "표지를 읽을 관리산출물을 선택하세요"
        : "표지를 읽을 개발산출물을 선택하세요",
    );
  }
  if (requirementName) {
    requirementName.textContent = fileSummary(requirementFiles?.files, "선택 시 파일명에서 SFR ID를 읽습니다");
  }
  if (proposalName) {
    proposalName.textContent = fileSummary(proposalFiles?.files, "선택 시 요구사항목록표 SFR ID를 읽습니다");
  }
}

function buildFormData(mode) {
  const body = new FormData(templateForm);
  body.append("request_id", crypto.randomUUID?.() || String(Date.now()));
  body.append("template_mode", mode);
  if (mode === "apply") {
    body.append("apply_mode", "true");
  }
  return body;
}

function renderDownloads(data) {
  const links = [];
  if (data.download_url) {
    links.push(`<a href="${escapeHtml(data.download_url)}">${escapeHtml(data.download_name || "전체 출력 ZIP")}</a>`);
  }
  for (const file of data.download_files || []) {
    links.push(`<a href="${escapeHtml(file.download_url)}">${escapeHtml(file.download_name || file.name)}</a>`);
  }
  if (!links.length) {
    downloadPanel.hidden = true;
    downloadPanel.innerHTML = "";
    return;
  }
  downloadPanel.hidden = false;
  downloadPanel.innerHTML = `
    <strong>다운로드</strong>
    <div class="template-download-list">${links.join("")}</div>
  `;
}

function renderRows(items) {
  templateRows.innerHTML = items.map((item) => {
    const identity = item.identity || {};
    const standard = item.standard_output || {};
    const statusClass = item.status === "matched" ? "ok" : "warn";
    return `
      <tr>
        <td>
          <strong>${escapeHtml(item.path)}</strong>
          <small>표준: ${escapeHtml(standard.id || "-")} ${escapeHtml(standard.name || "")}</small>
          <small>출력: ${escapeHtml(item.output_name || item.output_relative_path || "-")}</small>
          <small>${escapeHtml(item.output_path || "")}</small>
        </td>
        <td>
          <strong>${escapeHtml(item.body_template || "매칭 없음")}</strong>
          <small>${escapeHtml(item.body_match_type || "")}</small>
          <span class="template-cover-type">${escapeHtml(item.cover_type || "-")}</span>
          <small>${escapeHtml(item.cover_template || "표지 템플릿 없음")}</small>
        </td>
        <td>
          <strong>${escapeHtml(identity.document_title || "-")}</strong>
          <small>${escapeHtml(identity.document_number || "")}</small>
          <small>${escapeHtml(identity.project_title || "")}</small>
        </td>
        <td>
          <span class="template-status ${statusClass}">${escapeHtml(item.status)}</span>
          <small>${escapeHtml(item.message || "")}</small>
        </td>
      </tr>
    `;
  }).join("");
  emptyState.hidden = Boolean(items.length);
}

function validateInputs() {
  if (!standardFile?.files?.length) {
    throw new ValidationError("문서관리표준 파일을 선택하세요.");
  }
  if (!artifactFiles?.files?.length) {
    throw new ValidationError("산출물 입력폴더를 선택하세요.");
  }
}

async function runTemplateRequest(mode) {
  const isApply = mode === "apply";
  try {
    validateInputs();
    setLoading(true, isApply ? applyButton : runButton);
    setBadge("처리 중", "busy");
    buildState.textContent = "처리 중";
    resultMeta.textContent = isApply
      ? "매칭 결과를 기준으로 출력 폴더에 반영하는 중입니다."
      : "문서관리표준 ID, 입력 산출물 표지, 서버 템플릿 폴더를 매칭하는 중입니다.";
    const response = await fetch("/api/template-build", {
      method: "POST",
      body: buildFormData(mode),
    });
    const data = await response.json();
    if (!response.ok || data.ok === false) {
      throw new Error(data.error || (isApply ? "템플릿 반영 실패" : "템플릿 매칭 실패"));
    }

    templateCount.textContent = data.standard_output_count ?? 0;
    outputCount.textContent = data.artifact_count ?? 0;
    requirementCount.textContent = data.requirement_count ?? 0;
    buildState.textContent = "완료";
    resultMeta.textContent = isApply
      ? `반영 폴더: ${data.dump_root || "-"}`
      : `매칭 완료: ${data.matched_count ?? 0}건 매칭`;
    renderRows(data.items || []);
    renderDownloads(data);
    setBadge("완료", "done");
  } catch (error) {
    const isValidation = error instanceof ValidationError;
    buildState.textContent = isValidation ? "입력 필요" : "실패";
    resultMeta.textContent = error.message;
    renderRows([]);
    renderDownloads({});
    setBadge(isValidation ? "입력 필요" : "오류", isValidation ? "warn" : "error");
  } finally {
    setLoading(false);
  }
}

async function submitTemplateBuild(event) {
  event.preventDefault();
  await runTemplateRequest("match");
}

for (const input of [standardFile, artifactFiles, requirementFiles, proposalFiles]) {
  input?.addEventListener("change", syncFileLabels);
}
templateForm?.addEventListener("submit", submitTemplateBuild);
applyButton?.addEventListener("click", async () => {
  await runTemplateRequest("apply");
});
syncFileLabels();
setBadge("대기", "");
