// 폴더 산출물 검사 화면입니다. 업로드, API 호출, 매칭 결과 표시를 담당합니다.
const form = document.querySelector("#checkForm");
const standardFile = document.querySelector("#standardFile");
const folderFiles = document.querySelector("#folderFiles");
const standardName = document.querySelector("#standardName");
const folderName = document.querySelector("#folderName");
const statusBadge = document.querySelector("#statusBadge");
const runButton = document.querySelector("#runButton");
const applyButton = document.querySelector("#applyButton");
const outputCount = document.querySelector("#outputCount");
const fileCount = document.querySelector("#fileCount");
const matchCount = document.querySelector("#matchCount");
const matchedFileCount = document.querySelector("#matchedFileCount");
const resultMeta = document.querySelector("#resultMeta");
const resultFilter = document.querySelector("#resultFilter");
const results = document.querySelector("#results");
const resultRows = document.querySelector("#resultRows");
const emptyState = document.querySelector("#emptyState");
const visibleCount = document.querySelector("#visibleCount");
const loadingOverlay = document.querySelector("#loadingOverlay");
const applyReport = document.querySelector("#applyReport");
const runtimeMode = document.querySelector("#runtimeMode");

const expandedOutputs = new Set();
const excludedCandidatePaths = new Set();
let lastMatches = [];
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
  statusBadge.textContent = text;
  statusBadge.className = `status-pill ${mode || ""}`.trim();
}

function modeLabel(mode) {
  const labels = {
    ai_first: "AI 우선",
    rule_fallback_no_ollama: "규칙 기반",
  };
  return labels[mode] || mode || "-";
}

function modeClass(mode) {
  if (mode === "ai_first") return "ai";
  return "rule";
}

function setRuntimeMode(data) {
  currentRuntimeMode = data || currentRuntimeMode;
  if (!runtimeMode || !currentRuntimeMode) return;

  const mode = currentRuntimeMode.mode || "unknown";
  const modelText = currentRuntimeMode.model ? ` · ${currentRuntimeMode.model}` : "";
  runtimeMode.className = `mode-note ${modeClass(mode)}`;
  runtimeMode.textContent = `${currentRuntimeMode.label || modeLabel(mode)}${modelText}`;
}

async function loadRuntimeMode() {
  try {
    const response = await fetch("/api/runtime-mode", { cache: "no-store" });
    if (!response.ok) throw new Error("runtime mode request failed");
    setRuntimeMode(await response.json());
  } catch {
    setRuntimeMode({
      mode: "unknown",
      label: "실행 모드 확인 실패",
    });
  }
}

function setLoading(isLoading) {
  loadingOverlay.hidden = !isLoading;
  runButton.classList.toggle("is-loading", isLoading);
  applyButton.classList.toggle("is-loading", isLoading);
  runButton.disabled = isLoading;
  applyButton.disabled = isLoading;
}

function setFileSummary() {
  standardName.textContent = standardFile.files.length ? standardFile.files[0].name : "파일 선택";
  if (!folderFiles.files.length) {
    folderName.textContent = "기본 폴더: data\\테스트";
    return;
  }

  const firstPath = folderFiles.files[0].webkitRelativePath || folderFiles.files[0].name;
  const rootName = firstPath.includes("/") ? firstPath.split("/")[0] : firstPath;
  folderName.textContent = `${rootName} · ${folderFiles.files.length}개`;
}

function buildFormData() {
  const body = new FormData();
  body.append("standard_file", standardFile.files[0], standardFile.files[0].name);

  for (const file of folderFiles.files) {
    body.append("folder_files", file, file.webkitRelativePath || file.name);
  }

  for (const element of form.elements) {
    if (!element.name || element.type === "file") continue;
    if (element.type === "checkbox") {
      if (element.checked) body.append(element.name, "true");
    } else if (element.type === "radio") {
      if (element.checked) body.append(element.name, element.value);
    } else {
      body.append(element.name, element.value);
    }
  }
  body.append("excluded_candidate_paths", [...excludedCandidatePaths].join("\n"));

  return body;
}

function setSummary(data) {
  outputCount.textContent = data.output_count ?? "-";
  fileCount.textContent = data.scanned_files ?? "-";
  matchCount.textContent = data.matched_output_count ?? 0;
  matchedFileCount.textContent = data.matched_file_count ?? 0;
  if (data.match_mode) {
    setRuntimeMode({
      ...(currentRuntimeMode || {}),
      mode: data.match_mode,
      label: modeLabel(data.match_mode),
    });
  }
  resultMeta.textContent = `${data.standard_project_title || "-"} · ${modeLabel(data.match_mode || currentRuntimeMode?.mode)} · 반영 대상 파일 ${data.matched_file_count ?? 0}개`;
}

function renderMatches(matches) {
  const filter = resultFilter.value.trim().toLowerCase();
  const filtered = matches.filter((match) => {
    if (!filter) return true;
    const haystack = [
      match.output_id,
      match.output_name,
      ...match.candidates.flatMap((candidate) => [
        candidate.path,
        candidate.reason,
        candidate.identity?.project_title,
        candidate.identity?.document_title,
      ]),
    ].join(" ").toLowerCase();
    return haystack.includes(filter);
  });

  resultRows.replaceChildren();
  visibleCount.textContent = filtered.length;
  emptyState.hidden = filtered.length > 0;
  emptyState.textContent = "표시할 결과가 없습니다.";
  results.classList.toggle("is-empty", filtered.length === 0);

  for (const match of filtered) {
    const best = getBestCandidate(match);
    const key = outputKey(match);
    const expanded = expandedOutputs.has(key);
    const mainRow = renderOutputRow(match, best, expanded);
    const detailRow = renderDetailRow(match, expanded);

    mainRow.addEventListener("click", () => toggleOutput(key));
    mainRow.querySelector("button").addEventListener("click", (event) => {
      event.stopPropagation();
      toggleOutput(key);
    });

    resultRows.append(mainRow, detailRow);
  }
}

function clearApplyReport() {
  applyReport.hidden = true;
  applyReport.replaceChildren();
}

function renderApplyReport(data) {
  const items = data.apply_items || [];
  const failed = items.filter((item) => item.status === "error");
  const updated = items.filter((item) => item.status === "updated");
  const skipped = Number(data.skipped_file_count || 0);
  const downloadUrl = data.download_url || "";
  const downloadName = data.download_name || "결과.zip";

  applyReport.hidden = false;
  applyReport.className = `apply-report ${failed.length ? "has-errors" : "is-clean"}`;
  applyReport.innerHTML = `
    <div class="apply-report-head">
      <div>
        <strong>반영 결과</strong>
        <span>성공 ${updated.length}건 · 오류 ${failed.length}건 · 건너뜀 ${skipped}건</span>
      </div>
      <span class="apply-report-badge">${failed.length ? "확인 필요" : "오류 없음"}</span>
    </div>
    ${
      downloadUrl
        ? `<a class="apply-download-button" href="${escapeHtml(downloadUrl)}" download="${escapeHtml(downloadName)}">↓ 결과 ZIP 다운로드 <small>${escapeHtml(downloadName)}</small></a>`
        : ""
    }
    ${
      failed.length
        ? `<div class="apply-error-list">
            ${failed.map((item, index) => renderApplyError(item, index)).join("")}
          </div>`
        : `<p class="apply-report-empty">반영 중 오류가 난 파일이 없습니다.</p>`
    }
  `;
}

function renderApplyError(item, index) {
  return `
    <article class="apply-error-card">
      <div class="apply-error-top">
        <span>${index + 1}</span>
        <strong>${escapeHtml(item.output_name || "-")}</strong>
        <small>${escapeHtml(item.output_id || "-")}</small>
      </div>
      <dl>
        <dt>오류 위치</dt>
        <dd>${escapeHtml(item.old_path || "-")}</dd>
        <dt>오류 내용</dt>
        <dd>${escapeHtml(item.error || "알 수 없는 오류")}</dd>
        <dt>백업 위치</dt>
        <dd>${escapeHtml(item.backup_path || "백업 없음")}</dd>
      </dl>
    </article>
  `;
}

function getBestCandidate(match) {
  const candidates = match.candidates || [];
  if (!candidates.length) return null;
  return candidates.reduce((best, current) => {
    const bestScore = Number(best.score || 0);
    const currentScore = Number(current.score || 0);
    return currentScore > bestScore ? current : best;
  }, candidates[0]);
}

function outputKey(match) {
  return `${match.output_id || ""}|${match.output_name || ""}`;
}

function formatScore(value) {
  return Number(value || 0).toFixed(2);
}

function formatAiConfidence(candidate) {
  if (!candidate || candidate.ai_confidence === null || candidate.ai_confidence === undefined) return "";
  const value = Number(candidate.ai_confidence);
  return Number.isFinite(value) ? value.toFixed(2) : "";
}

function renderScoreBadges(candidate) {
  if (!candidate) return `<span class="score">-</span>`;
  const score = formatScore(candidate.score);
  const aiConfidence = formatAiConfidence(candidate);
  if (aiConfidence) {
    return `<span class="score-pair"><span class="ai-confidence">AI ${aiConfidence}</span></span>`;
  }
  return `
    <span class="score-pair">
      <span class="score">점수 ${score}</span>
    </span>
  `;
}

function toggleOutput(key) {
  if (expandedOutputs.has(key)) {
    expandedOutputs.delete(key);
  } else {
    expandedOutputs.add(key);
  }
  renderMatches(lastMatches);
}

function renderOutputRow(match, best, expanded) {
  const row = document.createElement("tr");
  const candidateCount = match.candidates?.length || 0;
  const activeCount = (match.candidates || []).filter((candidate) => !excludedCandidatePaths.has(candidate.path || "")).length;
  const excludedCount = candidateCount - activeCount;
  let fileTitle = "반영 대상 파일 없음";
  let fileSubtext = "반영 대상 없음";
  if (candidateCount === 1) {
    fileTitle = best?.path || "반영 대상 파일 1개";
    fileSubtext = excludedCount ? "1개 파일 제외됨" : "1개 파일 반영";
  } else if (candidateCount > 1) {
    fileTitle = excludedCount ? `${activeCount}개 반영 · ${excludedCount}개 제외` : `${candidateCount}개 파일 모두 반영 대상`;
    fileSubtext = `대표 표시: ${best?.path || "-"}`;
  }

  row.className = `output-row ${expanded ? "expanded" : ""} ${best ? "" : "missing"}`;
  row.innerHTML = `
    <td>
      <span class="output-cell">
        <button class="expand-button" type="button" aria-label="상세 보기">${expanded ? "−" : "+"}</button>
        <span class="output-title">
          <strong>${escapeHtml(match.output_name || "-")}</strong>
          <small>${escapeHtml(match.output_id || "-")}</small>
        </span>
      </span>
    </td>
    <td>
      <span class="file-cell">
        <strong>${escapeHtml(fileTitle)}</strong>
        <small>${escapeHtml(fileSubtext)}</small>
      </span>
    </td>
    <td>${renderScoreBadges(best)}</td>
  `;
  return row;
}

function renderDetailRow(match, expanded) {
  const row = document.createElement("tr");
  row.className = "detail-row";
  row.hidden = !expanded;

  const candidates = match.candidates || [];
  const activeCount = candidates.filter((candidate) => !excludedCandidatePaths.has(candidate.path || "")).length;
  const excludedCount = candidates.length - activeCount;
  row.innerHTML = `
    <td colspan="3">
      <div class="detail-panel">
        <div class="detail-head">
          <strong>${escapeHtml(match.output_name || "-")}</strong>
          <span>${candidates.length ? `${activeCount}개 반영 · ${excludedCount}개 제외` : "반영 대상 없음"}</span>
        </div>
        ${
          candidates.length
            ? candidates.map((candidate, index) => renderCandidateDetail(candidate, index)).join("")
            : `<p class="empty-inline">이 참고 산출물과 연결된 파일은 이번 검사에서 선택되지 않았습니다.</p>`
        }
      </div>
    </td>
  `;
  return row;
}

function renderCandidateDetail(candidate, index) {
  const identity = candidate.identity || {};
  const candidatePath = candidate.path || "";
  const excluded = excludedCandidatePaths.has(candidatePath);
  const documentTitle = identity.document_title || "-";
  const reason = candidate.reason || "-";

  return `
    <article class="candidate-row ${excluded ? "is-excluded" : ""}">
      <label class="row-exclude" title="반영 제외">
        <input type="checkbox" data-exclude-path="${escapeHtml(candidatePath)}" ${excluded ? "checked" : ""}>
      </label>
      <span class="candidate-index">${index + 1}</span>
      <div class="candidate-main">
        <strong>${escapeHtml(candidate.path || "-")}</strong>
        <small>
          <span>${escapeHtml(documentTitle)}</span>
          <span>${escapeHtml(reason)}</span>
        </small>
      </div>
      <div class="candidate-meta">
        ${renderScoreBadges(candidate)}
      </div>
    </article>
  `;
}

function clearResults(message) {
  resultRows.replaceChildren();
  visibleCount.textContent = "0";
  emptyState.hidden = false;
  emptyState.textContent = message || "표시할 결과가 없습니다.";
}

standardFile.addEventListener("change", setFileSummary);
folderFiles.addEventListener("change", setFileSummary);
resultFilter.addEventListener("input", () => renderMatches(lastMatches));
resultRows.addEventListener("change", (event) => {
  const input = event.target;
  if (!(input instanceof HTMLInputElement) || !input.dataset.excludePath) return;
  if (input.checked) {
    excludedCandidatePaths.add(input.dataset.excludePath);
  } else {
    excludedCandidatePaths.delete(input.dataset.excludePath);
  }
  renderMatches(lastMatches);
});

function validateRequiredFiles() {
  if (!standardFile.files.length) {
    setBadge("확인", "error");
    resultMeta.textContent = "문서관리표준 PDF를 선택하세요.";
    return false;
  }

  return true;
}

async function runRequest({ endpoint, busyText, preparingText, doneText, applyMode = false }) {
  if (!validateRequiredFiles()) return;

  setBadge("처리중", "busy");
  setLoading(true);
  resultMeta.textContent = busyText;
  expandedOutputs.clear();
  clearApplyReport();
  clearResults(preparingText);

  try {
    const response = await fetch(endpoint, {
      method: "POST",
      body: buildFormData(),
    });
    const data = await response.json().catch(() => ({ error: "서버 응답을 읽지 못했습니다." }));

    if (!response.ok) {
      throw new Error(data.error || "처리 실패");
    }

    setBadge(doneText, "done");
    if (!applyMode) {
      excludedCandidatePaths.clear();
    }
    lastMatches = data.matches || [];
    setSummary(data);
    if (applyMode) {
      resultMeta.textContent = [
        `반영 ${data.updated_file_count ?? 0}건`,
        `오류 ${data.failed_file_count ?? 0}건`,
        `반영 대상 파일 ${data.matched_file_count ?? 0}개`,
        data.download_url ? "다운로드 준비됨" : "다운로드 없음",
      ].join(" · ");
      renderApplyReport(data);
    }
    renderMatches(lastMatches);
  } catch (error) {
    setBadge("오류", "error");
    resultMeta.textContent = error.message;
    lastMatches = [];
    clearApplyReport();
    clearResults(`처리 중 오류: ${error.message}`);
  } finally {
    setLoading(false);
  }
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  await runRequest({
    endpoint: "/api/check",
    busyText: "파일을 읽는 중",
    preparingText: "검사 결과를 준비하고 있습니다.",
    doneText: "완료",
  });
});

applyButton.addEventListener("click", async () => {
  await runRequest({
    endpoint: "/api/folder-apply",
    busyText: "덤프 폴더를 만들고 문서를 반영하는 중",
    preparingText: "복사본을 만든 뒤 O 대상 문서를 수정하고 있습니다.",
    doneText: "반영완료",
    applyMode: true,
  });
});

loadRuntimeMode();
