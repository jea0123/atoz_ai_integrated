// 폴더 산출물 검사 화면입니다. 업로드, API 호출, 매칭 결과 표시를 담당합니다.
const form = document.querySelector("#checkForm");
const standardFile = document.querySelector("#standardFile");
const folderFiles = document.querySelector("#folderFiles");
const requirementFiles = document.querySelector("#requirementFiles");
const proposalFiles = document.querySelector("#proposalFiles");
const standardName = document.querySelector("#standardName");
const folderName = document.querySelector("#folderName");
const requirementName = document.querySelector("#requirementName");
const proposalName = document.querySelector("#proposalName");
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
const aiFallbackToggle = document.querySelector("#aiFallbackToggle");
const initialRevisionYear = document.querySelector("#initialRevisionYear");
const initialRevisionAuthor = document.querySelector("#initialRevisionAuthor");
const initialRevisionApprovalAuthor = document.querySelector("#initialRevisionApprovalAuthor");
const LAST_DUMP_ROOT_KEY = "atoz:lastDumpRoot";
const emptyFolderLabel = form?.dataset.emptyFolderLabel || "기본 폴더: data\\테스트";

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
    rule_ai_fallback: "규칙 매칭 · 미매칭 AI 확인",
    ai_first: "AI 우선",
    rule_only: "규칙 매칭만",
    rule_fallback_no_ollama: "규칙 매칭만",
  };
  return labels[mode] || mode || "-";
}

function modeClass(mode) {
  if (mode === "ai_first") return "ai";
  if (mode === "rule_only" || mode === "rule_fallback_no_ollama") return "no-ai";
  return "rule";
}

function setRuntimeMode(data) {
  currentRuntimeMode = data || currentRuntimeMode;
  if (!runtimeMode || !currentRuntimeMode) return;

  const mode = currentRuntimeMode.mode || "unknown";
  runtimeMode.className = `mode-note ${modeClass(mode)}`;
  runtimeMode.textContent = currentRuntimeMode.label || modeLabel(mode);
}

function selectedMatchMode() {
  return aiFallbackToggle?.checked ? "rule_ai_fallback" : "rule_only";
}

function syncMatchModeDisplay() {
  setRuntimeMode({
    ...(currentRuntimeMode || {}),
    mode: selectedMatchMode(),
    label: modeLabel(selectedMatchMode()),
  });
}

function selectedFolderFiles() {
  return UploadFilters.selectedFiles(folderFiles?.files, UploadFilters.CHECK_EXTENSIONS);
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
  if (requirementFiles) {
    if (!requirementFiles.files.length) {
      requirementName.textContent = "파일명 기준 SFR별 복제";
    } else {
      requirementName.textContent = requirementFiles.files.length > 1
        ? `${requirementFiles.files[0].name} 외 ${requirementFiles.files.length - 1}개`
        : requirementFiles.files[0].name;
    }
  }
  if (proposalFiles && proposalName) {
    if (!proposalFiles.files.length) {
      proposalName.textContent = "요구사항목록표 SFR ID 추출";
    } else {
      proposalName.textContent = proposalFiles.files.length > 1
        ? `${proposalFiles.files[0].name} 외 ${proposalFiles.files.length - 1}개`
        : proposalFiles.files[0].name;
    }
  }
  if (!folderFiles.files.length) {
    folderName.textContent = emptyFolderLabel;
    return;
  }

  const firstPath = folderFiles.files[0].webkitRelativePath || folderFiles.files[0].name;
  const rootName = firstPath.includes("/") ? firstPath.split("/")[0] : firstPath;
  folderName.textContent = `${rootName} · ${folderFiles.files.length}개`;
}

function buildFormData({ includeRequirementFiles = true } = {}) {
  const body = new FormData();
  body.append("standard_file", standardFile.files[0], standardFile.files[0].name);

  for (const file of selectedFolderFiles()) {
    body.append("folder_files", file, UploadFilters.relativePath(file));
  }
  if (includeRequirementFiles && requirementFiles) {
    for (const file of requirementFiles.files) {
      body.append("requirement_files", file, file.webkitRelativePath || file.name);
    }
  }
  if (includeRequirementFiles && proposalFiles) {
    for (const file of proposalFiles.files) {
      body.append("proposal_files", file, file.webkitRelativePath || file.name);
    }
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
  body.append("match_mode", selectedMatchMode());
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
  resultMeta.textContent = [
    "검사 성공",
    "오류 0건",
    data.standard_project_title || "-",
    modeLabel(data.match_mode || currentRuntimeMode?.mode),
    `반영 대상 파일 ${data.matched_file_count ?? 0}개`,
  ].join(" · ");
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
  const filenameUnchanged = Number(data.filename_unchanged_count || 0);
  const requirementEnabled = Boolean(data.requirement_generation_enabled);
  const requirementCreated = Number(data.requirement_generated_file_count || 0);
  const requirementFolders = Number(data.requirement_generated_folder_count || 0);
  const requirementRemoved = Number(data.requirement_generation_removed_file_count || 0);
  const requirementSkipped = Number(data.requirement_generation_skipped_count || 0);
  const requirementWarnings = Number(data.requirement_generation_warning_count || 0);
  const requirementErrors = Number(data.requirement_generation_error_count || 0);
  const requirementReadme = data.requirement_generation_readme_path || "";
  const requirementGenerationLabel = data.artifact_category === "management"
    ? "제안요청서 요구사항별 자동 생성"
    : "요구사항 파일명별 자동 생성";
  const initialRevisionUpdated = Number(data.initial_revision_updated_count || 0);
  const initialRevisionSkipped = Number(data.initial_revision_skipped_count || 0);
  const initialRevisionFailed = Number(data.initial_revision_failed_count || 0);
  const revisionDate = data.initial_revision_date || `${initialRevisionYear?.value || new Date().getFullYear()}-00-00`;
  const revisionAuthor = data.initial_revision_author || initialRevisionAuthor?.value || "송아름";
  const revisionApprovalAuthor = data.initial_revision_approval_author || initialRevisionApprovalAuthor?.value || "임채현";
  const hasErrors = failed.length || requirementErrors || initialRevisionFailed;
  const dumpRoot = data.dump_root || "";
  if (dumpRoot) {
    try {
      localStorage.setItem(LAST_DUMP_ROOT_KEY, dumpRoot);
    } catch {
      // localStorage를 쓸 수 없는 환경이면 화면 표시만 유지한다.
    }
  }

  applyReport.hidden = false;
  applyReport.className = `apply-report ${hasErrors ? "has-errors" : "is-clean"}`;
  applyReport.innerHTML = `
    <div class="apply-report-head">
      <div>
        <strong>반영 결과</strong>
        <span>성공 ${updated.length}건 · 오류 ${failed.length}건 · 건너뜀 ${skipped}건 · 파일명 미변경 ${filenameUnchanged}건</span>
      </div>
      <span class="apply-report-badge">${hasErrors ? "확인 필요" : "오류 없음"}</span>
    </div>
    ${
      requirementEnabled
        ? `<div class="apply-folder-result">
            <span>${requirementGenerationLabel}</span>
            <code>ID 폴더 ${requirementFolders}건 · 생성 ${requirementCreated}건 · 기존 삭제 ${requirementRemoved}건 · ID 없음 ${requirementSkipped}건 · 경고 ${requirementWarnings}건 · 오류 ${requirementErrors}건</code>
            ${requirementReadme ? `<code>${escapeHtml(requirementReadme)}</code>` : ""}
          </div>`
        : ""
    }
    <div class="apply-folder-result">
      <span>개정이력 v0.1 초기화</span>
      <code>날짜 ${escapeHtml(revisionDate)} · 작성자 ${escapeHtml(revisionAuthor)} · 승인자 ${escapeHtml(revisionApprovalAuthor)} · 성공 ${initialRevisionUpdated}건 · 스킵 ${initialRevisionSkipped}건 · 오류 ${initialRevisionFailed}건</code>
    </div>
    ${
      dumpRoot
        ? `<div class="apply-folder-result">
            <span>결과 폴더</span>
            <code>${escapeHtml(dumpRoot)}</code>
            <a href="/metadata.html?dump_root=${encodeURIComponent(dumpRoot)}">메타데이터 반영으로 이동</a>
          </div>`
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
requirementFiles?.addEventListener("change", setFileSummary);
proposalFiles?.addEventListener("change", setFileSummary);
aiFallbackToggle?.addEventListener("change", syncMatchModeDisplay);
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

function validateInitialRevisionInputs() {
  const year = initialRevisionYear?.value.trim() || "";
  const author = initialRevisionAuthor?.value.trim() || "";
  const approvalAuthor = initialRevisionApprovalAuthor?.value.trim() || "";
  if (!/^\d{4}$/.test(year)) {
    resultMeta.textContent = "개정연도는 4자리 숫자로 입력하세요.";
    initialRevisionYear?.focus();
    return false;
  }
  if (!author) {
    resultMeta.textContent = "개정이력 작성자를 입력하세요.";
    initialRevisionAuthor?.focus();
    return false;
  }
  if (!approvalAuthor) {
    resultMeta.textContent = "개정이력 승인자를 입력하세요.";
    initialRevisionApprovalAuthor?.focus();
    return false;
  }
  return true;
}

function validateRequiredFiles({ applyMode = false } = {}) {
  if (!standardFile.files.length) {
    setBadge("확인", "error");
    resultMeta.textContent = "문서관리표준 PDF를 선택하세요.";
    return false;
  }
  const requireFolder = Boolean(form.querySelector('[name="require_folder"][value="true"]'));
  if (requireFolder && !folderFiles.files.length) {
    setBadge("확인", "error");
    resultMeta.textContent = "검사할 폴더를 선택하세요.";
    return false;
  }
  if (requireFolder && !selectedFolderFiles().length) {
    setBadge("확인", "error");
    resultMeta.textContent = "지원하는 산출물 문서가 없습니다. bak/backup/old와 비문서 파일은 제외됩니다.";
    return false;
  }
  if (applyMode && !validateInitialRevisionInputs()) {
    setBadge("확인", "error");
    return false;
  }

  return true;
}

async function runRequest({ endpoint, busyText, preparingText, doneText, applyMode = false }) {
  if (!validateRequiredFiles({ applyMode })) return;

  setBadge("처리중", "busy");
  setLoading(true);
  resultMeta.textContent = busyText;
  expandedOutputs.clear();
  clearApplyReport();
  clearResults(preparingText);

  try {
    const response = await fetch(endpoint, {
      method: "POST",
      body: buildFormData({ includeRequirementFiles: applyMode }),
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
        data.standard_project_title || "-",
        modeLabel(data.match_mode || currentRuntimeMode?.mode),
        `매칭 파일 ${data.matched_file_count ?? 0}개 표시 중`,
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
