const form = document.querySelector("#metadataForm");
const wbsFile = document.querySelector("#wbsFile");
const standardFile = document.querySelector("#standardFile");
const documentFiles = document.querySelector("#documentFiles");
const dumpRootInput = document.querySelector("#dumpRootInput");
const wbsName = document.querySelector("#wbsName");
const standardName = document.querySelector("#standardName");
const folderName = document.querySelector("#folderName");
const statusBadge = document.querySelector("#statusBadge");
const previewButton = document.querySelector("#previewButton");
const applyButton = document.querySelector("#applyButton");
const wbsCount = document.querySelector("#wbsCount");
const docCount = document.querySelector("#docCount");
const matchedCount = document.querySelector("#matchedCount");
const reviewCount = document.querySelector("#reviewCount");
const resultMeta = document.querySelector("#resultMeta");
const resultFilter = document.querySelector("#resultFilter");
const resultRows = document.querySelector("#resultRows");
const visibleCount = document.querySelector("#visibleCount");
const emptyState = document.querySelector("#emptyState");
const loadingOverlay = document.querySelector("#loadingOverlay");
const loadingCancelButton = document.querySelector("#loadingCancelButton");
const applyReport = document.querySelector("#applyReport");
const LAST_DUMP_ROOT_KEY = "atoz:lastDumpRoot";

let lastTargets = [];
let lastPreviewApprovalAuthor = "";
let activeRequest = null;
const excludedPaths = new Set();

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

function setLoading(isLoading) {
  loadingOverlay.hidden = !isLoading;
  previewButton.disabled = isLoading;
  applyButton.disabled = isLoading;
  previewButton.classList.toggle("is-loading", isLoading);
  applyButton.classList.toggle("is-loading", isLoading);
  if (loadingCancelButton) loadingCancelButton.disabled = !isLoading;
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

function buildRequestFormData(request) {
  const body = buildFormData();
  if (request?.requestId) body.append("request_id", request.requestId);
  return body;
}

function updateFileLabels() {
  wbsName.textContent = wbsFile.files.length ? wbsFile.files[0].name : ".xlsx, .xlsm";
  standardName.textContent = standardFile.files.length ? standardFile.files[0].name : ".pdf, .hwp, .hwpx";
  if (!documentFiles.files.length) {
    folderName.textContent = "직접 업로드할 때 선택";
    return;
  }
  const firstPath = documentFiles.files[0].webkitRelativePath || documentFiles.files[0].name;
  const rootName = firstPath.includes("/") ? firstPath.split("/")[0] : firstPath;
  folderName.textContent = `${rootName} · ${documentFiles.files.length}개`;
}

function buildFormData() {
  const body = new FormData();
  if (wbsFile.files.length) {
    body.append("wbs_file", wbsFile.files[0], wbsFile.files[0].name);
  }
  if (standardFile.files.length) {
    body.append("standard_file", standardFile.files[0], standardFile.files[0].name);
  }
  for (const file of documentFiles.files) {
    body.append("document_files", file, file.webkitRelativePath || file.name);
  }
  if (dumpRootInput.value.trim()) {
    body.append("dump_root", dumpRootInput.value.trim());
  }
  body.append("excluded_paths", [...excludedPaths].join("\n"));
  return body;
}

function validateRequiredInputs() {
  if (!wbsFile.files.length) {
    throw new Error("WBS 파일을 선택하세요.");
  }
  if (!standardFile.files.length) {
    throw new Error("문서관리표준 파일을 선택하세요.");
  }
  if (!documentFiles.files.length && !dumpRootInput.value.trim()) {
    throw new Error("산출물 폴더를 업로드하거나 산출물 매핑 결과 폴더 경로를 입력하세요.");
  }
}

function setSummary(data) {
  wbsCount.textContent = data.wbs_record_count ?? "-";
  docCount.textContent = data.document_count ?? "-";
  matchedCount.textContent = data.matched_count ?? "-";
  reviewCount.textContent = Number(data.ambiguous_count || 0) + Number(data.unmatched_count || 0);
  resultMeta.textContent = `${data.folder_root || "-"} · 반영 가능 ${data.matched_count || 0}개 · 승인자 ${data.approval_author || "-"}`;
}

function statusLabel(status) {
  const labels = {
    matched: "반영 가능",
    no_change: "수정 위치 없음",
    ambiguous: "확인 필요",
    unmatched: "매칭 없음",
    error: "오류",
  };
  return labels[status] || status || "-";
}

function targetStatusRank(status) {
  const ranks = {
    unmatched: 0,
    no_change: 1,
    ambiguous: 1,
    matched: 2,
  };
  return ranks[status] ?? 3;
}

function renderTargets() {
  const filter = resultFilter.value.trim().toLowerCase();
  const targets = lastTargets
    .filter((target) => {
      if (!filter) return true;
      const haystack = [
        target.relative_path,
        target.output_name,
        target.author,
        target.revision_date,
        target.current?.author,
        target.current?.revision_date,
        target.message,
      ].join(" ").toLowerCase();
      return haystack.includes(filter);
    });
  targets.sort((left, right) => {
      const statusDiff = targetStatusRank(left.status) - targetStatusRank(right.status);
      if (statusDiff !== 0) return statusDiff;
      return String(left.relative_path || "").localeCompare(String(right.relative_path || ""), "ko");
    });

  resultRows.replaceChildren();
  visibleCount.textContent = targets.length;
  emptyState.hidden = targets.length > 0;

  for (const target of targets) {
    const row = document.createElement("tr");
    const canApply = target.status === "matched";
    const checked = canApply && !excludedPaths.has(target.relative_path);
    row.innerHTML = `
      <td>
        <label>
          <input type="checkbox" ${checked ? "checked" : ""} ${canApply ? "" : "disabled"} data-path="${escapeHtml(target.relative_path)}">
        </label>
      </td>
      <td>
        <span class="meta-file">
          <strong>${escapeHtml(target.relative_path)}</strong>
          <small>${escapeHtml(target.output_name || "-")}</small>
        </span>
      </td>
      <td>
        <span class="meta-values">
          <strong>${escapeHtml(target.current?.author || "-")}</strong>
          <small>${escapeHtml(target.current?.revision_date || "-")}</small>
          <small>이력: ${escapeHtml(target.current?.revision_author || "-")} / ${escapeHtml(target.current?.revision_history_date || "-")}</small>
        </span>
      </td>
      <td>
        <span class="meta-values">
          <strong>${escapeHtml(target.author || "-")}</strong>
          <small>${escapeHtml(target.revision_date || "-")}</small>
          <small>승인자: ${escapeHtml(lastPreviewApprovalAuthor || "-")}</small>
        </span>
      </td>
      <td>
        <span class="meta-status">
          <span class="meta-badge ${escapeHtml(target.status)}">${escapeHtml(statusLabel(target.status))}</span>
          <small>${escapeHtml(target.message || `${target.candidate_count || 0}개 WBS 후보`)}</small>
        </span>
      </td>
    `;
    const checkbox = row.querySelector("input[type='checkbox']");
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) {
        excludedPaths.delete(target.relative_path);
      } else {
        excludedPaths.add(target.relative_path);
      }
    });
    resultRows.append(row);
  }
}

function renderApplyReport(data) {
  const items = data.apply_items || [];
  const failed = items.filter((item) => item.status === "error");
  applyReport.hidden = false;
  applyReport.className = `apply-report metadata-apply-report ${failed.length ? "has-errors" : "is-clean"}`;
  applyReport.innerHTML = `
    <div class="apply-report-head">
      <div>
        <strong>반영 결과</strong>
        <span>성공 ${data.updated_file_count || 0}건 · 오류 ${data.failed_file_count || 0}건 · 제외/스킵 ${data.skipped_file_count || 0}건</span>
      </div>
      <span class="apply-report-badge">${failed.length ? "확인 필요" : "완료"}</span>
    </div>
    ${
      data.dump_root
        ? `<div class="apply-folder-result">
            <span>결과 폴더</span>
            <code>${escapeHtml(data.dump_root || "")}</code>
            ${data.download_url ? `<a href="${escapeHtml(data.download_url)}">${escapeHtml(data.download_name || "결과 ZIP 다운로드")}</a>` : ""}
          </div>`
        : ""
    }
    ${
      failed.length
        ? `<div class="apply-error-list">${failed.map(renderApplyError).join("")}</div>`
        : `<p class="apply-report-empty">표지, 머리말, 개정이력 반영이 완료되었습니다.</p>`
    }
  `;
}

function renderApplyError(item) {
  return `
    <article class="apply-error-card">
      <div class="apply-error-top">
        <strong>${escapeHtml(item.relative_path || item.old_path || "-")}</strong>
        <small>${escapeHtml(item.output_name || "-")}</small>
      </div>
      <dl>
        <dt>오류</dt>
        <dd>${escapeHtml(item.error || "-")}</dd>
        <dt>백업</dt>
        <dd>${escapeHtml(item.backup_path || "-")}</dd>
      </dl>
    </article>
  `;
}

async function requestPreview() {
  setLoading(true);
  setBadge("확인 중", "busy");
  applyReport.hidden = true;
  const request = beginCancelableRequest();
  try {
    validateRequiredInputs();
    const response = await fetch("/api/metadata-preview", { method: "POST", body: buildRequestFormData(request), signal: request.controller.signal });
    const data = await response.json();
    if (!response.ok || !data.ok) throw new Error(data.error || "미리보기에 실패했습니다.");
    lastTargets = data.targets || [];
    lastPreviewApprovalAuthor = data.approval_author || "";
    setSummary(data);
    renderTargets();
    setBadge("완료", "done");
  } catch (error) {
    if (isAbortError(error)) {
      markRequestCanceled();
      return;
    }
    setBadge("오류", "error");
    resultMeta.textContent = error.message;
  } finally {
    endCancelableRequest(request);
    setLoading(false);
  }
}

async function requestApply() {
  if (!lastTargets.length) {
    await requestPreview();
    if (!lastTargets.length) return;
  }
  setLoading(true);
  setBadge("반영 중", "busy");
  const request = beginCancelableRequest();
  try {
    validateRequiredInputs();
    const response = await fetch("/api/metadata-apply", { method: "POST", body: buildRequestFormData(request), signal: request.controller.signal });
    const data = await response.json();
    if (!response.ok || !data.ok) {
      renderApplyReport(data);
      throw new Error(data.error || "일부 문서 반영에 실패했습니다.");
    }
    renderApplyReport(data);
    setBadge("완료", "done");
  } catch (error) {
    setBadge("확인 필요", "error");
    resultMeta.textContent = error.message;
  } finally {
    if (request.controller.signal.aborted) markRequestCanceled();
    endCancelableRequest(request);
    setLoading(false);
  }
}

wbsFile.addEventListener("change", updateFileLabels);
standardFile.addEventListener("change", updateFileLabels);
documentFiles.addEventListener("change", updateFileLabels);
resultFilter.addEventListener("input", renderTargets);
form.addEventListener("submit", (event) => {
  event.preventDefault();
  requestPreview();
});
applyButton.addEventListener("click", requestApply);
loadingCancelButton?.addEventListener("click", () => {
  sendCancelRequest(activeRequest);
  activeRequest?.controller.abort();
});

function initializeDumpRootInput() {
  const params = new URLSearchParams(window.location.search);
  const queryDumpRoot = params.get("dump_root") || "";
  let storedDumpRoot = "";
  try {
    storedDumpRoot = localStorage.getItem(LAST_DUMP_ROOT_KEY) || "";
  } catch {
    storedDumpRoot = "";
  }
  dumpRootInput.value = queryDumpRoot || storedDumpRoot;
}

initializeDumpRootInput();
updateFileLabels();
