const header = document.querySelector("[data-app-header]");

if (header) {
  const eyebrow = header.dataset.eyebrow || "Document Tools";
  const title = header.dataset.title || "Document Tools";
  const statusId = header.dataset.statusId || "statusBadge";
  const active = header.dataset.active || "check";

  header.innerHTML = [
    `<div class="brand-block">`,
    `  <span class="brand-mark">DT</span>`,
    `  <span class="brand-copy">`,
    `    <span class="eyebrow">${escapeHeaderHtml(eyebrow)}</span>`,
    `    <strong>${escapeHeaderHtml(title)}</strong>`,
    `  </span>`,
    `</div>`,
    `<div class="topbar-actions">`,
    `  <nav class="app-nav-tabs" aria-label="${text("screenSelect")}" style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;">`,
    `    <a href="/management.html" style="${tabStyle(active === "management")}">${text("managementTab")}</a>`,
    `    <a href="/check.html" style="${tabStyle(active === "check")}">${text("developmentTab")}</a>`,
    `    <a href="/management-template.html" style="${tabStyle(active === "managementTemplate")}">${text("managementTemplateTab")}</a>`,
    `    <a href="/development-template.html" style="${tabStyle(active === "developmentTemplate")}">${text("developmentTemplateTab")}</a>`,
//    `    <a href="/metadata.html" style="${tabStyle(active === "metadata")}">${text("metadataTab")}</a>`,
    `     <a href="/qa.html" style="${tabStyle(active === "qa")}">QA ${text("generation")}</a>`,
    `  </nav>`,
    `  <div class="status-pill" id="${escapeHeaderHtml(statusId)}">${text("standby")}</div>`,
    `</div>`,
  ].join("");
}

function text(key) {
  const labels = {
    screenSelect: "화면 선택",
    developmentTab: "개발산출물",
    managementTab: "관리산출물",
    developmentTemplateTab: "개발템플릿",
    managementTemplateTab: "관리템플릿",
    metadataTab: "메타데이터",
    generation: "생성",
    standby: "대기",
  };
  return labels[key] || "";
}

function tabStyle(isActive) {
  const base = [
    "display:inline-flex",
    "align-items:center",
    "justify-content:center",
    "min-height:34px",
    "padding:0 11px",
    "border-radius:8px",
    "font-size:12px",
    "font-weight:900",
    "text-decoration:none",
    "white-space:nowrap",
  ];
  if (isActive) {
    return [
      ...base,
      "border:1px solid #1f5fbf",
      "background:#edf4ff",
      "color:#173f82",
    ].join(";");
  }
  return [
    ...base,
    "border:1px solid #d7dee8",
    "background:#ffffff",
    "color:#64748b",
  ].join(";");
}

function escapeHeaderHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
