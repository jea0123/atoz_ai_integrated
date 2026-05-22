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
    `  <nav class="app-nav-tabs" aria-label="${text("screenSelect")}" style="display:flex;align-items:center;gap:6px;">`,
    `    <a href="/check.html" style="${tabStyle(active === "check")}">${text("mappingTab")}</a>`,
    `    <a href="/metadata.html" style="${tabStyle(active === "metadata")}">${text("metadataTab")}</a>`,
    `  </nav>`,
    `  <div class="status-pill" id="${escapeHeaderHtml(statusId)}">${text("standby")}</div>`,
    `</div>`,
  ].join("");
}

function text(key) {
  const labels = {
    screenSelect: "\uD654\uBA74 \uC120\uD0DD",
    mappingTab: "\uC0B0\uCD9C\uBB3C \uB9E4\uD551",
    metadataTab: "\uBA54\uD0C0\uB370\uC774\uD130",
    generation: "\uC0DD\uC131",
    standby: "\uB300\uAE30",
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
