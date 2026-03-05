// Shared provider routing helpers for SharePoint sequences (Rhino compatible).

function sppr_isBlank(value) {
  return value === null || value === undefined || String(value).trim().length === 0;
}

function sppr_safeString(value) {
  return value === null || value === undefined ? "" : String(value);
}

function sppr_defaultString(value, defaultValue) {
  return sppr_isBlank(value) ? defaultValue : String(value);
}

function sppr_normalizeProvider(value, defaultValue) {
  var normalized = sppr_defaultString(value, defaultValue).trim().toLowerCase();
  if (normalized === "") {
    normalized = sppr_defaultString(defaultValue, "graph").trim().toLowerCase();
  }
  if (normalized !== "graph" && normalized !== "onprem") {
    throw new java.lang.IllegalArgumentException("Unsupported provider: " + normalized + " (expected graph or onprem)");
  }
  return normalized;
}

function sppr_ensureOnPremDispatch() {
  if (typeof sppr_onPremResolveSite !== "function") {
    include("js/sharepoint_onprem_dispatch.js");
  }
}
