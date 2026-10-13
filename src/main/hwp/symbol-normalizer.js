const DEFAULT_SYMBOL_MAP = Object.freeze({
  "\uF03DA": "\u25A1",
  "\uF0A7": "\u25CB",
  "\uF0B7": "\u00B7"
});

function normalizeSpecialSymbols(text, options = {}) {
  const source = String(text || "");
  if (!source) {
    return "";
  }
  const customMap = options && typeof options === "object" ? options.symbolMap : null;
  const lookup = customMap && typeof customMap === "object" ? customMap : DEFAULT_SYMBOL_MAP;
  let normalized = source;
  for (const [from, to] of Object.entries(lookup)) {
    if (!from) {
      continue;
    }
    normalized = normalized.split(from).join(String(to || ""));
  }
  return normalized;
}

function normalizeListMarker(marker) {
  const normalized = normalizeSpecialSymbols(marker);
  return String(normalized || "").trim();
}

module.exports = {
  DEFAULT_SYMBOL_MAP,
  normalizeSpecialSymbols,
  normalizeListMarker
};
