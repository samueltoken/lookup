const {
  parseJsonValue,
  findTableRecord,
  resolveTableSize
} = require("./hwp-json-compat");

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function summarizeHwpJsonLayout(jsonValue) {
  const parsed = parseJsonValue(jsonValue);
  const sections = Array.isArray(parsed?.body_text?.sections) ? parsed.body_text.sections : [];
  let paragraphCount = 0;
  let tableCount = 0;
  let explicitPageBreakCount = 0;
  const tableShapes = [];

  for (const section of sections) {
    const paragraphs = Array.isArray(section?.paragraphs) ? section.paragraphs : [];
    paragraphCount += paragraphs.length;
    for (const paragraph of paragraphs) {
      const divideTypes = Array.isArray(paragraph?.para_header?.column_divide_type)
        ? paragraph.para_header.column_divide_type
        : [];
      if (divideTypes.includes("page")) {
        explicitPageBreakCount += 1;
      }
      const table = findTableRecord(paragraph);
      if (table) {
        const size = resolveTableSize(table);
        tableCount += 1;
        tableShapes.push({
          rows: size.rows,
          cols: size.cols,
          pageBreak: normalizeText(table?.attributes?.attribute?.page_break)
        });
      }
    }
  }

  return {
    sectionCount: sections.length,
    paragraphCount,
    tableCount,
    explicitPageBreakCount,
    tableShapes
  };
}

module.exports = {
  summarizeHwpJsonLayout
};
