const { buildHwpPageModelFromJson } = require("./page-model");
const { escapeHtml } = require("./hwp-json-compat");

function buildPageModelCss(fontFaceCss = "") {
  return `
    ${fontFaceCss || ""}
    @page { size: A4; margin: 0; }
    html, body {
      margin: 0;
      padding: 0;
      background: #fff;
      color: #050b16;
      font-family: "HCR Batang", "Haansoft Batang", Batang, "BatangChe", "Malgun Gothic", "LookupNotoSansKR", serif;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }
    .lookup-hwp-page {
      width: 210mm;
      height: 297mm;
      box-sizing: border-box;
      padding: 18mm 18mm 14mm 18mm;
      overflow: hidden;
      page-break-after: always;
      break-after: page;
      background: #fff;
      position: relative;
    }
    .lookup-hwp-page:last-child {
      page-break-after: auto;
      break-after: auto;
    }
    .lookup-hwp-page-content {
      --lookup-page-scale: 1;
      transform: scale(var(--lookup-page-scale));
      transform-origin: top left;
      width: calc(100% / var(--lookup-page-scale));
    }
    .lookup-hwp-paragraph {
      margin: 0 0 8px;
      line-height: 1.55;
      word-break: keep-all;
      overflow-wrap: anywhere;
    }
    .lookup-section-title-rendered {
      display: flex;
      align-items: center;
      gap: 10px;
      margin: 12px 0 14px;
      border-top: 3px solid #111;
      border-bottom: 3px solid #111;
      padding: 7px 0;
      font-family: "Malgun Gothic", "HCR Dotum", "LookupNotoSansKR", sans-serif;
      font-weight: 800;
    }
    .lookup-section-title-mark {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 40px;
      height: 40px;
      background: #1e4268;
      border: 2px solid #0b1726;
      color: #fff;
      font-size: 25px;
      line-height: 1;
    }
    .lookup-section-title-text {
      font-size: 28px;
      letter-spacing: -0.03em;
      color: #000;
    }
    .lookup-list {
      margin: 0 0 8px;
      padding: 0;
      list-style: none;
    }
    .lookup-list-item {
      margin: 0 0 4px;
      line-height: 1.52;
      padding-left: 1.9em;
      text-indent: -1.9em;
      word-break: keep-all;
      overflow-wrap: anywhere;
    }
    .lookup-list-marker {
      display: inline-block;
      min-width: 1.45em;
      text-indent: 0;
      font-weight: 700;
    }
    table.lookup-table-trace {
      border-collapse: collapse;
      table-layout: fixed;
      margin: 8px 0 10px;
      color: #050b16;
      page-break-inside: avoid;
      break-inside: avoid;
    }
    table.lookup-section-title-table {
      margin-top: 10px;
      margin-bottom: 14px;
    }
    .lookup-table-structured th,
    .lookup-table-structured td {
      box-sizing: border-box;
      padding: 4px 5px;
      word-break: keep-all;
      overflow-wrap: anywhere;
      line-height: 1.35;
      font-size: 13px;
    }
    .lookup-section-title-table th,
    .lookup-section-title-table td {
      font-size: 24px;
      font-weight: 700;
      line-height: 1.15;
    }
    .lookup-split-table {
      margin-top: 4px;
    }
    .lookup-split-table th,
    .lookup-split-table td {
      font-size: 16px;
      line-height: 1.45;
    }
    .lookup-table-structured th {
      font-weight: 700;
    }
    strong {
      font-weight: 700;
    }
  `;
}

function buildHwpPageModelCompatibilityHtmlFromJson(jsonValue, title, options = {}) {
  const model = buildHwpPageModelFromJson(jsonValue, {
    targetPageCount: options.targetPageCount
  });
  if (!model.available || !Array.isArray(model.pages) || !model.pages.length) {
    return {
      html: "",
      diagnostics: model.diagnostics || { reason: "page_model_unavailable" }
    };
  }

  const pageHtml = model.pages
    .map((page) => {
      const scale = Math.max(0.1, Number(page.scale || 1));
      return `<section class="lookup-hwp-page" data-page="${page.index}">
        <div class="lookup-hwp-page-content" style="--lookup-page-scale:${scale.toFixed(4)}">
          ${page.blocks.map((block) => block.html).join("\n")}
        </div>
      </section>`;
    })
    .join("\n");

  return {
    html: `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <title>${escapeHtml(title || "lookup HWP")}</title>
  <style>${buildPageModelCss(options.fontFaceCss || "")}</style>
</head>
<body>
${pageHtml}
</body>
</html>`,
    diagnostics: model.diagnostics
  };
}

module.exports = {
  buildHwpPageModelCompatibilityHtmlFromJson
};
