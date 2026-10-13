const { escapeHtml } = require("./hwp-json-compat");
const { hwpUnitToPx } = require("./style-profile");

function hwpUnitToMm(value, fallback) {
  const amount = Number(value);
  return Number.isFinite(amount) && amount >= 0 ? (amount * 25.4 / 7200).toFixed(3) : String(fallback);
}

function buildHwpxSectionLayout(def = {}, index = 0) {
  const width = Number(def.pageWidth) > 0 ? hwpUnitToMm(def.pageWidth, 210) : "210";
  const height = Number(def.pageHeight) > 0 ? hwpUnitToMm(def.pageHeight, 297) : "297";
  const margins = def.pageMargin || {};
  const name = `lookupSection${Math.max(0, Math.floor(Number(index) || 0))}`;
  const margin = [margins.top, margins.right, margins.bottom, margins.left].map(value => `${hwpUnitToMm(value, 18)}mm`).join(" ");
  const columns = Math.max(1, Math.floor(Number(def.columns?.count) || 1));
  const style = `page:${name};${columns > 1 ? `column-count:${columns};column-gap:${Math.max(0, hwpUnitToPx(def.columns.gap))}px;` : ""}`;
  return { name, style, css: `@page ${name} { size:${width}mm ${height}mm; margin:${margin}; }` };
}

function buildHwpxPictureHtml(picture = {}, binaryItems = []) {
  const binary = binaryItems.find(item => Number(item?.id) === Number(picture.binDataIDRef));
  if (!binary?.data?.length) return "";
  const format = String(binary.format || "").toLowerCase();
  const mime = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", tiff: "image/tiff", svg: "image/svg+xml" }[format];
  if (!mime) return "";
  const width = Math.max(0, hwpUnitToPx(picture.width));
  const height = Math.max(0, hwpUnitToPx(picture.height));
  const dimensions = `${width ? `width:${width.toFixed(3)}px;` : ""}${height ? `height:${height.toFixed(3)}px;` : ""}`;
  const encoded = Buffer.from(binary.data).toString("base64");
  return `<img class="lookup-hwpx-picture" src="data:${mime};base64,${encoded}" alt="${escapeHtml(binary.name || "")}" style="${dimensions}max-width:100%;object-fit:contain;" />`;
}

function normalizeHwpxDocumentStyles(document = {}) {
  const head = document.head || {};
  const fonts = Array.isArray(head.fontFaces) ? head.fontFaces : [];
  const hangulFonts = fonts.filter(face => String(face.lang || "").toUpperCase() === "HANGUL");
  return {
    ...document,
    head: {
      ...head,
      charProperties: (head.charProperties || []).map(property => {
        const color = Number(property.textColor);
        const fontId = Number(property.fontRef?.hangul ?? 0);
        const font = hangulFonts[fontId] || fonts[fontId];
        return {
          ...property,
          fontSize: property.height,
          textColor: Number.isFinite(color) ? {r:color & 255,g:(color >> 8) & 255,b:(color >> 16) & 255} : property.textColor,
          fontFamily: String(font?.fontName || ""),
          letterSpacing: property.spacing,
          underlineType: property.underline && String(property.underline).toUpperCase() !== "NONE" ? 1 : 0,
          strike: property.strikeout && String(property.strikeout).toUpperCase() !== "NONE" ? 1 : 0
        };
      }),
      paraProperties: (head.paraProperties || []).map(property => ({
        ...property,
        align: String(property.alignment || "").toLowerCase(),
        lineSpacing: property.lineSpacing?.type === "PERCENT" ? property.lineSpacing.value : 0,
        leftMargin: property.paraMargin?.left,
        rightMargin: property.paraMargin?.right,
        indent: property.paraMargin?.indent,
        topSpacing: property.paraMargin?.prevSpacing,
        bottomSpacing: property.paraMargin?.nextSpacing
      }))
    }
  };
}

module.exports = { buildHwpxSectionLayout, buildHwpxPictureHtml, normalizeHwpxDocumentStyles };
