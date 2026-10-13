function clamp(value, min, max) {
  return Math.min(max, Math.max(min, Number(value) || 0));
}

function hwpUnitToPx(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return 0;
  }
  return (numeric * 96) / 7200;
}

function hwpFontUnitToPx(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric <= 0) {
    return 0;
  }
  const pointSize = numeric / 100;
  return pointSize * (96 / 72);
}

function parseColorComponents(color) {
  if (color === null || color === undefined || color === "") {
    return null;
  }
  if (typeof color === "object") {
    const r = Number(color.r ?? color.red);
    const g = Number(color.g ?? color.green);
    const b = Number(color.b ?? color.blue);
    if (Number.isFinite(r) && Number.isFinite(g) && Number.isFinite(b)) {
      return {
        r: clamp(r, 0, 255),
        g: clamp(g, 0, 255),
        b: clamp(b, 0, 255)
      };
    }
    return (
      parseColorComponents(color.color)
      || parseColorComponents(color.value)
      || parseColorComponents(color.rgb)
      || parseColorComponents(color.argb)
      || null
    );
  }
  if (typeof color === "string") {
    const value = color.trim();
    if (!value) {
      return null;
    }
    const rgbMatch = value.match(/^rgba?\((\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})/i);
    if (rgbMatch) {
      return {
        r: clamp(Number(rgbMatch[1]), 0, 255),
        g: clamp(Number(rgbMatch[2]), 0, 255),
        b: clamp(Number(rgbMatch[3]), 0, 255)
      };
    }
    const hexMatch = value.match(/^#?([0-9a-f]{6})$/i);
    if (hexMatch) {
      const numeric = Number.parseInt(hexMatch[1], 16);
      return {
        r: (numeric >> 16) & 255,
        g: (numeric >> 8) & 255,
        b: numeric & 255
      };
    }
    if (/^\d+$/.test(value)) {
      return parseColorComponents(Number(value));
    }
    return null;
  }
  const numeric = Number(color);
  if (!Number.isFinite(numeric)) {
    return null;
  }
  const intValue = Math.max(0, Math.floor(numeric)) & 0xffffff;
  return {
    r: (intValue >> 16) & 255,
    g: (intValue >> 8) & 255,
    b: intValue & 255
  };
}

function toCssColor(color) {
  const parsed = parseColorComponents(color);
  if (!parsed) {
    return "";
  }
  return `rgb(${Math.round(parsed.r)}, ${Math.round(parsed.g)}, ${Math.round(parsed.b)})`;
}

function isWhiteColor(color) {
  const parsed = parseColorComponents(color);
  if (!parsed) {
    return false;
  }
  return Number(parsed.r) === 255 && Number(parsed.g) === 255 && Number(parsed.b) === 255;
}

function normalizeAlign(rawValue) {
  const value = String(rawValue || "").trim().toLowerCase();
  switch (value) {
    case "center":
      return "center";
    case "right":
      return "right";
    case "justify":
      return "justify";
    case "left":
      return "left";
    default:
      return "";
  }
}

function readNumericFromAny(sources, keys, fallback = 0) {
  for (const source of Array.isArray(sources) ? sources : []) {
    if (!source || typeof source !== "object") {
      continue;
    }
    for (const key of Array.isArray(keys) ? keys : []) {
      const value = Number(source[key]);
      if (Number.isFinite(value)) {
        return value;
      }
    }
  }
  return fallback;
}

function resolveShadowOffsetPx(value, fallback) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return fallback;
  }
  if (Math.abs(numeric) > 30) {
    return clamp(hwpUnitToPx(numeric), -4, 4);
  }
  return clamp(numeric, -4, 4);
}

function normalizeInlineStyle(input = {}) {
  const style = input && typeof input === "object" ? input : {};
  const attributes = style.attributes && typeof style.attributes === "object" ? style.attributes : {};
  const fontWeightRaw = Number(
    attributes.font_weight
    || attributes.fontWeight
    || attributes.weight
    || style.font_weight
    || style.fontWeight
    || style.weight
    || 0
  );
  const boldRaw = attributes.bold ?? style.bold;
  const textColor = style.text_color || style.textColor || style.color || null;
  const shadeColor = style.shading_color || style.shadingColor || style.backgroundColor || style.bgColor || null;
  const fontSizePx = hwpFontUnitToPx(
    style.base_size
    || style.baseSize
    || style.font_size
    || style.fontSize
    || style.sz
    || 0
  );
  const letterSpacingRaw = style.letter_spacing?.korean ?? style.letterSpacing ?? attributes.letter_spacing ?? 0;
  const letterSpacingEm = Number.isFinite(Number(letterSpacingRaw)) ? Number(letterSpacingRaw) / 100 : 0;
  const underlineType = Number(attributes.underline_type || attributes.underlineType || style.underline_type || style.underlineType || 0);
  const underlineStyle = Number(attributes.underline_style || attributes.underlineStyle || style.underline_style || style.underlineStyle || 0);
  const strikeType = Number(attributes.strikethrough || style.strikethrough || style.strike || 0);
  const strikeStyle = Number(attributes.strikethrough_style || attributes.strikethroughStyle || style.strikethrough_style || style.strikethroughStyle || 0);
  const outlineType = readNumericFromAny(
    [attributes, style],
    ["outline_type", "outlineType", "text_outline_type", "textOutlineType"],
    0
  );
  const shadowType = readNumericFromAny(
    [attributes, style],
    ["shadow_type", "shadowType", "text_shadow_type", "textShadowType"],
    0
  );
  const emboss =
    attributes.emboss === true
    || attributes.emboss === 1
    || style.emboss === true
    || style.emboss === 1;
  const engrave =
    attributes.engrave === true
    || attributes.engrave === 1
    || style.engrave === true
    || style.engrave === 1;
  const outlineColor =
    style.outline_color
    || style.outlineColor
    || attributes.outline_color
    || attributes.outlineColor
    || (isWhiteColor(textColor) ? { r: 15, g: 23, b: 42 } : textColor);
  const shadowColor =
    style.shadow_color
    || style.shadowColor
    || attributes.shadow_color
    || attributes.shadowColor
    || { r: 30, g: 41, b: 59 };
  const strokeWidthRaw = readNumericFromAny(
    [attributes, style],
    ["outline_width", "outlineWidth", "text_stroke_width", "textStrokeWidth"],
    0
  );
  const shadowX = readNumericFromAny(
    [attributes, style],
    ["shadow_offset_x", "shadowOffsetX", "shadow_spacing_x", "shadowSpacingX"],
    1
  );
  const shadowY = readNumericFromAny(
    [attributes, style],
    ["shadow_offset_y", "shadowOffsetY", "shadow_spacing_y", "shadowSpacingY"],
    1
  );
  const strokeWidthPx = outlineType > 0
    ? clamp(strokeWidthRaw > 0 ? hwpUnitToPx(strokeWidthRaw) : Math.max(0.28, Math.min(0.7, (fontSizePx || 14) * 0.035)), 0.25, 1.4)
    : 0;
  const textShadow = shadowType > 0
    ? `${resolveShadowOffsetPx(shadowX, 1).toFixed(1)}px ${resolveShadowOffsetPx(shadowY, 1).toFixed(1)}px 0 ${toCssColor(shadowColor) || "rgba(15,23,42,0.45)"}`
    : emboss
      ? "0.7px 0.7px 0 rgba(255,255,255,0.9), -0.7px -0.7px 0 rgba(15,23,42,0.45)"
      : engrave
        ? "-0.7px -0.7px 0 rgba(255,255,255,0.9), 0.7px 0.7px 0 rgba(15,23,42,0.45)"
        : "";
  const normalized = {
    bold:
      boldRaw === true
      || boldRaw === 1
      || String(boldRaw).toLowerCase() === "true"
      || (Number.isFinite(fontWeightRaw) && fontWeightRaw >= 600),
    italic: attributes.italic === true || attributes.italic === 1 || style.italic === true,
    fontFamily: style.fontFamily || style.font_family || "",
    underline: underlineType > 0 && underlineStyle !== 15,
    strike: strikeType > 0 && strikeStyle !== 15,
    fontSizePx: fontSizePx > 0 ? clamp(fontSizePx, 8, 40) : 0,
    color: toCssColor(textColor),
    backgroundColor: shadeColor && !isWhiteColor(shadeColor) ? toCssColor(shadeColor) : "",
    letterSpacingEm: Number.isFinite(letterSpacingEm) ? clamp(letterSpacingEm, -0.3, 0.6) : 0,
    textStrokeWidthPx: strokeWidthPx,
    textStrokeColor: strokeWidthPx > 0 ? (toCssColor(outlineColor) || "currentColor") : "",
    textShadow
  };
  return normalized;
}

function normalizeParagraphStyle(shape = {}) {
  const source = shape && typeof shape === "object" ? shape : {};
  const attributes1 = source.attributes1 && typeof source.attributes1 === "object" ? source.attributes1 : {};
  const lineSpacingRaw = Number(source.line_spacing || source.line_spacing_old || source.lineSpacing || source.lineSpacingOld || 0);
  const hasPercentLikeLineHeight = Number.isFinite(lineSpacingRaw) && lineSpacingRaw >= 100 && lineSpacingRaw <= 400;
  const lineHeight = hasPercentLikeLineHeight ? clamp(lineSpacingRaw / 100, 1, 2.8) : 0;
  const indentPx = hwpUnitToPx(source.indent || 0);
  const leftPx = hwpUnitToPx(source.left_margin || source.leftMargin || 0);
  const rightPx = hwpUnitToPx(source.right_margin || source.rightMargin || 0);
  const topPx = hwpUnitToPx(source.top_spacing || source.topSpacing || 0);
  const bottomPx = hwpUnitToPx(source.bottom_spacing || source.bottomSpacing || 0);
  return {
    textAlign: normalizeAlign(attributes1.align || source.align),
    lineHeight,
    textIndentPx: Number.isFinite(indentPx) ? clamp(indentPx, -120, 200) : 0,
    marginLeftPx: Number.isFinite(leftPx) ? clamp(leftPx, 0, 260) : 0,
    marginRightPx: Number.isFinite(rightPx) ? clamp(rightPx, 0, 260) : 0,
    marginTopPx: Number.isFinite(topPx) ? clamp(topPx, 0, 140) : 0,
    marginBottomPx: Number.isFinite(bottomPx) ? clamp(bottomPx, 0, 140) : 0
  };
}

function toCssStyleText(styleObject = {}) {
  const entries = Object.entries(styleObject).filter(([, value]) => value !== undefined && value !== null && value !== "");
  if (!entries.length) {
    return "";
  }
  return entries.map(([key, value]) => `${key}:${value}`).join(";");
}

function inlineStyleToCss(style = {}) {
  const source = style && typeof style === "object" ? style : {};
  const css = {};
  if (source.bold) {
    css["font-weight"] = "700";
  }
  if (source.fontFamily) {
    css["font-family"] = `"${String(source.fontFamily).replace(/"/g, "")}", "HCR Batang", "Haansoft Batang", Batang, "Malgun Gothic", sans-serif`;
  }
  if (source.italic) {
    css["font-style"] = "italic";
  }
  if (source.underline && source.strike) {
    css["text-decoration"] = "underline line-through";
  } else if (source.underline) {
    css["text-decoration"] = "underline";
  } else if (source.strike) {
    css["text-decoration"] = "line-through";
  }
  if (Number(source.fontSizePx) > 0) {
    css["font-size"] = `${Number(source.fontSizePx).toFixed(2)}px`;
  }
  if (source.color) {
    css.color = source.color;
  }
  if (source.backgroundColor) {
    css["background-color"] = source.backgroundColor;
  }
  if (Number(source.letterSpacingEm)) {
    css["letter-spacing"] = `${Number(source.letterSpacingEm).toFixed(3)}em`;
  }
  if (Number(source.textStrokeWidthPx) > 0 && source.textStrokeColor) {
    css["-webkit-text-stroke"] = `${Number(source.textStrokeWidthPx).toFixed(2)}px ${source.textStrokeColor}`;
    css["paint-order"] = "stroke fill";
  }
  if (source.textShadow) {
    css["text-shadow"] = source.textShadow;
  }
  return toCssStyleText(css);
}

function paragraphStyleToCss(style = {}) {
  const source = style && typeof style === "object" ? style : {};
  const css = {};
  if (source.textAlign) {
    css["text-align"] = source.textAlign;
  }
  if (Number(source.lineHeight) > 0) {
    css["line-height"] = Number(source.lineHeight).toFixed(2);
  }
  if (Number(source.textIndentPx)) {
    css["text-indent"] = `${Number(source.textIndentPx).toFixed(1)}px`;
  }
  if (Number(source.marginLeftPx) > 0) {
    css["margin-left"] = `${Number(source.marginLeftPx).toFixed(1)}px`;
  }
  if (Number(source.marginRightPx) > 0) {
    css["margin-right"] = `${Number(source.marginRightPx).toFixed(1)}px`;
  }
  if (Number(source.marginTopPx) > 0) {
    css["margin-top"] = `${Number(source.marginTopPx).toFixed(1)}px`;
  }
  if (Number(source.marginBottomPx) > 0) {
    css["margin-bottom"] = `${Number(source.marginBottomPx).toFixed(1)}px`;
  }
  return toCssStyleText(css);
}

function createHwpCharStyleResolver(document = {}) {
  const charShapes = Array.isArray(document?.doc_info?.char_shapes) ? document.doc_info.char_shapes : [];
  const faceNames = Array.isArray(document?.doc_info?.face_names) ? document.doc_info.face_names : [];
  const resolveFontFamily = (shape = {}) => {
    const fontId = Number(
      shape?.font_ids?.korean
      ?? shape?.fontIds?.korean
      ?? shape?.font_id
      ?? shape?.fontId
      ?? -1
    );
    if (!Number.isFinite(fontId) || fontId < 0 || fontId >= faceNames.length) {
      return "";
    }
    const face = faceNames[fontId];
    return String(face?.name || face?.default_font_name || "").trim();
  };
  return (shapeId) => {
    const index = Number(shapeId);
    if (!Number.isFinite(index) || index < 0 || index >= charShapes.length) {
      return normalizeInlineStyle({});
    }
    const shape = charShapes[index] || {};
    return normalizeInlineStyle({
      ...shape,
      fontFamily: resolveFontFamily(shape)
    });
  };
}

function createHwpParaStyleResolver(document = {}) {
  const paraShapes = Array.isArray(document?.doc_info?.para_shapes) ? document.doc_info.para_shapes : [];
  return (shapeId) => {
    const index = Number(shapeId);
    if (!Number.isFinite(index) || index < 0 || index >= paraShapes.length) {
      return normalizeParagraphStyle({});
    }
    return normalizeParagraphStyle(paraShapes[index]);
  };
}

function createStructuredCharStyleResolver(document = {}) {
  const props = Array.isArray(document?.head?.charProperties) ? document.head.charProperties : [];
  const byId = new Map();
  props.forEach((item, index) => {
    if (!item || typeof item !== "object") {
      return;
    }
    const explicitId = Number(item.id);
    if (Number.isFinite(explicitId)) {
      byId.set(explicitId, item);
    }
    byId.set(index, item);
  });
  return (charPrIDRef) => {
    const id = Number(charPrIDRef);
    if (!Number.isFinite(id)) {
      return normalizeInlineStyle({});
    }
    return normalizeInlineStyle(byId.get(id) || {});
  };
}

function createStructuredParaStyleResolver(document = {}) {
  const props = Array.isArray(document?.head?.paraProperties) ? document.head.paraProperties : [];
  const byId = new Map();
  props.forEach((item, index) => {
    if (!item || typeof item !== "object") {
      return;
    }
    const explicitId = Number(item.id);
    if (Number.isFinite(explicitId)) {
      byId.set(explicitId, item);
    }
    byId.set(index, item);
  });
  return (paraPrIDRef) => {
    const id = Number(paraPrIDRef);
    if (!Number.isFinite(id)) {
      return normalizeParagraphStyle({});
    }
    return normalizeParagraphStyle(byId.get(id) || {});
  };
}

module.exports = {
  hwpUnitToPx,
  toCssColor,
  normalizeInlineStyle,
  normalizeParagraphStyle,
  inlineStyleToCss,
  paragraphStyleToCss,
  createHwpCharStyleResolver,
  createHwpParaStyleResolver,
  createStructuredCharStyleResolver,
  createStructuredParaStyleResolver
};
