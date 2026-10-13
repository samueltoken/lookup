const {
  PDFDocument,
  PDFArray,
  PDFDict,
  PDFName,
  PDFNumber,
  PDFRawStream,
  decodePDFRawStream
} = require("pdf-lib");

const MAX_CONTENT_BYTES = 8 * 1024 * 1024;
const MAX_TOTAL_CONTENT_BYTES = 64 * 1024 * 1024;
const MAX_OPERATORS = 500000;
const MAX_FORM_DEPTH = 16;
const SAFE_OPERATORS = new Set((
  "q Q cm w J j M d ri i gs BT ET Tc Tw Tz TL Tf Tr Ts Td TD Tm T* " +
  "CS cs SC SCN sc scn G g RG rg K k m l c v y h re W W* n " +
  "MP DP BMC BDC EMC BX EX d0 d1"
).split(/\s+/));
const PAINTING_OPERATORS = new Set(["S", "s", "f", "F", "f*", "B", "B*", "b", "b*", "sh"]);
const TEXT_OPERATORS = new Set(["Tj", "TJ", "'", '"']);

function isSpace(char) {
  return char === "\0" || char === " " || char === "\t" || char === "\n" || char === "\r" || char === "\f";
}

function isDelimiter(char) {
  return !char || isSpace(char) || "()<>[]{}/%".includes(char);
}

// Content streams contain binary strings and comments. Searching them with a
// regular expression would mistake text such as "(Do Tj)" for PDF operators.
function makeTokenReader(source) {
  let offset = 0;
  function skipSpace() {
    while (offset < source.length) {
      if (isSpace(source[offset])) offset += 1;
      else if (source[offset] === "%") {
        while (offset < source.length && !"\r\n".includes(source[offset])) offset += 1;
      } else break;
    }
  }
  function read(depth = 0) {
    if (depth > 32) throw new Error("PDF operand nesting limit exceeded");
    skipSpace();
    if (offset >= source.length) return null;
    const char = source[offset++];
    if (char === "(") {
      let nesting = 1;
      let hasText = false;
      while (offset < source.length && nesting > 0) {
        const next = source[offset++];
        if (next === "\\") {
          if (offset >= source.length) break;
          const escaped = source[offset++];
          if (escaped === "\r" && source[offset] === "\n") offset += 1;
          if (escaped !== "\r" && escaped !== "\n") hasText = true;
        } else if (next === "(") {
          nesting += 1;
          hasText = true;
        } else if (next === ")") {
          nesting -= 1;
          if (nesting > 0) hasText = true;
        } else hasText = true;
      }
      if (nesting !== 0) throw new Error("Unterminated PDF string");
      return { type: "string", hasText };
    }
    if (char === "<" && source[offset] !== "<") {
      let hasText = false;
      while (offset < source.length && source[offset] !== ">") {
        if (!isSpace(source[offset])) hasText = true;
        offset += 1;
      }
      if (source[offset++] !== ">") throw new Error("Unterminated PDF hex string");
      return { type: "string", hasText };
    }
    if (char === "[" || (char === "<" && source[offset] === "<")) {
      const isArray = char === "[";
      if (!isArray) offset += 1;
      let hasText = false;
      while (true) {
        skipSpace();
        if ((isArray && source[offset] === "]") || (!isArray && source.slice(offset, offset + 2) === ">>")) {
          offset += isArray ? 1 : 2;
          return { type: isArray ? "array" : "dict", hasText };
        }
        const nested = read(depth + 1);
        if (!nested) throw new Error("Unterminated PDF composite operand");
        if (nested.type === "string" || nested.type === "array") hasText ||= nested.hasText;
      }
    }
    if (char === "/") {
      const start = offset;
      while (!isDelimiter(source[offset])) offset += 1;
      return { type: "name", value: source.slice(start, offset).replace(/#([\da-f]{2})/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16))) };
    }
    if (char === ")" || char === "]" || char === ">") throw new Error("Unexpected PDF delimiter");
    const start = offset - 1;
    while (!isDelimiter(source[offset])) offset += 1;
    const value = source.slice(start, offset);
    const operand = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(value) || ["true", "false", "null"].includes(value);
    return { type: operand ? "operand" : "operator", value };
  }
  return read;
}

function emptyPageReport(pageNumber) {
  return {
    pageNumber,
    placeholderOnly: false,
    imageCount: 0,
    placeholderImageCount: 0,
    hasText: false,
    hasOtherContent: false,
    uncertain: false
  };
}

function markUncertain(report, error) {
  report.uncertain = true;
  if (!report.error) report.error = error instanceof Error ? error.message : String(error);
}

function readContent(stream, budget) {
  if (!(stream instanceof PDFRawStream)) throw new Error("Unsupported PDF content stream");
  if (budget.decoded.has(stream)) return budget.decoded.get(stream);
  // Use a bounded read, rather than decompressing an entire oversized stream.
  const bytes = decodePDFRawStream(stream).getBytes(MAX_CONTENT_BYTES + 1);
  if (bytes.length > MAX_CONTENT_BYTES) throw new Error("PDF content stream size limit exceeded");
  budget.byteCount += bytes.length;
  if (budget.byteCount > MAX_TOTAL_CONTENT_BYTES) throw new Error("PDF content inspection size limit exceeded");
  const source = Buffer.from(bytes).toString("latin1");
  budget.decoded.set(stream, source);
  return source;
}

function inspectStream(source, resources, report, budget, activeForms, depth) {
  if (depth > MAX_FORM_DEPTH) throw new Error("PDF form nesting limit exceeded");
  const readToken = makeTokenReader(source);
  let operands = [];
  let token;
  while ((token = readToken())) {
    if (token.type !== "operator") {
      operands.push(token);
      continue;
    }
    budget.operatorCount += 1;
    if (budget.operatorCount > MAX_OPERATORS) throw new Error("PDF content operator limit exceeded");
    const operator = token.value;
    if (TEXT_OPERATORS.has(operator)) {
      const lastOperand = operands[operands.length - 1];
      const expectedType = operator === "TJ" ? "array" : "string";
      const expectedCount = operator === '"' ? 3 : 1;
      if (operands.length !== expectedCount || !lastOperand || lastOperand.type !== expectedType) throw new Error("Invalid PDF text invocation");
      report.hasText ||= operands.some((operand) => (operand.type === "string" || operand.type === "array") && operand.hasText);
    } else if (PAINTING_OPERATORS.has(operator)) {
      report.hasOtherContent = true;
    } else if (operator === "Do") {
      const name = operands[operands.length - 1];
      if (operands.length !== 1 || !name || name.type !== "name") throw new Error("Invalid PDF XObject invocation");
      const xObjects = resources instanceof PDFDict ? resources.lookup(PDFName.of("XObject")) : undefined;
      const object = xObjects instanceof PDFDict ? xObjects.lookup(PDFName.of(name.value)) : undefined;
      if (!(object instanceof PDFRawStream)) throw new Error("Missing PDF XObject stream");
      const subtype = object.dict.lookup(PDFName.of("Subtype"));
      if (subtype === PDFName.of("Image")) {
        const width = object.dict.lookup(PDFName.of("Width"));
        const height = object.dict.lookup(PDFName.of("Height"));
        if (!(width instanceof PDFNumber) || !(height instanceof PDFNumber)) throw new Error("Invalid PDF image dimensions");
        const w = width.asNumber();
        const h = height.asNumber();
        if (!Number.isInteger(w) || !Number.isInteger(h) || w <= 0 || h <= 0) throw new Error("Invalid PDF image dimensions");
        const tiny = w <= 2 && h <= 2;
        report.imageCount += 1;
        if (tiny) report.placeholderImageCount += 1;
      } else if (subtype === PDFName.of("Form")) {
        if (activeForms.has(object)) throw new Error("Cyclic PDF form reference");
        const ownResources = object.dict.lookup(PDFName.of("Resources"));
        if (ownResources && !(ownResources instanceof PDFDict)) throw new Error("Invalid PDF form resources");
        activeForms.add(object);
        try {
          inspectStream(readContent(object, budget), ownResources || resources, report, budget, activeForms, depth + 1);
        } finally {
          activeForms.delete(object);
        }
      } else throw new Error("Unsupported PDF XObject subtype");
    } else if (operator === "BI") {
      // Inline image binary data needs a renderer/parser. Do not label such a
      // legitimate PDF as a blank capture based on incomplete inspection.
      throw new Error("Inline PDF image requires visual inspection");
    } else if (!SAFE_OPERATORS.has(operator)) {
      throw new Error(`Unsupported PDF operator: ${operator}`);
    }
    operands = [];
  }
  if (operands.length) throw new Error("Unconsumed PDF content operands");
}

/**
 * Detect the known failed-capture shape without rejecting ordinary scanned or
 * text PDFs. Only images actually drawn by Do count; unused resource entries
 * do not affect the result. A blank page or any uncertain content prevents the
 * document-wide hard failure. This is deliberately not a general PDF renderer.
 *
 * placeholderPages and blankPages contain one-based page numbers.
 */
async function inspectPdfCaptureContent(buffer, { stopWhenNonPlaceholder = false } = {}) {
  const result = {
    checked: false,
    placeholderOnly: false,
    pageCount: 0,
    placeholderPages: [],
    blankPages: [],
    uncertainPages: [],
    imageCount: 0,
    placeholderImageCount: 0,
    pages: []
  };
  try {
    const document = await PDFDocument.load(buffer, { ignoreEncryption: true, updateMetadata: false });
    result.pageCount = document.getPageCount();
    if (document.isEncrypted) throw new Error("Encrypted PDF requires visual inspection");
    const budget = { byteCount: 0, operatorCount: 0, decoded: new WeakMap() };
    for (const [index, page] of document.getPages().entries()) {
      const report = emptyPageReport(index + 1);
      result.pages.push(report);
      try {
        const contents = page.node.Contents();
        const streams = contents instanceof PDFArray ? contents.asArray().map((ref) => document.context.lookup(ref)) : contents ? [contents] : [];
        const source = streams.map((stream) => readContent(stream, budget)).join("\n");
        inspectStream(source, page.node.Resources(), report, budget, new Set(), 0);
        const annotations = page.node.Annots();
        if (annotations && annotations.size()) report.hasOtherContent = true;
      } catch (error) {
        markUncertain(report, error);
      }
      report.placeholderOnly = !report.uncertain && !report.hasText && !report.hasOtherContent && report.imageCount > 0 && report.imageCount === report.placeholderImageCount;
      if (report.placeholderOnly) result.placeholderPages.push(report.pageNumber);
      if (!report.uncertain && !report.hasText && !report.hasOtherContent && report.imageCount === 0) result.blankPages.push(report.pageNumber);
      if (report.uncertain) result.uncertainPages.push(report.pageNumber);
      result.imageCount += report.imageCount;
      result.placeholderImageCount += report.placeholderImageCount;
      if (stopWhenNonPlaceholder && !report.placeholderOnly) {
        result.partial = index + 1 < result.pageCount;
        break;
      }
    }
    result.checked = result.uncertainPages.length === 0;
    result.placeholderOnly = result.pageCount > 0 && result.placeholderPages.length === result.pageCount;
  } catch (error) {
    result.error = error instanceof Error ? error.message : String(error);
  }
  return result;
}

module.exports = { inspectPdfCaptureContent };
