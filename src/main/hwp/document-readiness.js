const HWP_MAGIC = Buffer.from("d0cf11e0a1b11ae1", "hex");
const ZIP_MAGIC = Buffer.from("504b0304", "hex");
const path = require("node:path");
const { unzipSync, zipSync } = require("fflate");

function inspectHwpFamilyBytes(bytes, sourceExt = ".hwp") {
  const raw = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes || []);
  if (raw.length === 0) {
    const error = new Error("empty_document: hwp_input_empty");
    error.code = "EMPTY_DOCUMENT";
    throw error;
  }
  const format = raw.length >= 512 && raw.subarray(0, 8).equals(HWP_MAGIC)
    ? "hwp5"
    : raw.length >= 22 && raw.subarray(0, 4).equals(ZIP_MAGIC) ? "hwpx" : "unknown";
  if (format === "unknown") {
    const error = new Error("hwp_invalid_file_header");
    error.code = "INVALID_HWP";
    error.publicMessage = "손상된 한글 파일";
    throw error;
  }
  return {
    format,
    sourceExt: String(sourceExt || "").toLowerCase(),
    documentExt: format === "hwp5" ? ".hwp" : ".hwpx",
    extensionMismatch: (format === "hwp5" ? ".hwp" : ".hwpx") !== String(sourceExt || "").toLowerCase(),
    byteLength: raw.length
  };
}

function readHwpDocumentVariants(parser, bytes) {
  const raw = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes || []);
  const variants = { html: "", json: "", markdown: "", failures: [] };
  for (const [method, field] of [["toHtml", "html"], ["toJson", "json"], ["toMarkdown", "markdown"]]) {
    if (typeof parser?.[method] !== "function") {
      variants.failures.push({ method, reason: "method_unavailable" });
      continue;
    }
    try {
      const result = method === "toHtml" ? parser[method](raw, { includeVersion: false, includePageInfo: false }) : parser[method](raw);
      variants[field] = typeof result === "string" ? result : String(result?.[field] || "");
      if (!variants[field].trim()) variants.failures.push({ method, reason: "empty_result" });
    } catch (error) {
      variants.failures.push({ method, reason: String(error?.message || "parse_failed").slice(0, 250) });
    }
  }
  return variants;
}

function normalizeHwpxPackageForRenderer(bytes) {
  const raw = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes || []);
  if (!raw.subarray(0, 4).equals(ZIP_MAGIC)) return raw;
  const files = unzipSync(raw);
  let changed = false;
  for (const entry of Object.keys(files).filter(name => /(?:^|\/)content\.hpf$/i.test(name))) {
    const content = Buffer.from(files[entry]).toString("utf8");
    const rewritten = content.replace(/\bhref=(['"])([^'"]+)\1/g, (match, quote, href) => {
      if (files[href]) return match;
      const resolved = path.posix.normalize(path.posix.join(path.posix.dirname(entry), href));
      if (!files[resolved]) return match;
      changed = true;
      return `href=${quote}${resolved}${quote}`;
    });
    if (rewritten !== content) files[entry] = new Uint8Array(Buffer.from(rewritten));
  }
  return changed ? Buffer.from(zipSync(files, { level: 0 })) : raw;
}

module.exports = { inspectHwpFamilyBytes, readHwpDocumentVariants, normalizeHwpxPackageForRenderer };
