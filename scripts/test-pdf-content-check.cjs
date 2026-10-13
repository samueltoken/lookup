const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const { PDFDocument, PDFName, StandardFonts, rgb } = require("pdf-lib");
const { inspectPdfCaptureContent } = require("../src/main/web/pdf-content-check");

function addImage(document, width = 1, height = 1) {
  const stream = document.context.flateStream(new Uint8Array(width * height * 3).fill(255), {
    Type: "XObject", Subtype: "Image", Width: width, Height: height,
    ColorSpace: "DeviceRGB", BitsPerComponent: 8
  });
  return document.context.register(stream);
}

function addPage(document, source = "/Tiny Do", objects = { Tiny: addImage(document) }) {
  const page = document.addPage([595.28, 841.89]);
  page.node.set(PDFName.of("Resources"), document.context.obj({ XObject: objects }));
  page.node.set(PDFName.of("Contents"), document.context.register(document.context.flateStream(source)));
  return page;
}

async function inspect(document) {
  return inspectPdfCaptureContent(await document.save({ useObjectStreams: false }));
}

test("all pages of 1x1 and 2x2 drawn images are rejected", async () => {
  const document = await PDFDocument.create();
  addPage(document, "q 595.28 0 0 595.28 0 246.61 cm /Tiny Do Q");
  addPage(document, "/Tiny Do", { Tiny: addImage(document, 2, 2) });
  const result = await inspect(document);
  assert.equal(result.checked, true);
  assert.equal(result.placeholderOnly, true);
  assert.equal(result.pageCount, 2);
  assert.deepEqual(result.placeholderPages, [1, 2]);
});

test("a normal high resolution scanned page is accepted", async () => {
  const document = await PDFDocument.create();
  addPage(document, "/Scan Do", { Scan: addImage(document, 300, 400) });
  const result = await inspect(document);
  assert.equal(result.checked, true);
  assert.equal(result.placeholderOnly, false);
  assert.equal(result.imageCount, 1);
});

test("ordinary text and vector documents are accepted", async () => {
  const document = await PDFDocument.create();
  const page = document.addPage();
  const font = await document.embedFont(StandardFonts.Helvetica);
  page.drawText("Text document", { font, x: 40, y: 500 });
  page.drawRectangle({ x: 40, y: 400, width: 100, height: 50, color: rgb(1, 0, 0) });
  const result = await inspect(document);
  assert.equal(result.checked, true);
  assert.equal(result.placeholderOnly, false);
  assert.equal(result.pages[0].hasText, true);
  assert.equal(result.pages[0].hasOtherContent, true);
});

test("tiny images alongside real text or painted vectors are accepted", async () => {
  for (const source of ["/Tiny Do BT (Do TJ) Tj ET", "/Tiny Do 10 10 30 30 re f"]) {
    const document = await PDFDocument.create();
    addPage(document, source);
    const result = await inspect(document);
    assert.equal(result.checked, true);
    assert.equal(result.placeholderOnly, false);
  }
});

test("a deliberately blank page prevents document-wide rejection", async () => {
  const document = await PDFDocument.create();
  addPage(document);
  document.addPage();
  const result = await inspect(document);
  assert.equal(result.placeholderOnly, false);
  assert.deepEqual(result.placeholderPages, [1]);
  assert.deepEqual(result.blankPages, [2]);
});

test("unused Resources images do not count as page content", async () => {
  const document = await PDFDocument.create();
  addPage(document, "/Tiny Do", { Tiny: addImage(document), UnusedScan: addImage(document, 200, 200) });
  const result = await inspect(document);
  assert.equal(result.placeholderOnly, true);
  assert.equal(result.imageCount, 1);
});

test("drawing one real image among tiny images is accepted", async () => {
  const document = await PDFDocument.create();
  addPage(document, "/Tiny Do /Scan Do", { Tiny: addImage(document), Scan: addImage(document, 200, 200) });
  const result = await inspect(document);
  assert.equal(result.placeholderOnly, false);
  assert.equal(result.imageCount, 2);
});

test("content arrays, comments and escaped image names are handled", async () => {
  const document = await PDFDocument.create();
  const page = addPage(document, "", { Im1: addImage(document) });
  const first = document.context.register(document.context.flateStream("% (text) Tj /Scan Do\nq /Im#31"));
  const second = document.context.register(document.context.flateStream("Do Q"));
  page.node.set(PDFName.of("Contents"), document.context.obj([first, second]));
  const result = await inspect(document);
  assert.equal(result.checked, true);
  assert.equal(result.placeholderOnly, true);
});

test("inherited page Resources resolve the actual image", async () => {
  const document = await PDFDocument.create();
  const page = addPage(document);
  const resources = page.node.get(PDFName.of("Resources"));
  page.node.Parent().set(PDFName.of("Resources"), resources);
  page.node.delete(PDFName.of("Resources"));
  const result = await inspect(document);
  assert.equal(result.checked, true);
  assert.equal(result.placeholderOnly, true);
});

test("Form XObjects use their own Resources and ignore unused objects", async () => {
  const document = await PDFDocument.create();
  const form = document.context.register(document.context.flateStream("/Tiny Do", {
    Type: "XObject", Subtype: "Form", BBox: [0, 0, 500, 500],
    Resources: { XObject: { Tiny: addImage(document), Scan: addImage(document, 200, 200) } }
  }));
  addPage(document, "/Form Do", { Form: form, Tiny: addImage(document, 200, 200) });
  const result = await inspect(document);
  assert.equal(result.checked, true);
  assert.equal(result.placeholderOnly, true);
  assert.equal(result.imageCount, 1);
});

test("Form text prevents rejection and literal operators are not executed", async () => {
  const document = await PDFDocument.create();
  const form = document.context.register(document.context.flateStream("BT (text with \\(Do\\) and TJ) Tj ET", {
    Type: "XObject", Subtype: "Form", BBox: [0, 0, 500, 500]
  }));
  addPage(document, "/Tiny Do /Form Do", { Tiny: addImage(document), Form: form });
  const result = await inspect(document);
  assert.equal(result.checked, true);
  assert.equal(result.placeholderOnly, false);
  assert.equal(result.pages[0].hasText, true);
});

test("Form XObjects without Resources inherit the invocation Resources", async () => {
  const document = await PDFDocument.create();
  const form = document.context.register(document.context.flateStream("/Tiny Do", {
    Type: "XObject", Subtype: "Form", BBox: [0, 0, 500, 500]
  }));
  addPage(document, "/Form Do", { Tiny: addImage(document), Form: form });
  const result = await inspect(document);
  assert.equal(result.checked, true);
  assert.equal(result.placeholderOnly, true);
});

test("recursive forms are accepted conservatively without infinite recursion", async () => {
  const document = await PDFDocument.create();
  const form = document.context.flateStream("/Self Do", {
    Type: "XObject", Subtype: "Form", BBox: [0, 0, 500, 500]
  });
  const ref = document.context.register(form);
  form.dict.set(PDFName.of("Resources"), document.context.obj({ XObject: { Self: ref } }));
  addPage(document, "/Tiny Do /Form Do", { Tiny: addImage(document), Form: ref });
  const result = await inspect(document);
  assert.equal(result.checked, false);
  assert.equal(result.placeholderOnly, false);
});

test("inline images and unresolved image references remain uncertain", async () => {
  for (const source of ["/Tiny Do BI /W 100 /H 100 ID binary EI", "/Tiny Do /Missing Do"]) {
    const document = await PDFDocument.create();
    addPage(document, source);
    const result = await inspect(document);
    assert.equal(result.checked, false);
    assert.equal(result.placeholderOnly, false);
    assert.deepEqual(result.uncertainPages, [1]);
  }
});

test("invalid text calls do not conceal uncertain page content", async () => {
  const document = await PDFDocument.create();
  addPage(document, "/Tiny Do BT /Invalid Tj ET");
  const result = await inspect(document);
  assert.equal(result.checked, false);
  assert.equal(result.placeholderOnly, false);
});

test("malformed input never claims a known blank capture", async () => {
  const result = await inspectPdfCaptureContent(Buffer.from("not a PDF"));
  assert.equal(result.checked, false);
  assert.equal(result.placeholderOnly, false);
  assert.equal(result.pageCount, 0);
  assert.ok(result.error);
});

const reportedPdf = process.env.LOOKUP_PDF_REGRESSION_SAMPLE;
test("optional regression sample detects all 12 placeholder pages", { skip: !reportedPdf || !fs.existsSync(reportedPdf) }, async () => {
  const result = await inspectPdfCaptureContent(fs.readFileSync(reportedPdf));
  assert.equal(result.checked, true);
  assert.equal(result.placeholderOnly, true);
  assert.equal(result.pageCount, 12);
  assert.deepEqual(result.placeholderPages, Array.from({ length: 12 }, (_, index) => index + 1));
});
