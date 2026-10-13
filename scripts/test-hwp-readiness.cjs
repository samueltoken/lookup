const assert = require("node:assert/strict");
const { inspectHwpFamilyBytes, readHwpDocumentVariants, normalizeHwpxPackageForRenderer } = require("../src/main/hwp/document-readiness");
const { buildHwpxSectionLayout, buildHwpxPictureHtml, normalizeHwpxDocumentStyles } = require("../src/main/hwp/hwpx-layout");
const { HWPXBuilder, write, read, utils } = require("hwpx-js");

assert.throws(() => inspectHwpFamilyBytes(Buffer.alloc(0)), error => error.code === "EMPTY_DOCUMENT");
assert.throws(() => inspectHwpFamilyBytes(Buffer.from("not a HWP file")), error => error.code === "INVALID_HWP");
assert.throws(() => inspectHwpFamilyBytes(Buffer.from("d0cf11e0a1b11ae1", "hex")), error => error.code === "INVALID_HWP");
const ole = Buffer.alloc(512); Buffer.from("d0cf11e0a1b11ae1", "hex").copy(ole);
assert.equal(inspectHwpFamilyBytes(ole, ".hwpx").documentExt, ".hwp");
assert.equal(inspectHwpFamilyBytes(ole, ".hwpx").extensionMismatch, true);

const fixture = new HWPXBuilder().setPageSettings({width:297,height:210,marginLeft:22,marginTop:13}).addParagraph("첫 문단", {fontSize:20,bold:true,color:"#CC0033"})
  .addTable([["병합 셀",""],["왼쪽","오른쪽"]], {merges:[{row:0,col:0,rowSpan:1,colSpan:2}],colWidths:[30,50]}).build();
const zipped = write(fixture);
assert.equal(inspectHwpFamilyBytes(zipped, ".hwp").documentExt, ".hwpx");
const roundTrip = read(zipped);
const normalized = normalizeHwpxDocumentStyles(roundTrip);
const resolveStyle = require("../src/main/hwp/style-profile").createStructuredCharStyleResolver(normalized);
const resolved = resolveStyle(roundTrip.sections[0].paragraphs[0].runs[0].charPrIDRef);
assert.equal(resolved.bold,true);
assert.equal(resolved.color,"rgb(204, 0, 51)");
assert.ok(resolved.fontSizePx>15,"source font height is retained");
assert.ok(resolved.fontFamily.length>0,"source font reference is retained");
const unchanged = Buffer.from(zipped);
const repairedPackage = require("fflate").unzipSync(normalizeHwpxPackageForRenderer(zipped));
assert.match(Buffer.from(repairedPackage["Contents/content.hpf"]).toString("utf8"), /href="Contents\/section0.xml"/);
assert.deepEqual(Buffer.from(zipped), unchanged);
assert.equal(roundTrip.sections[0].paragraphs[1].runs[0].table.rows[0].cells[0].colSpan, 2);
const layout = buildHwpxSectionLayout(roundTrip.sections[0].def);
assert.match(layout.css, /size:297\.\d+mm 210\.\d+mm/);
const leftMargin = Number(layout.css.match(/margin:[^;]+\s([\d.]+)mm;/)[1]);
assert.ok(Math.abs(leftMargin-22)<0.02);
assert.match(layout.style, /page:lookupSection0/);
const image = buildHwpxPictureHtml({binDataIDRef:7,width:utils.mmToHwpunit(20),height:utils.mmToHwpunit(10)}, [{id:7,name:'図 " test',format:'png',data:Uint8Array.from([1,2,3])}]);
assert.match(image, /data:image\/png;base64,AQID/);
assert.ok(Math.abs(Number(image.match(/width:([\d.]+)px/)[1])-75.59)<0.03);
assert.ok(Math.abs(Number(image.match(/height:([\d.]+)px/)[1])-37.80)<0.03);
assert.match(image, /&quot;/);

const calls = [];
const recovered = readHwpDocumentVariants({
  toHtml: (_bytes, options) => { calls.push("html"); assert.equal(options.includeVersion,false); return "<p>살아있는 본문</p>"; },
  toJson: () => { calls.push("json"); throw new Error("unsupported record"); },
  toMarkdown: () => { calls.push("markdown"); throw new Error("bad table"); }
}, ole);
assert.deepEqual(calls,["html","json","markdown"]);
assert.equal(recovered.html,"<p>살아있는 본문</p>");
assert.equal(recovered.failures.length,2);
const structured = readHwpDocumentVariants({toHtml:()=>{throw Error("layout error")},toJson:()=>'{"table":true}',toMarkdown:()=>({markdown:"본문"})},ole);
assert.equal(structured.json,'{"table":true}');
assert.equal(structured.markdown,"본문");
console.log("HWP input, format recovery, independent parsers, page settings, image and merged-table checks passed.");
