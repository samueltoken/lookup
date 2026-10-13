// Run with Electron. Each PDF is produced by the real application conversion path.
const { app, BrowserWindow } = require("electron");
const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");
const assert = require("node:assert/strict");
const crypto = require("node:crypto");
const { HWPXBuilder, write } = require("hwpx-js");
const { PNG } = require("pngjs");
const { PDFDocument } = require("pdf-lib");
const root = path.resolve(__dirname,"..");
const output = path.join(root,"tmp","hwp-verification");
fs.mkdirSync(output,{recursive:true});
app.setPath("userData",path.join(output,"profile"));

app.whenReady().then(async()=>{
  const keeper = new BrowserWindow({ show:false, webPreferences:{ offscreen:true, backgroundThrottling:false } });
  try {
    const filename = path.join(root,"src/main/app-main.js");
    let source = fs.readFileSync(filename,"utf8");
    source = source.slice(0,source.lastIndexOf("\nif (gotSingleInstanceLock) {"));
    source = source.replace("app.requestSingleInstanceLock()","true");
    source += "\nmodule.exports={convertDocumentToPdf,convertHwpToPdfWithRhwp,buildStructuredCompatibilityHtmlFromDocument};";
    const compiled = new Module(filename);
    compiled.filename=filename; compiled.paths=Module._nodeModulePaths(path.dirname(filename));
    compiled._compile(source,filename);
    const test = compiled.exports;
    const image = new PNG({width:12,height:12});
    for(let i=0;i<image.data.length;i+=4){image.data[i]=220;image.data[i+1]=40;image.data[i+2]=60;image.data[i+3]=255;}
    const fixture = new HWPXBuilder().setPageSettings({width:297,height:210,marginLeft:22,marginTop:13})
      .addParagraph("HWPX_START_MARKER",{fontSize:20,bold:true,color:"#CC0033"})
      .addTable([["MERGED_MARKER",""],["LEFT_CELL_MARKER","RIGHT_CELL_MARKER"]],{merges:[{row:0,col:0,rowSpan:1,colSpan:2}],colWidths:[35,55]})
      .addImage(PNG.sync.write(image),"png",{width:30,height:20})
      .addParagraph("PAGE_BREAK_MARKER")
      .addParagraph("HWPX_END_MARKER").build();
    fixture.sections[0].paragraphs[3].runs.unshift({t:"break",breakType:"PAGE",charPrIDRef:0});
    const structuredHtml=test.buildStructuredCompatibilityHtmlFromDocument(require("../src/main/hwp/hwpx-layout").normalizeHwpxDocumentStyles(fixture),"fixture",{sourceExt:".hwpx"});
    assert.ok(structuredHtml.includes('colspan="2"'));
    assert.ok(structuredHtml.includes("data:image/png;base64,"));
    assert.ok(structuredHtml.includes("lookup-hwpx-page-break"));
    assert.ok(structuredHtml.includes("size:297."));
    assert.ok(structuredHtml.includes("rgb(204, 0, 51)"));
    assert.ok(structuredHtml.includes("font-size:26.67px"));
    assert.equal(structuredHtml.includes("원본 형식:"),false);
    const fixturePath=path.join(output,"한글 공백 그림 표.hwpx");
    fs.writeFileSync(fixturePath,write(fixture));
    const pdfjs=await import("pdfjs-dist/legacy/build/pdf.mjs");
    pdfjs.GlobalWorkerOptions.workerSrc=require("node:url").pathToFileURL(require.resolve("pdfjs-dist/legacy/build/pdf.worker.mjs")).href;
    async function inspect(bytes){
      const pdf=await PDFDocument.load(bytes),handle=await pdfjs.getDocument({data:new Uint8Array(bytes),disableFontFace:true}).promise;
      const text=[];let images=0;
      for(let n=1;n<=handle.numPages;n++){const page=await handle.getPage(n);text.push((await page.getTextContent()).items.map(item=>item.str).join(" "));const ops=await page.getOperatorList();images+=ops.fnArray.filter(fn=>[pdfjs.OPS.paintImageXObject,pdfjs.OPS.paintInlineImageXObject].includes(fn)).length;}
      await handle.destroy();return{pages:pdf.getPageCount(),sizes:pdf.getPages().map(page=>page.getSize()),text,images};
    }
    const rows=[];
    const start=Date.now(),converted=await test.convertDocumentToPdf(fixturePath,{forceCompatibility:true});
    const info=await inspect(converted.pdfBuffer);assert.ok(info.pages>=2);assert.ok(info.sizes[0].width>info.sizes[0].height);assert.ok(info.images>=1);
    for(const marker of ["HWPX_START_MARKER","MERGED_MARKER","LEFT_CELL_MARKER","RIGHT_CELL_MARKER","PAGE_BREAK_MARKER","HWPX_END_MARKER"])assert.ok(info.text.join(" ").replace(/\s/g,"").includes(marker),marker);
    assert.ok(info.text[0].includes("HWPX_START_MARKER"));assert.equal(info.text[0].includes("PAGE_BREAK_MARKER"),false);
    rows.push({kind:"hwpx",mode:converted.convertMode,ms:Date.now()-start,...info});
    fs.writeFileSync(path.join(output,"hwpx-fixture.pdf"),converted.pdfBuffer);
    const alternateFixture = await test.convertHwpToPdfWithRhwp(fixturePath,path.join(output,"rhwp-hwpx-fixture.pdf"));
    assert.equal(alternateFixture.ok,false,"empty alternative output must be rejected");
    assert.match(alternateFixture.reason,/rhwp_no_visible_content/);
    rows.push({kind:"rhwp-hwpx-empty-rejected",reason:alternateFixture.reason});
    const renamed=path.join(output,"실제 HWPX 확장자 hwp.hwp");fs.copyFileSync(fixturePath,renamed);
    const recovered=await test.convertDocumentToPdf(renamed,{forceCompatibility:true});assert.ok((await inspect(recovered.pdfBuffer)).text.join(" ").includes("HWPX_END_MARKER"));
    rows.push({kind:"mismatched-extension",mode:recovered.convertMode});
    const hwpSource=process.argv[2];
    const expectedHwpText=process.argv[3];
    if(hwpSource&&fs.existsSync(hwpSource)){
      const before=crypto.createHash("sha256").update(fs.readFileSync(hwpSource)).digest("hex");
      const hwpStart=Date.now();const real=await test.convertDocumentToPdf(hwpSource,{forceCompatibility:true});const realInfo=await inspect(real.pdfBuffer);
      assert.ok(realInfo.pages>0);assert.ok(realInfo.text.join(" ").trim().length>0);if(expectedHwpText)assert.ok(realInfo.text.join(" ").includes(expectedHwpText));assert.equal(crypto.createHash("sha256").update(fs.readFileSync(hwpSource)).digest("hex"),before);
      fs.writeFileSync(path.join(output,"public-hwp.pdf"),real.pdfBuffer);rows.push({kind:"real-public-hwp",mode:real.convertMode,ms:Date.now()-hwpStart,pages:realInfo.pages,chars:realInfo.text.join(" ").length});
      const alternative=await test.convertHwpToPdfWithRhwp(hwpSource,path.join(output,"rhwp-public-hwp.pdf"));assert.equal(alternative.ok,true,alternative.reason);
      const alternativeInfo=await inspect(fs.readFileSync(path.join(output,"rhwp-public-hwp.pdf")));assert.ok(alternativeInfo.text.join(" ").trim().length>0);if(expectedHwpText)assert.ok(alternativeInfo.text.join(" ").includes(expectedHwpText));rows.push({kind:"rhwp-recovery",originalPages:alternative.pageCount,geometry:alternative.pageGeometry,pages:alternativeInfo.pages,chars:alternativeInfo.text.join(" ").length});
    }
    assert.equal(BrowserWindow.getAllWindows().length,1,"all conversion windows close");
    fs.writeFileSync(path.join(output,"diagnostics.json"),JSON.stringify(rows,null,2));console.log(JSON.stringify(rows.map(({text,sizes,...row})=>row),null,2));
  }catch(error){console.error(error);process.exitCode=1;}finally{keeper.destroy();app.quit();}
});
