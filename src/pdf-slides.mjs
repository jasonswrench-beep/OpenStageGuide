// Slide pictures made from a PDF of the presentation. A presentation's slides cannot be drawn without the presentation program, but the
// program can save the whole deck as one PDF (File > Save a Copy > PDF), and that is far easier for an author than exporting a picture
// of every slide. The file window opens the PDF with pdf.js (public/vendor/pdf.min.js), draws each page on a canvas and hands the
// pictures on as if the author had chosen them. This file holds the decisions (which page is which slide) and the drawing loop, with
// the PDF library and the canvas passed in, so both can be tested without a browser.
export const PICTURE_WIDTH=1280;     // pixels across a slide picture (a 16:9 slide is then 720 tall)
export const JPEG_QUALITY=0.88;

export const isPdfFile=f=>/\.pdf$/i.test(String(f?.name||''))||String(f?.type||'')==='application/pdf';

// Which slide each page of the PDF shows, by the slide's number in the presentation. PowerPoint leaves hidden slides out of a PDF unless
// told not to, so a PDF with one page for each slide that is not hidden is as good as one with a page for every slide.
export function pageSlideOrders(pageCount,slides){
  const all=slides.map(s=>s.order),shown=slides.filter(s=>!s.hidden).map(s=>s.order);
  if(pageCount===all.length)return all;
  if(pageCount===shown.length)return shown;
  throw new Error(`The PDF has ${pageCount} page${pageCount===1?'':'s'}, but the presentation has ${all.length} slide${all.length===1?'':'s'}${shown.length!==all.length?` (${shown.length} not hidden)`:''}. Save the PDF from this same presentation with every slide, then choose it again.`);
}

// data: the PDF's bytes. pdfjs: the pdf.js library (getDocument). makeCanvas(width,height): {canvas,context}. toFile(canvas,name): a picture
// file for the pane. Resolves the list of pictures, one for each page, named Slide<number>.jpg so the pane matches them to slides.
export async function renderPdfSlides({data,slides,pdfjs,makeCanvas,toFile,width=PICTURE_WIDTH,onProgress=()=>{}}){
  let pdf;
  // eval stays off: a crafted PDF must not be able to run script through the library (CVE-2024-4367 in the pdf.js that shipped here).
  try{pdf=await pdfjs.getDocument({data,isEvalSupported:false}).promise;}
  catch(e){
    if(e?.name==='PasswordException')throw new Error('That PDF is protected with a password. Save a copy without one, then choose it again.');
    throw new Error('That file could not be read as a PDF. Save the presentation as a PDF again, then choose it.');
  }
  try{
    const orders=pageSlideOrders(pdf.numPages,slides),files=[];
    for(let i=0;i<pdf.numPages;i++){
      onProgress(i+1,pdf.numPages);
      const page=await pdf.getPage(i+1);
      const scale=width/page.getViewport({scale:1}).width,viewport=page.getViewport({scale});
      const {canvas,context}=makeCanvas(Math.max(1,Math.round(viewport.width)),Math.max(1,Math.round(viewport.height)));
      context.fillStyle='#ffffff';context.fillRect(0,0,canvas.width,canvas.height); // a page has no background of its own; a picture needs one
      await page.render({canvasContext:context,viewport}).promise;
      files.push(await toFile(canvas,`Slide${orders[i]}.jpg`));
      page.cleanup?.();
    }
    return files;
  }finally{await pdf.destroy?.();}
}
