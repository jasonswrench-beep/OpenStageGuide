import test from 'node:test';
import assert from 'node:assert/strict';
import {isPdfFile,pageSlideOrders,renderPdfSlides,PICTURE_WIDTH} from '../src/pdf-slides.mjs';

const slides=(n,hidden=[])=>Array.from({length:n},(_,i)=>({order:i+1,title:`T${i+1}`,hidden:hidden.includes(i+1)}));

test('a PDF is known by its name or its type',()=>{
  assert.equal(isPdfFile({name:'Deck.PDF',type:''}),true);
  assert.equal(isPdfFile({name:'deck',type:'application/pdf'}),true);
  assert.equal(isPdfFile({name:'Slide1.png',type:'image/png'}),false);
  assert.equal(isPdfFile(null),false);
});

test('each page is a slide: all of them, or all that are not hidden, and anything else is refused in plain words',()=>{
  assert.deepEqual(pageSlideOrders(3,slides(3)),[1,2,3]);
  assert.deepEqual(pageSlideOrders(3,slides(4,[2])),[1,3,4],'PowerPoint leaves hidden slides out of a PDF');
  assert.deepEqual(pageSlideOrders(4,slides(4,[2])),[1,2,3,4],'or the author kept them in');
  assert.throws(()=>pageSlideOrders(5,slides(4,[2])),/The PDF has 5 pages, but the presentation has 4 slides \(3 not hidden\)\. Save the PDF from this same presentation/);
  assert.throws(()=>pageSlideOrders(1,slides(2)),/The PDF has 1 page, but the presentation has 2 slides\./);
});

// A stand-in for pdf.js and for a canvas: each page is 960 x 540 points, and drawing one records what was done.
function fakeLibrary({pages=3,fail=null}={}){
  const log=[];
  const pdfjs={getDocument:({data})=>({promise:fail?Promise.reject(fail):Promise.resolve({
    numPages:pages,
    getPage:async n=>({
      getViewport:({scale})=>({width:960*scale,height:540*scale}),
      render:({canvasContext,viewport})=>({promise:(async()=>{log.push(['render',n,Math.round(viewport.width),Math.round(viewport.height)]);})()}),
      cleanup:()=>log.push(['cleanup',n])
    }),
    destroy:async()=>log.push(['destroy'])
  })})};
  const makeCanvas=(w,h)=>({canvas:{width:w,height:h},context:{fillRect:(...a)=>log.push(['fill',...a]),set fillStyle(v){log.push(['style',v]);}}});
  const toFile=async(canvas,name)=>({name,width:canvas.width,height:canvas.height});
  return {pdfjs,makeCanvas,toFile,log};
}

test('every page is drawn on a white ground at the picture width and named for its slide',async()=>{
  const lib=fakeLibrary({pages:3});
  const progress=[];
  const files=await renderPdfSlides({data:new Uint8Array(4),slides:slides(3),...lib,onProgress:(i,n)=>progress.push([i,n])});
  assert.deepEqual(files.map(f=>[f.name,f.width,f.height]),[['Slide1.jpg',PICTURE_WIDTH,720],['Slide2.jpg',PICTURE_WIDTH,720],['Slide3.jpg',PICTURE_WIDTH,720]]);
  assert.deepEqual(progress,[[1,3],[2,3],[3,3]]);
  assert.deepEqual(lib.log.filter(e=>e[0]==='style'),[['style','#ffffff'],['style','#ffffff'],['style','#ffffff']]);
  assert.equal(lib.log.filter(e=>e[0]==='fill').length,3);
  assert.equal(lib.log.at(-1)[0],'destroy','the document is let go at the end');
});

test('with a hidden slide left out of the PDF, the pictures carry the slides\' own numbers',async()=>{
  const files=await renderPdfSlides({data:new Uint8Array(4),slides:slides(4,[2]),...fakeLibrary({pages:3})});
  assert.deepEqual(files.map(f=>f.name),['Slide1.jpg','Slide3.jpg','Slide4.jpg']);
});

test('a PDF that does not match the presentation, is protected or is not a PDF gives a reason the author can act on',async()=>{
  const wrongCount=fakeLibrary({pages:9});
  await assert.rejects(renderPdfSlides({data:new Uint8Array(4),slides:slides(3),...wrongCount}),/The PDF has 9 pages, but the presentation has 3 slides/);
  assert.equal(wrongCount.log.at(-1)[0],'destroy','the document is let go even when it is refused');
  const locked=Object.assign(new Error('x'),{name:'PasswordException'});
  await assert.rejects(renderPdfSlides({data:new Uint8Array(4),slides:slides(3),...fakeLibrary({fail:locked})}),/protected with a password/);
  await assert.rejects(renderPdfSlides({data:new Uint8Array(4),slides:slides(3),...fakeLibrary({fail:new Error('Invalid PDF structure')})}),/could not be read as a PDF/);
});
