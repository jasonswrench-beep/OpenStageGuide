// The file window (picker.html). Opened from the pane with Office's dialog API, because a file chooser cannot be opened from inside the
// pane itself. It reads and shrinks what was chosen, sends it to the pane as text messages, and the pane closes this window.
import {prepareImage} from './image-prepare.mjs';
import {readPptx} from './pptx.mjs';
import {readChosen,readRawFiles,resultMessages} from './file-pick.mjs';
import {isPdfFile,renderPdfSlides,JPEG_QUALITY} from './pdf-slides.mjs';

const $=id=>document.getElementById(id);
const query=new URLSearchParams(location.search),asked=query.get('kind');
const kind=asked==='slides'?'slides':asked==='files'?'files':'picture';
const say=text=>{$('pick-status').textContent=text;};
const send=text=>globalThis.Office.context.ui.messageParent(text);

$('picker-heading').textContent=kind==='slides'?'Choose a presentation and its slide pictures':kind==='files'?(query.get('label')?`Choose: ${query.get('label')}`:'Choose a file'):'Choose a picture';
$('picker-picture').hidden=kind!=='picture';
$('picker-slides').hidden=kind!=='slides';
$('picker-files').hidden=kind!=='files';
// For any other file the pane's own file box wanted: the same kinds of file, one or several, with the box's own name.
if(kind==='files'){
  const accept=query.get('accept');if(accept)$('pick-any').accept=accept;
  $('pick-any').multiple=query.get('multiple')==='1';
  if(query.get('label'))$('pick-any-label').textContent=query.get('label');
}

// A PDF of the presentation, drawn page by page with pdf.js into one picture for each slide.
async function picturesFromPdf(pdfFile,slides){
  const pdfjs=globalThis.pdfjsLib;
  if(!pdfjs)throw new Error('The PDF reader did not load. Close this window and try again.');
  // The worker sits beside pdf.min.js in the build folder (/build/<id>/vendor/), found from this module's own address: on the public
  // site there is no /vendor/ at the root, so a plain '/vendor/...' answered 404 there (the local server happens to serve one).
  pdfjs.GlobalWorkerOptions.workerSrc=new URL('../vendor/pdf.worker.min.js',import.meta.url).href;
  const data=new Uint8Array(await pdfFile.arrayBuffer());
  const files=await renderPdfSlides({data,slides,pdfjs,
    makeCanvas:(w,h)=>{const canvas=document.createElement('canvas');canvas.width=w;canvas.height=h;return {canvas,context:canvas.getContext('2d')};},
    toFile:async(canvas,name)=>{
      const blob=await new Promise((resolve,reject)=>canvas.toBlob(b=>b?resolve(b):reject(new Error('A slide picture could not be made.')),'image/jpeg',JPEG_QUALITY));
      return new File([blob],name,{type:'image/jpeg'});
    },
    onProgress:(i,n)=>say(`Making slide picture ${i} of ${n}...`)});
  const out=[];
  for(const f of files)out.push(await prepareImage(f));
  return out;
}

let warnedNoPictures=false;
// Choosing pictures (or a PDF) after the warning puts the window back as it was.
function pictureChoiceChanged(){warnedNoPictures=false;$('pick-go').textContent='Use these files';}
async function go(){
  $('pick-go').disabled=true;
  try{
    let deckFile=null,pictureFiles,pdfFile=null,rawFiles=null;
    if(kind==='files'){
      rawFiles=Array.from($('pick-any').files||[]);
      if(!rawFiles.length)throw new Error('Choose a file first.');
    }else if(kind==='slides'){
      deckFile=$('pick-deck').files?.[0]||null;
      if(!deckFile)throw new Error('Choose the presentation (.pptx) first.');
      const chosenPictures=Array.from($('pick-slide-pictures').files||[]),pdfs=chosenPictures.filter(isPdfFile);
      if(pdfs.length>1)throw new Error('Choose one PDF, not several.');
      if(pdfs.length&&pdfs.length!==chosenPictures.length)throw new Error('Choose either one PDF or the slides saved as pictures, not both.');
      pdfFile=pdfs[0]||null;pictureFiles=pdfFile?[]:chosenPictures;
      // Slides with no pictures go in as text only. Said once, before anything is sent, so it is never a surprise in the document
      // (the author's Word, 2026-11-07: "we're still not seeing the slides"); pressing the button again goes on without them.
      if(!chosenPictures.length&&!warnedNoPictures){
        warnedNoPictures=true;
        say('You have not chosen a PDF or slide pictures, so the slides will go in as text only, with no pictures. To see the slides, choose a PDF of the presentation above (in your presentation program: save a copy as a PDF). Or choose "Use without pictures" to go on as it is.');
        $('pick-go').textContent='Use without pictures';
        $('pick-go').disabled=false;
        return;
      }
    }else{
      pictureFiles=Array.from($('pick-picture').files||[]);
      if(!pictureFiles.length)throw new Error('Choose a picture first.');
    }
    say('Reading the files...');
    // Any other file goes to the pane exactly as it is; the pane reads it the way it always has.
    const chosen=rawFiles?{deck:null,files:await readRawFiles(rawFiles)}:await readChosen({deckFile,pictureFiles,prepare:prepareImage,
      readDeck:async f=>readPptx(await f.arrayBuffer(),f.name,globalThis.JSZip,{details:true})});
    if(pdfFile)chosen.files=await picturesFromPdf(pdfFile,chosen.deck.slides);
    const messages=resultMessages(chosen);
    say('Sending to the pane...');
    for(let i=0;i<messages.length;i++){
      send(messages[i]);
      if(i%8===7)await new Promise(r=>setTimeout(r,0)); // a breath now and then, so the messages do not pile up
    }
    say('Done. This window will close.');
  }catch(e){
    say((e&&e.message)||String(e));
    $('pick-go').disabled=false;
  }
}

function start(){
  $('pick-go').onclick=go;
  $('pick-slide-pictures').onchange=pictureChoiceChanged;
  $('pick-cancel').onclick=()=>send(JSON.stringify({t:'cancel'}));
  $('picker-heading').focus();
}
if(globalThis.Office?.onReady)globalThis.Office.onReady(start);else start();
