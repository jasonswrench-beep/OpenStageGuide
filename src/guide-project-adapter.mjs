// Keeps a guide in a project file while the pane works on it.
//
// It wraps a document that can export and import its state (the in-memory document of guide-word.mjs) and gives the pane the very same
// interface. When the pane changes anything, the project is saved a moment later. Opening reads the file once, before the pane looks at
// the document.
//
// Three rules matter more than the rest:
//  * A file that could not be read is never written over. The pane is told why, and saving stays off until the program is started again.
//  * Saving never loses a change: if a save fails, the guide stays marked as not saved and the next change (or a retry) saves it all.
//  * The guide cannot be changed before it has opened, or after it failed to open: work that cannot be saved is refused, not lost.
//
// read() gives the project's text, or nothing when there is no file yet. write(text) stores it. Both are supplied by the host: the
// browser asks the program's own server; a test hands in a few lines.
import {parseProject,serializeProject} from './guide-project.mjs';

// What the program says when asked for a PDF. The pane shows a message that starts this way as it is (guide-pane.mjs, makePdf).
export const NO_PDF='The program does not make PDFs yet. Choose File > Save as Word file, open that file in Word, and choose File > Export > Create PDF/XPS, with "Document structure tags for accessibility" ticked.';

// What the pane may ask of its document without changing it. Everything else counts as a change and is followed by a save.
// The few settings below it only describe the session (a stand-in file picker, the theme the pane draws with) and are not part of the guide.
const READ_ONLY=new Set([
  'kind','paragraphs','title','customCues','getLook','slideMemory','versionMemory','cursorIndex','reviewParts','sourcePackage','captureFile',
  'exportPdf','pickFiles','keepPaneOpen','onSelectionChanged','cursor','snapshot','lookState','autoShow','openedDocuments','downloads','exportState','formSettings','design','sessions',
  'setTheme','setPicker','refuse','setAutoShow','fireSelection','offerDownload','setCapturedFile','place','historyState','mediaOf','listStyles','styleEdits','styleUse'
]);

// onChange() is called after every change to the guide, before it is saved, so what shows the guide can draw it again at once.
// `exportPdf` is how the program makes a PDF (File > Save as PDF; it gives the PDF as base64): without it the document says the program makes none.
export function createProjectAdapter(document,{read,write,delayMs=400,soon=(fn,ms)=>setTimeout(fn,ms),cancel=t=>clearTimeout(t),onStatus=()=>{},onChange=()=>{},exportPdf}={}){
  let timer=null,inFlight=Promise.resolve(),dirty=false,blocked=false,loaded=false,state='new',message='';
  const set=(s,m='')=>{state=s;message=m;onStatus({state:s,message:m});};
  const save=()=>{
    if(blocked||!dirty)return inFlight;
    dirty=false;
    set('saving');
    inFlight=inFlight.then(async()=>{
      try{
        await write(serializeProject(document.exportState()));
        if(!dirty)set('saved');
      }catch(e){
        // A page that is out of date (another guide was opened in another tab or window) stops saving for good: it must never write its guide over
        // the one that is open now. The person is told to reload; what was typed here since is not saved from this page.
        if(e?.stale){
          blocked=true;dirty=false;
          set('error',`${String(e.message).replace(/\.$/,'')}. Nothing more is saved from this page.`);
          return;
        }
        dirty=true;
        set('error',`The guide could not be saved: ${e?.message||'the file could not be written'}. Your changes are still here; the next change tries again.`);
      }
    });
    return inFlight;
  };
  const schedule=()=>{
    if(blocked)return;
    dirty=true;
    set('unsaved');
    if(timer!==null)cancel(timer);
    timer=soon(()=>{timer=null;save();},delayMs);
  };
  const flush=()=>{
    if(timer!==null){cancel(timer);timer=null;}
    return save();
  };
  const ready=(async()=>{
    let text;
    try{text=await read();}catch(e){
      blocked=true;set('error',`The project could not be read: ${e?.message||'the file could not be opened'}. Nothing was changed.`);
      throw e;
    }
    if(text===null||text===undefined||text===''){loaded=true;set('new');return {opened:false};}
    try{
      const {guide}=parseProject(text);
      document.importState(guide);
    }catch(e){
      blocked=true;set('error',`${e.message} The file was not changed, and saving is off until the program is started again.`);
      throw e;
    }
    loaded=true;
    set('saved');
    return {opened:true};
  })();
  ready.catch(()=>{});
  // The program has no Word to ask for a PDF (the stand-in document would give an empty file), so it says so instead of saving one that is not real.
  const own={ready,flush,status:()=>({state,message}),blocked:()=>blocked,exportPdf:exportPdf?async()=>exportPdf():async()=>{throw new Error(NO_PDF);}};
  return new Proxy(document,{
    get(target,prop){
      if(Object.hasOwn(own,prop))return own[prop];
      const v=target[prop];
      if(typeof v!=='function'||typeof prop==='symbol'||READ_ONLY.has(prop))return v;
      return (...args)=>{
        if(blocked)throw new Error(message||'This guide could not be opened, so it cannot be changed.');
        if(!loaded)throw new Error('The guide is still opening. Try again in a moment.');
        const out=v.apply(target,args);
        const changed=()=>{schedule();try{onChange();}catch{/* a failing view must not stop the guide being saved */}};
        if(out&&typeof out.then==='function')return out.then(x=>{changed();return x;});
        changed();
        return out;
      };
    }
  });
}
