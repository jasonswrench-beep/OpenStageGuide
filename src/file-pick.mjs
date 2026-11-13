// Choosing files when the pane cannot. In Word's task pane an <input type="file"> does nothing (no file window ever opens: seen in the
// author's Word, 2026-11-06, with the mouse and with the keyboard), so the choice is made in a small window of its own, opened with
// Office's dialog API. That window is the page picker.html: it has the file inputs, reads and shrinks the pictures, and sends the
// result back to the pane as a run of small text messages (the dialog API carries text only). This file is both ends of that road:
// resultMessages / createReceiver (the messages), pickFilesWithDialog (the pane's end) and pickWithBrowserInputs (the plain-browser
// preview, where an <input type="file"> does work).
import {toBase64,prepareImage} from './image-prepare.mjs';

export const CHUNK_BYTES=48*1024;      // raw bytes in one message (a little under 66 KB of text)
export const MAX_FILES=300;           // a deck of up to 300 slides, one picture each
export const MAX_TOTAL_BYTES=64*1024*1024;

export const chunkBase64=(bytes,size=CHUNK_BYTES)=>{const out=[];for(let i=0;i<bytes.length;i+=size)out.push(toBase64(bytes.subarray(i,i+size)));return out.length?out:[''];};
export function fromBase64(text){
  if(typeof Buffer!=='undefined')return new Uint8Array(Buffer.from(text,'base64'));
  const s=atob(text),out=new Uint8Array(s.length);for(let i=0;i<s.length;i++)out[i]=s.charCodeAt(i);return out;
}

// What the dialog sends for {deck, files}: the deck (titles, notes and so on, as plain data), then each file's description and its
// bytes in pieces, then a closing "done". Each message is one text.
export function resultMessages({deck=null,files=[]}){
  if(files.length>MAX_FILES)throw new Error(`Choose ${MAX_FILES} files or fewer at a time.`);
  const total=files.reduce((a,f)=>a+f.bytes.length,0);
  if(total>MAX_TOTAL_BYTES)throw new Error(`Those files add up to ${(total/1048576).toFixed(0)} MB. Choose fewer or smaller files (the limit is ${MAX_TOTAL_BYTES/1048576} MB).`);
  const messages=[];
  if(deck)messages.push({t:'deck',deck});
  files.forEach((f,i)=>{
    const chunks=chunkBase64(f.bytes);
    messages.push({t:'file',i,meta:{name:f.name,type:f.type,mime:f.mime,width:f.width,height:f.height,shrunk:!!f.shrunk,size:f.bytes.length},parts:chunks.length});
    chunks.forEach((data,k)=>messages.push({t:'chunk',i,k,data}));
  });
  messages.push({t:'done',files:files.length});
  return messages.map(m=>JSON.stringify(m));
}

// The pane's end: feed it each message text; it says 'done' when the last one has arrived, and refuses anything that did not arrive whole.
export function createReceiver(){
  let deck=null,expected=null;
  const files=[];
  const problem=()=>new Error('The files did not arrive complete. Close the file window and try again.');
  return {
    receive(text){
      let m;
      try{m=JSON.parse(text);}catch{throw problem();}
      if(!m||typeof m!=='object')throw problem();
      if(m.t==='deck'){deck=m.deck;return 'more';}
      if(m.t==='file'){
        if(m.i!==files.length||!m.meta||!(m.parts>=1))throw problem();
        files.push({meta:m.meta,parts:new Array(m.parts).fill(null)});return 'more';
      }
      if(m.t==='chunk'){
        const f=files[m.i];
        if(!f||!(m.k>=0&&m.k<f.parts.length)||typeof m.data!=='string')throw problem();
        f.parts[m.k]=m.data;return 'more';
      }
      if(m.t==='done'){expected=m.files;return 'done';}
      if(m.t==='cancel')return 'cancel';
      throw problem();
    },
    result(){
      if(expected!==files.length)throw problem();
      const out=files.map(f=>{
        if(f.parts.some(p=>p===null))throw problem();
        const pieces=f.parts.map(fromBase64),bytes=new Uint8Array(pieces.reduce((a,p)=>a+p.length,0));
        let at=0;for(const p of pieces){bytes.set(p,at);at+=p.length;}
        if(bytes.length!==f.meta.size)throw problem();
        return {name:f.meta.name,type:f.meta.type,mime:f.meta.mime,width:f.meta.width,height:f.meta.height,shrunk:!!f.meta.shrunk,bytes};
      });
      return {deck,files:out};
    }
  };
}

// Office's reasons a dialog cannot open, in words an author can act on.
const DIALOG_ERRORS={
  12004:'Word does not trust the address of the file window.',
  12005:'The file window needs a secure (https) address.',
  12007:'A file window is already open. Close it, then try again.',
  12009:'Word asked whether this add-in may open a window, and the answer was no. Try again and choose Allow.',
  12011:'Word could not open the file window in this place. Try again with your cursor in the document.'
};
const dialogMessage=err=>DIALOG_ERRORS[err?.code]||`Word could not open the file window${err?.message?` (${err.message})`:''}. Close it if it is open, then try again.`;

// The address of the file window: its page, what to choose (`kind`: 'picture', 'slides' or 'files') and, for 'files', what the pane's
// own file box asked for (accept, multiple, label) so the window offers the same choice.
export function pickerAddress(url,kind,{accept='',multiple=false,label=''}={}){
  const q=new URLSearchParams({kind});
  if(kind==='files'){if(accept)q.set('accept',String(accept));if(multiple)q.set('multiple','1');if(label)q.set('label',String(label).slice(0,120));}
  return `${url}?${q.toString()}`;
}

// The pane's end of the dialog: opens picker.html for `kind` ('picture', 'slides' or 'files'), collects what it sends, closes it and
// resolves {deck, files} (or {cancelled:true} when the window was closed first).
export function pickFilesWithDialog({Office,url,kind,accept,multiple,label,timeoutMs=600000}){
  return new Promise((resolve,reject)=>{
    const ui=Office?.context?.ui;
    if(!ui?.displayDialogAsync){reject(new Error('This version of Word cannot open the file window. Update Word, then try again.'));return;}
    ui.displayDialogAsync(pickerAddress(url,kind,{accept,multiple,label}),{height:58,width:44,displayInIframe:false},res=>{
      if(res.status!=='succeeded'){reject(new Error(dialogMessage(res.error)));return;}
      const dialog=res.value,rx=createReceiver();
      let settled=false,timer=null;
      const finish=fn=>{if(settled)return;settled=true;clearTimeout(timer);try{dialog.close();}catch{/* it may already be closed */}fn();};
      timer=setTimeout(()=>finish(()=>reject(new Error('The file window took too long. Close it and try again.'))),timeoutMs);
      dialog.addEventHandler(Office.EventType.DialogMessageReceived,arg=>{
        try{
          const state=rx.receive(arg.message);
          if(state==='done')finish(()=>resolve(rx.result()));
          else if(state==='cancel')finish(()=>resolve({cancelled:true}));
        }
        catch(e){finish(()=>reject(e));}
      });
      // The author closed the window (12006) or it went away: nothing was chosen.
      dialog.addEventHandler(Office.EventType.DialogEventReceived,()=>finish(()=>resolve({cancelled:true})));
    });
  });
}

// What the picker page does with the files the author chose: read and shrink each picture, read the presentation. Used by the picker
// page and by the preview's plain inputs.
export async function readChosen({deckFile=null,pictureFiles=[],prepare=prepareImage,readDeck=null}){
  let deck=null;
  if(deckFile){
    if(!readDeck)throw new Error('This window cannot read a presentation.');
    const source=await readDeck(deckFile);
    deck={name:deckFile.name,slides:source.slides.map(s=>({order:s.order,title:s.title,text:s.text,notes:s.notes,clicks:s.clicks??0,hidden:!!s.hidden,titleGuessed:!!s.titleGuessed,...(Number.isInteger(s.sourceId)?{sourceId:s.sourceId}:{})}))};
  }
  const files=[];
  for(const f of pictureFiles)files.push(await prepare(f));
  return {deck,files};
}

// Files taken as they are (a bank, a backup, a saved settings file): their names and bytes, nothing read or shrunk.
export async function readRawFiles(list){
  const files=[];
  for(const f of list)files.push({name:f.name,type:'',mime:f.type||'',bytes:new Uint8Array(await f.arrayBuffer())});
  return files;
}

// The plain-browser version (the preview): ordinary file inputs, one prompt after another.
export function pickWithBrowserInputs(kind,{document=globalThis.document,prepare=prepareImage,readDeck=null,accept='',multiple=false}={}){
  const ask=(accept,multiple)=>new Promise(resolve=>{
    const input=document.createElement('input');input.type='file';input.accept=accept;input.multiple=!!multiple;
    input.onchange=()=>resolve(Array.from(input.files||[]));
    input.oncancel=()=>resolve([]);
    input.click();
  });
  return (async()=>{
    if(kind==='files'){
      const chosen=await ask(accept,multiple);
      if(!chosen.length)return {cancelled:true};
      return {deck:null,files:await readRawFiles(chosen)};
    }
    if(kind==='slides'){
      const [deckFile]=await ask('.pptx',false);
      if(!deckFile)return {cancelled:true};
      const pictureFiles=await ask('image/png,image/jpeg',true);
      return readChosen({deckFile,pictureFiles,prepare,readDeck});
    }
    const [file]=await ask('image/png,image/jpeg,image/gif',false);
    if(!file)return {cancelled:true};
    return readChosen({pictureFiles:[file],prepare});
  })();
}
