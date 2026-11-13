// Every file box in the pane, made to work in Word. In Word's task pane an <input type="file"> does nothing: no file window opens
// (seen in the author's Word, 2026-11-06 for pictures and 2026-11-07 for the OpenAssess bank). The pane already has a file window of
// its own (picker.html, opened with Office's dialog API; see file-pick.mjs). This file connects the two: when a file box in the pane
// is pressed, the file window opens, and what the author chooses there is put into the same file box and its "change" is announced, so
// the code that reads a file box works exactly as it did before and needs no change.
//
// Not used in the plain-browser preview, where a file box works by itself.

// Files made from what the file window sent back, as the list a file box holds.
export function fileListFrom(entries,{FileCtor=globalThis.File,DataTransferCtor=globalThis.DataTransfer}={}){
  const transfer=new DataTransferCtor();
  for(const e of entries)transfer.items.add(new FileCtor([e.bytes],e.name,{type:e.mime||''}));
  return transfer.files;
}

// The words next to a file box: its <label>, or its own name for people who use a screen reader.
export function labelOfInput(input,document){
  const byFor=input.id&&document?.querySelector?.(`label[for="${input.id}"]`);
  const text=(byFor?.textContent||input.getAttribute?.('aria-label')||'').replace(/\s+/g,' ').trim();
  return text;
}

// Starts listening on the whole page. `pick` opens the file window ({kind:'files', accept, multiple, label} -> {files} or
// {cancelled:true}); `onError` is told about a problem in words. Returns a function that stops listening.
export function installFileInputBridge({document,pick,onError=()=>{},enabled=true,FileCtor,DataTransferCtor,EventCtor=globalThis.Event}){
  if(!enabled||!document?.addEventListener)return ()=>{};
  let open=false;
  const handler=async event=>{
    const input=event.target;
    if(!input||String(input.tagName).toUpperCase()!=='INPUT'||String(input.type).toLowerCase()!=='file')return;
    // The file box's own window never opens in Word, so the click is taken over entirely.
    event.preventDefault();
    event.stopPropagation?.();
    if(open)return; // a window is already open (or opening): a second press does nothing
    open=true;
    try{
      const res=await pick({kind:'files',accept:input.accept||'',multiple:!!input.multiple,label:labelOfInput(input,document)});
      if(!res||res.cancelled||!res.files?.length)return;
      input.files=fileListFrom(res.files,{FileCtor,DataTransferCtor});
      input.dispatchEvent(new EventCtor('input',{bubbles:true}));
      input.dispatchEvent(new EventCtor('change',{bubbles:true}));
    }catch(e){
      onError(e);
    }finally{
      open=false;
    }
  };
  document.addEventListener('click',handler,true);
  return ()=>document.removeEventListener('click',handler,true);
}
