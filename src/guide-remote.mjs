// One guide document, used from two pages.
//
// In the program the guide's document lives in the program's own page, together with its saving (guide-project-adapter.mjs) and, later,
// the editor. The tools (the pane of the first version) run in a frame of that page. They use the document through the same interface
// they always used, but each call is carried to the program's page as a message and answered the same way:
//
//   frame: createRemoteAdapter(...)  --  {osg:'call', id, method, args}  -->  page: serveAdapter(guide, ...)
//   frame:                           <--  {osg:'result', id, ok, value|message}  --
//   frame:                           <--  {osg:'event', name:'status'|'selection', ...}  --
//
// Only the methods in REMOTE_METHODS can be called. The document's other methods (importing and exporting its whole state, moving its cursor,
// undoing) are the program's own and cannot be reached from the frame. Nothing is evaluated: a call names a method and carries plain data.
import {makeContext} from './guide-ooxml.mjs';

export const REMOTE_METHODS=Object.freeze([
  'paragraphs','title','setTitle','customCues','setCustomCues','getLook','setLook','slideMemory','setSlideMemory','versionMemory','setVersionMemory',
  'formSettings','setFormSettings','applyLook','setHeaderFooter','cursorIndex','plainLineAfter','go','deleteRange','moveRange','applyEdits',
  'reviewParts','sourcePackage','refreshFields','captureFile','exportPdf','keepPaneOpen','pickFiles','insert','ready','status','flush'
]);

// The ordinary downloads of a page: what the add-in's Word adapter does when Word cannot open a file.
function browserDownload(base64,fileName,type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'){
  const bytes=Uint8Array.from(atob(base64),ch=>ch.charCodeAt(0));
  const url=URL.createObjectURL(new Blob([bytes],{type}));
  const a=document.createElement('a');a.href=url;a.download=fileName;document.body.append(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),60000);
}

// The frame's end. post(message) sends to the program's page; listen(fn) calls fn(data) for every message that came from it (the caller has
// already checked who sent it). onStatus(state) hears how the saving stands.
export function createRemoteAdapter({post,listen,onStatus=()=>{},download=browserDownload,build=makeContext}){
  let counter=0,selection=null,last={state:'new',message:''};
  const pending=new Map();
  listen(data=>{
    if(!data||typeof data!=='object')return;
    if(data.osg==='result'){
      const p=pending.get(data.id);
      if(!p)return;
      pending.delete(data.id);
      if(data.ok)p.resolve(data.value);else p.reject(new Error(String(data.message||'The guide could not do that.')));
    }else if(data.osg==='event'){
      if(data.name==='status'){last={state:String(data.state||''),message:String(data.message||'')};onStatus(last);}
      else if(data.name==='selection')selection?.();
    }
  });
  const call=(method,...args)=>new Promise((resolve,reject)=>{
    const id=++counter;
    pending.set(id,{resolve,reject});
    try{post({osg:'call',id,method,args});}catch(e){pending.delete(id);reject(e);}
  });
  const adapter={kind:'demo'};
  for(const name of REMOTE_METHODS)if(!['insert','ready','status'].includes(name))adapter[name]=(...args)=>call(name,...args);
  // The content is built here, where the tool's own code is, and only the finished Word-format text is sent.
  adapter.insert=(make,opts)=>{
    const body=make(build());
    if(typeof body!=='string')throw new Error('The block could not be made.');
    return call('insert',body,opts);
  };
  adapter.ready=call('ready');
  adapter.ready.catch(()=>{});
  // How the saving stands when the document opened: the program told the frame nothing before it was listening.
  adapter.ready.then(()=>call('status')).then(s=>{if(s&&typeof s==='object'&&s.state){last={state:String(s.state),message:String(s.message||'')};onStatus(last);}},()=>{});
  adapter.status=()=>last;
  adapter.blocked=()=>last.state==='error';
  adapter.onSelectionChanged=handler=>{selection=handler;return true;};
  // The program has no Word window to open a file in: the tools fall back to a download, as they do when Word cannot open a file.
  adapter.openDocument=()=>Promise.reject(new Error('The program has no Word window to open the file in.'));
  adapter.offerDownload=(base64,fileName,type)=>download(base64,fileName,type);
  return adapter;
}

// The program's end. guide is the guide's document (with its saving); post(message) sends to the frame; listen(fn) calls fn(data) for every
// message that came from the frame (the caller has already checked it is the frame). Returns {stop, status(state)}: tell the frame how the
// saving stands with status, and stop listening with stop.
export function serveAdapter(guide,{post,listen,methods=REMOTE_METHODS}){
  const allowed=new Set(methods);
  const stop=listen(async data=>{
    if(!data||typeof data!=='object'||data.osg!=='call'||!Number.isSafeInteger(data.id)||typeof data.method!=='string')return;
    const reply=r=>post({osg:'result',id:data.id,...r});
    if(!allowed.has(data.method)){reply({ok:false,message:'The guide does not offer that.'});return;}
    const args=Array.isArray(data.args)?data.args:[];
    try{
      let value;
      if(data.method==='ready')value=await guide.ready;
      else if(data.method==='status')value=guide.status();
      else if(data.method==='insert'){
        if(typeof args[0]!=='string')throw new Error('The block could not be made.');
        value=await guide.insert(()=>args[0],args[1]);
      }else value=await guide[data.method](...args);
      reply({ok:true,value});
    }catch(e){reply({ok:false,message:String(e?.message||e)});}
  });
  guide.onSelectionChanged?.(()=>post({osg:'event',name:'selection'}));
  return {
    stop:()=>{if(typeof stop==='function')stop();},
    status:s=>post({osg:'event',name:'status',state:s.state,message:s.message||''})
  };
}
