// The browser's side of a project file: it asks the program's own server (scripts/project-server.mjs) to read and write the one project
// that program was started with. The page never sees a file name or a folder; it only says "read the project" and "store the project".
const call=(...args)=>globalThis.fetch(...args);

// `token` gives the token the program handed out when this guide was opened or made. It goes with every read and save, so the program can tell this
// page's guide from the one that is open now (a page left over from another tab or window is refused instead of writing over the other guide).
export function webProjectStore({fetchFn=call,path='/project',token=()=>null}={}){
  const headers=extra=>{const t=token();return t?{...extra,'X-Guide-Token':String(t)}:extra;};
  // The program says a page is out of date with a 409 and a code. Such a page must stop saving, so its error says so.
  async function stale(r){
    if(r.status!==409)return null;
    let body=null;
    try{body=JSON.parse(await r.clone().text());}catch{body=null;}
    return body&&body.code==='stale-guide'?Object.assign(new Error(String(body.error||'This guide was closed in another window.')),{stale:true}):null;
  }
  return {
    // The project's text, or null when no file has been saved yet.
    async read(){
      const r=await fetchFn(path,{headers:headers({Accept:'text/plain'}),cache:'no-store'});
      const old=await stale(r);
      if(old)throw old;
      if(r.status===204||r.status===404)return null;
      if(!r.ok)throw new Error(`the program answered ${r.status}`);
      return r.text();
    },
    async write(text){
      const r=await fetchFn(path,{method:'PUT',headers:headers({'Content-Type':'text/plain; charset=utf-8'}),body:text,cache:'no-store'});
      if(r.ok)return;
      const old=await stale(r);
      if(old)throw old;
      let why='';
      try{why=(await r.text()).slice(0,200);}catch{why='';}
      throw new Error(why||`the program answered ${r.status}`);
    }
  };
}
