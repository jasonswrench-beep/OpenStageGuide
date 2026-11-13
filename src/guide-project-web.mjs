// The browser's side of a project file: it asks the program's own server (scripts/project-server.mjs) to read and write the one project
// that program was started with. The page never sees a file name or a folder; it only says "read the project" and "store the project".
const call=(...args)=>globalThis.fetch(...args);

export function webProjectStore({fetchFn=call,path='/project'}={}){
  return {
    // The project's text, or null when no file has been saved yet.
    async read(){
      const r=await fetchFn(path,{headers:{Accept:'text/plain'},cache:'no-store'});
      if(r.status===204||r.status===404)return null;
      if(!r.ok)throw new Error(`the program answered ${r.status}`);
      return r.text();
    },
    async write(text){
      const r=await fetchFn(path,{method:'PUT',headers:{'Content-Type':'text/plain; charset=utf-8'},body:text,cache:'no-store'});
      if(r.ok)return;
      let why='';
      try{why=(await r.text()).slice(0,200);}catch{why='';}
      throw new Error(why||`the program answered ${r.status}`);
    }
  };
}
