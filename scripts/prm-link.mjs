// The program's connection to Prism Reference Manager, kept on the server side (a web page cannot call it, and the token it is given is never shown to the page).
// One action at a time, from the design desk's Models screen: status, connect (ask to be paired; the person chooses Allow in Prism Reference Manager's window), cite (the
// citation of a source of the catalogue) and forget. The port and the token are kept in a file beside the guides, `OpenStage Guide Reference Manager.json`: what the person
// allowed, and nothing else. A token that Prism Reference Manager no longer knows (the person removed the program there) is dropped, with a message that says so.
import fs from 'node:fs/promises';
import path from 'node:path';
import {SOURCES} from '../src/design-models.mjs';
import {PrmError,cleanPort,probe,startPairing,pollPairing,call,citeSource,PRM_DEFAULT_PORT} from '../src/prm-client.mjs';

export const PRM_FILE='OpenStage Guide Reference Manager.json';

/** @param {{file:string, fetch?:typeof fetch, read?:(f:string)=>Promise<string>, write?:(f:string,t:string)=>Promise<void>, remove?:(f:string)=>Promise<void>}} deps */
export function createPrmLink({file,fetch=globalThis.fetch,read=f=>fs.readFile(f,'utf8'),write=async(f,t)=>{await fs.mkdir(path.dirname(f),{recursive:true});await fs.writeFile(f,t,{encoding:'utf8',mode:0o600});},remove=f=>fs.rm(f,{force:true})}){
  let pairing=null;   // {id, port}: a request that waits for the person's Allow
  let busy=false;
  async function load(){
    try{
      const o=JSON.parse(await read(file));
      if(o&&typeof o.token==='string'&&o.token&&Number.isInteger(o.port))return {port:o.port,token:o.token};
    }catch{/* nothing kept, or not readable: not connected */}
    return null;
  }
  const state=(s,extra={})=>({ok:true,state:s,...extra});
  const fail=(e,extra={})=>({ok:false,message:e instanceof PrmError?e.message:'The connection to Prism Reference Manager failed.',code:e instanceof PrmError?e.code:'error',...extra});

  async function status(){
    if(pairing){
      try{
        const p=await pollPairing({fetch,port:pairing.port,id:pairing.id});
        if(p.state==='approved'){await write(file,JSON.stringify({port:pairing.port,token:p.token}));pairing=null;return state('connected',{message:'Connected. This program may now ask Prism Reference Manager to format citations.'});}
        if(p.state==='denied'){const port=pairing.port;pairing=null;return state('denied',{port,message:'Prism Reference Manager was told not to allow this program. Choose Connect to ask again.'});}
        return state('pending',{port:pairing.port,message:'Waiting for you to choose Allow in Prism Reference Manager\'s window.'});
      }catch(e){pairing=null;return fail(e,{state:'not-connected'});}
    }
    const kept=await load();
    if(kept){
      try{
        await call({fetch,port:kept.port,token:kept.token,method:'library.info',args:{}});
        return state('connected',{port:kept.port});
      }catch(e){
        if(e instanceof PrmError&&e.code==='unauthorized'){await remove(file);return state('not-connected',{port:kept.port,message:'Prism Reference Manager no longer allows this program. Choose Connect to ask again.'});}
        // Reachable, but no library is open (it says so in its own words), or not reachable now: the connection is kept.
        if(e instanceof PrmError&&e.code==='unreachable')return state('connected',{port:kept.port,reachable:false,message:e.message});
        return state('connected',{port:kept.port,message:e.message});
      }
    }
    const up=await probe({fetch,port:PRM_DEFAULT_PORT});
    return state('not-connected',{port:PRM_DEFAULT_PORT,reachable:up==='up'});
  }

  async function connect(body){
    const port=cleanPort(body.port);
    const out=await startPairing({fetch,port});
    pairing={id:out.id,port};
    return state('pending',{port,message:'Prism Reference Manager is asking you. Choose Allow in its window.'});
  }

  async function cite(body){
    const id=String(body.source??'');
    const source=Object.hasOwn(SOURCES,id)?SOURCES[id]:null;
    if(!source)return {ok:false,message:'That source is not one of the program\'s.',code:'source'};
    const kept=await load();
    if(!kept)return {ok:false,message:'This program is not connected to Prism Reference Manager. Choose Connect first.',code:'unauthorized'};
    try{
      const made=await citeSource({fetch,port:kept.port,token:kept.token,source,style:typeof body.style==='string'&&body.style?body.style.slice(0,60):'apa-7'});
      return {ok:true,...made,source:id};
    }catch(e){
      if(e instanceof PrmError&&e.code==='unauthorized')await remove(file);
      throw e;
    }
  }

  return {
    /** One action of the desk's Models screen. @param {{action?:string, port?:unknown, source?:unknown, style?:unknown}} body */
    async handle(body){
      if(busy)return {ok:false,message:'Prism Reference Manager is being asked something already. Wait a moment.',code:'busy'};
      busy=true;
      try{
        switch(body?.action){
          case 'status':return await status();
          case 'connect':return await connect(body);
          case 'cite':return await cite(body);
          case 'forget':pairing=null;await remove(file);return state('not-connected',{message:'This program no longer holds a connection. You can also remove it in Prism Reference Manager\'s settings.'});
          default:return {ok:false,message:'That is not something the connection does.',code:'action'};
        }
      }catch(e){return fail(e);}
      finally{busy=false;}
    },
  };
}
