import path from 'node:path';
import {PANE_CSP,pageHeaders,stripOfficeScript} from './security-headers.mjs';
const types={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.mjs':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.png':'image/png','.svg':'image/svg+xml','.json':'application/json','.txt':'text/plain'};
export function createHandler(snapshot){return async(req,res)=>{
  // Every HTML page says what it may load and that only a page of this server may frame it (scripts/security-headers.mjs).
  const reply=(code,type,data,cache='no-store')=>{res.writeHead(code,{'Content-Type':type,'Cache-Control':cache,'X-Content-Type-Options':'nosniff','Referrer-Policy':'no-referrer',...(/^text\/html/.test(type)?pageHeaders(PANE_CSP):{})});res.end(req.method==='HEAD'?undefined:data);};
  try{
    if(!['GET','HEAD'].includes(req.method)){reply(405,'text/plain','Method not allowed');return;}
    const url=new URL(req.url,'http://localhost'),decoded=decodeURIComponent(url.pathname);
    if(decoded==='/health'){reply(200,'application/json',JSON.stringify({app:'OpenStage Guide',version:snapshot.version,buildId:snapshot.buildId}));return;}
    const prefix='/build/'+snapshot.buildId;
    if(decoded===prefix+'/build-info.json'){reply(200,'application/json',JSON.stringify(snapshot.record));return;}
    if(decoded.startsWith('/build/')&&!decoded.startsWith(prefix+'/')){reply(404,'text/plain','Build unavailable');return;}
    let name=decoded.startsWith(prefix+'/')?decoded.slice(prefix.length):decoded;if(name==='/')name='/taskpane.html';
    if(name.includes('\\')||name.split('/').some(s=>s==='..'||s==='.')||/[\x00-\x1f]/.test(name)){reply(403,'text/plain','Unsupported path');return;}
    const data=snapshot.read(name);if(!data){reply(404,'text/plain','Not found');return;}
    const cache=/\.(png|svg)$/i.test(name)?'public, max-age=86400':'no-store';
    reply(200,types[path.extname(name)]||'application/octet-stream',name.endsWith('.html')?Buffer.from(stripOfficeScript(data)):data,cache);
  }catch{reply(404,'text/plain','Not found');}
};}
