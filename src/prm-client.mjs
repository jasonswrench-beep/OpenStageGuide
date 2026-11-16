// Prism Reference Manager's local API, as far as the design desk needs it: to cite a learning design model's source in the style the author's guide uses. Prism Reference Manager
// (PRM) runs on this computer, off until the person turns on "Allow local connections" in its settings, and answers only the programs on this computer that the person has
// allowed. Its API is documented in its docs/local-api.md and, for the shapes below, read from its code (api.rs, dispatch.rs, capture.rs, library/cite.rs, model.rs; PRM 0.1.3):
//   * the port is fixed (23119 unless the person changed it), with no way to find it out, so the program keeps the port the person gives;
//   * a program asks to be paired (POST /v1/pair), the person chooses Allow in PRM's own window, and the program is then given a token once (GET /v1/pair/<id>);
//   * every other call is POST /v1/call {"method","args"} with `Authorization: Bearer <token>`; an answer is {"result"} or {"error"};
//   * the calls used: capture.check (is it in the library?), refs.add (put it there), cite.render (the reference-list entry and the in-text citation, in a style).
// A browser page cannot make these calls (PRM allows no web origin), so the program's own server makes them; the token never reaches the page. Nothing is sent but the
// details of a published work, and only to this computer. Without PRM every part of the design desk still works: the citation is then the program's own APA 7 text.
export const PRM_DEFAULT_PORT=23119;
export const PRM_SCOPES=Object.freeze(['read','format','write','import']);
export const PRM_NAME='OpenStage Guide';

export class PrmError extends Error{
  constructor(message,code='error'){super(message);this.name='PrmError';this.code=code;}
}
export const validPort=p=>Number.isInteger(p)&&p>=1024&&p<=65535;
export function cleanPort(value){
  if(value===undefined||value===null||value==='')return PRM_DEFAULT_PORT;
  const n=Number(String(value).trim());
  if(!validPort(n))throw new PrmError('The port is a whole number from 1024 to 65535. Prism Reference Manager shows it in Settings, under Local connections.','port');
  return n;
}

async function post({fetch,port,path,body,token,timeoutMs=30000}){
  const headers={'Content-Type':'application/json',...(token?{Authorization:`Bearer ${token}`}:{})};
  let res;
  try{res=await fetch(`http://127.0.0.1:${port}${path}`,{method:'POST',headers,body:JSON.stringify(body??{}),signal:AbortSignal.timeout(timeoutMs)});}
  catch{throw new PrmError('Prism Reference Manager did not answer. It must be running, with "Allow local connections" turned on in its Settings, and on this port.','unreachable');}
  return read(res);
}
async function get({fetch,port,path,timeoutMs=15000}){
  let res;
  try{res=await fetch(`http://127.0.0.1:${port}${path}`,{method:'GET',signal:AbortSignal.timeout(timeoutMs)});}
  catch{throw new PrmError('Prism Reference Manager did not answer. It must be running, with "Allow local connections" turned on in its Settings, and on this port.','unreachable');}
  return read(res);
}
async function read(res){
  const text=await res.text();
  let data={};
  try{data=text?JSON.parse(text):{};}catch{data={};}
  if(res.ok)return data;
  const message=typeof data.error==='string'&&data.error?data.error:`Prism Reference Manager answered with an error (${res.status}).`;
  throw new PrmError(message,res.status===401?'unauthorized':res.status===403?'forbidden':res.status===429?'busy':'error');
}

/** Whether something answers on the port as Prism Reference Manager does: an unsigned call is refused with 401 when the API is on, and nothing answers when it is off. */
export async function probe({fetch=globalThis.fetch,port=PRM_DEFAULT_PORT}={}){
  try{await post({fetch,port,path:'/v1/call',body:{method:'library.info',args:{}},timeoutMs:4000});return 'up';}
  catch(e){return e instanceof PrmError&&e.code==='unauthorized'?'up':e instanceof PrmError&&e.code==='unreachable'?'down':'up';}
}
/** Ask to be allowed. Prism Reference Manager shows the request in its window; nothing is given until the person chooses Allow. @returns {Promise<{id:string}>} */
export async function startPairing({fetch=globalThis.fetch,port=PRM_DEFAULT_PORT}={}){
  const out=await post({fetch,port,path:'/v1/pair',body:{name:PRM_NAME,scopes:[...PRM_SCOPES]}});
  if(typeof out.id!=='string'||!out.id)throw new PrmError('Prism Reference Manager did not say how to follow the request. Try again.','pair');
  return {id:out.id};
}
/** Where the request is: pending, approved (with the token, given only once) or denied. */
export async function pollPairing({fetch=globalThis.fetch,port=PRM_DEFAULT_PORT,id}){
  const out=await get({fetch,port,path:`/v1/pair/${encodeURIComponent(id)}`});
  if(out.state==='approved'&&typeof out.token==='string'&&out.token)return {state:'approved',token:out.token};
  if(out.state==='denied')return {state:'denied'};
  return {state:'pending'};
}
/** One call of the API. Gives the result. */
export async function call({fetch=globalThis.fetch,port=PRM_DEFAULT_PORT,token,method,args={}}){
  if(!token)throw new PrmError('This program is not connected to Prism Reference Manager yet.','unauthorized');
  const out=await post({fetch,port,path:'/v1/call',body:{method,args},token});
  return out.result;
}

// ---- a source of the catalogue as a reference of PRM's own shape (not CSL-JSON: refs.add takes the library's Reference)

const TYPES={'article-journal':'article-journal',book:'book',report:'report'};
export function prmReference(s){
  const role=s.editors?'editor':'author';
  const fields={title:s.title};
  if(s.container)fields.container_title=s.container;
  if(s.volume)fields.volume=s.volume;
  if(s.issue)fields.issue=s.issue;
  if(s.pages)fields.page=s.pages;
  if(s.doi)fields.doi=s.doi;
  if(s.publisher)fields.publisher=s.publisher;
  if(s.edition)fields.edition=s.edition;
  return {type:TYPES[s.type]||'book',creators:s.authors.map(a=>({role,name:{family:a.family,given:a.given}})),issued:{year:s.year},fields};
}

/**
 * The citation of a source, from Prism Reference Manager: the work is looked for in the library (by its DOI, then its title) and added when it is not there, and the reference-list
 * entry and the in-text citation are rendered in the style. `added` says whether the work was put in the library by this call.
 * @returns {Promise<{id:string, added:boolean, reference:string, inText:string, style:string}>}
 */
export async function citeSource({fetch=globalThis.fetch,port=PRM_DEFAULT_PORT,token,source,style='apa-7'}){
  const c=args=>call({fetch,port,token,...args});
  const found=await c({method:'capture.check',args:{doi:source.doi||'',title:source.title}});
  let id,added=false;
  if(found?.found&&found.id)id=found.id;
  else{
    const made=await c({method:'refs.add',args:{ref:prmReference(source)}});
    id=made?.id;
    added=true;
  }
  if(!id)throw new PrmError('Prism Reference Manager did not give the work an identity, so it could not be cited.','error');
  const entry=await c({method:'cite.render',args:{request:{mode:'reference',ref_ids:[id],style}}});
  const inText=await c({method:'cite.render',args:{request:{mode:'parenthetical',items:[{ref_id:id}],style}}});
  const text=r=>String(r?.text??'').replace(/\s+/g,' ').trim();
  if(!text(entry)||!text(inText))throw new PrmError('Prism Reference Manager gave no text for the citation.','error');
  return {id,added,reference:text(entry),inText:text(inText),style:String(entry.style_title||entry.style||style)};
}
