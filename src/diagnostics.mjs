const HASH=/^[a-f0-9]{64}$/;
const core=['/src/app.mjs','/src/model.mjs','/src/word.mjs','/src/document-store.mjs','/src/diagnostics.mjs'];
// A module is pinned to a build when it comes from the local server, or, on a public host, from the same https origin as the page that is running it.
export function moduleIdentity(value){try{const u=new URL(value),here=globalThis.location?.origin;const local=['http:','https:'].includes(u.protocol)&&['localhost','127.0.0.1'].includes(u.hostname)&&['3843','3844'].includes(u.port);const samePublicHost=u.protocol==='https:'&&!!here&&u.origin===here&&!['localhost','127.0.0.1'].includes(u.hostname);if(!local&&!samePublicHost)return null;const match=u.pathname.match(/^\/build\/([a-f0-9]{64})(\/src\/[a-z0-9-]+\.mjs)$/);return match?{buildId:match[1],path:match[2]}:null;}catch{return null;}}
export function moduleAsset(path,metaURL){const id=moduleIdentity(metaURL);return id?`/build/${id.buildId}/${path.replace(/^\//,'')}`:'/'+path.replace(/^\//,'');}
const safeCode=path=>/^(?:src\/[a-z0-9-]+\.mjs|scripts\/[a-z0-9-]+\.mjs|package(?:-lock)?\.json|public\/(?:taskpane\.html|help\.html|style\.css|vendor\/jszip\.min\.js))$/.test(path);
function records(values){if(!Array.isArray(values)||values.length>4096)throw new Error('Invalid build record');const seen=new Set();return values.map(v=>{if(!v||typeof v.path!=='string'||v.path.length>1024||v.path.includes('..')||/[\\:\x00-\x1f]/.test(v.path)||!HASH.test(v.sha256)||!Number.isSafeInteger(v.bytes)||v.bytes<0||seen.has(v.path))throw new Error('Invalid build record');seen.add(v.path);return {path:v.path,sha256:v.sha256,bytes:v.bytes};});}
function capabilities(value={}){return {mode:value.mode==='word'?'word':'preview',host:value.host==='Word'?'Word':'unknown',platform:['PC','Mac','OfficeOnline','iOS','Android','Universal'].includes(value.platform)?value.platform:'unknown',hostVersion:typeof value.hostVersion==='string'&&/^\d+(?:\.\d+){0,5}$/.test(value.hostVersion)?value.hostVersion:'unknown',wordApi13:value.wordApi13===true,sharedRuntime11:value.sharedRuntime11===true};}
async function json(fetcher,path){const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),5000);try{const response=await fetcher(path,{cache:'no-store',credentials:'omit',redirect:'error',signal:controller.signal});if(!response.ok)throw new Error('Build unavailable');const declared=Number(response.headers?.get?.('content-length')||0);if(declared>1024*1024)throw new Error('Build record too large');const text=await response.text();if(new TextEncoder().encode(text).length>1024*1024)throw new Error('Build record too large');return JSON.parse(text);}finally{clearTimeout(timer);}}
async function digest(value){return Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(JSON.stringify(value)))),x=>x.toString(16).padStart(2,'0')).join('');}
export async function inspectBuild({entryURL,registry,fetcher=globalThis.fetch,runtime={}}){
  const entry=moduleIdentity(entryURL),modules=[],observed=new Set();let unpinned=false;
  let examined=0;for(const value of registry||[]){if(++examined>4096){unpinned=true;break;}const id=moduleIdentity(value);if(!id||new URL(value).origin!==new URL(entryURL).origin)unpinned=true;else{const key=id.buildId+id.path;if(!observed.has(key)){observed.add(key);modules.push(id);}}}
  const report={format:'openstage-guide-diagnostic',version:1,checkedAt:new Date().toISOString(),loadedBuildId:entry?.buildId||null,servedBuildId:null,loadedRecordValidated:false,status:'unknown',reason:'Loaded module identity is unavailable.',modules,capabilities:capabilities(runtime),buildFiles:[],workspaceBuildVerified:false,nativeManifestVerified:false,scope:'Observed instrumented code identity and local server metadata only. A server match does not prove newest workspace sources are loaded. External host/vendor execution, native manifest and feature acceptance remain unverified.'};
  if(!entry||entry.path!=='/src/app.mjs')return report;
  if(unpinned||modules.length>4096||modules.some(m=>m.buildId!==entry.buildId)){report.status='mixed';report.reason='Loaded modules include an unpinned or different build. Preserve work and reload normally before testing.';return report;}
  if(core.some(path=>!modules.some(m=>m.path===path))){report.reason='The executing core module set has not been fully observed.';return report;}
  try{
    const [archive,current]=await Promise.allSettled([json(fetcher,`/build/${entry.buildId}/build-info.json`),json(fetcher,'/health')]);
    const health=current.status==='fulfilled'?current.value:null;
    if(health?.app==='OpenStage Guide'&&HASH.test(health.buildId)){report.servedBuildId=health.buildId;if(health.buildId!==entry.buildId){report.status='stale';report.reason='Loaded code and current server build differ. Preserve work and reload normally before testing.';}}
    if(archive.status!=='fulfilled')throw new Error('Loaded record unavailable');const record=archive.value;
    if(record.schema!==1||record.app!=='OpenStage Guide'||typeof record.version!=='string'||!/^[0-9]+\.[0-9]+\.[0-9]+$/.test(record.version)||record.buildId!==entry.buildId)throw new Error('Invalid build record');
    const inputs=records(record.inputs),served=records(record.served),descriptor={schema:1,app:record.app,version:record.version,inputs};
    if(await digest(descriptor)!==entry.buildId||modules.some(m=>!served.some(f=>f.path===m.path)))throw new Error('Build record mismatch');
    report.loadedRecordValidated=true;report.buildFiles=inputs.filter(f=>safeCode(f.path));
    if(!report.servedBuildId)throw new Error('Invalid server metadata');
    if(health.buildId!==entry.buildId){report.status='stale';report.reason='Loaded code and current server build differ. Preserve work and reload normally before testing.';}
    else{report.status='matched';report.reason='Observed modules and running server snapshot match. Compare this ID with the intended source build before acceptance; newest workspace files and native behavior remain unverified.';}
  }catch{if(report.status!=='stale')report.reason='The loaded snapshot could not be checked against the local server. No current-build acceptance claim can be made.';}
  return report;
}
