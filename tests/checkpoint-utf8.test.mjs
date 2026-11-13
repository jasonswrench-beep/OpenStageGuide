import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import JSZip from 'jszip';
import {createCheckpoint,readCheckpoint,sha256} from '../src/checkpoint.mjs';
import {masterFixture,installParser,assets} from './helpers/checkpoint-fixture.mjs';
import {xml} from '../src/worksheet.mjs';
installParser();
const encoder=new TextEncoder(),decoder=new TextDecoder('utf-8',{fatal:true});
const browserContext=vm.createContext({setTimeout,clearTimeout,setImmediate,clearImmediate,Promise,Uint8Array,ArrayBuffer,DataView,TextEncoder,TextDecoder,console,Buffer:undefined});
vm.runInContext(fs.readFileSync(new URL('../public/vendor/jszip.min.js',import.meta.url),'utf8'),browserContext,{filename:'browser-jszip.min.js'});
const BrowserZip=browserContext.JSZip;
async function titledFixture(title){
 const f=await masterFixture(),before=f.project.title,z=await JSZip.loadAsync(f.master);f.project.title=title;
 for(const name of ['word/document.xml','word/webextensions/webextension1.xml'])z.file(name,encoder.encode((await z.file(name).async('string')).split(xml(before)).join(xml(title))),{binary:true});
 return {project:f.project,master:await z.generateAsync({type:'uint8array',compression:'DEFLATE'})};
}
async function titlePrefix(){
 const f=await masterFixture(),b=await createCheckpoint(f.master,f.project,JSZip,assets);
 return JSON.stringify({...b.manifest,course:{...b.manifest.course,title:'MARKER'}},null,2).indexOf('MARKER');
}
async function verify(backup,fixture,Zip){
 const loaded=await JSZip.loadAsync(backup.data,{checkCRC32:true}),raw=await loaded.file('manifest.json').async('uint8array'),expected=encoder.encode(JSON.stringify(backup.manifest,null,2));
 assert.deepEqual(raw,expected);assert.equal(decoder.decode(raw),JSON.stringify(backup.manifest,null,2));
 const checked=await readCheckpoint(backup.data,Zip);assert.deepEqual(checked.master,fixture.master);assert.deepEqual(checked.project,fixture.project);
 for(const item of backup.manifest.files){const data=await loaded.file(item.path).async('uint8array');assert.equal(data.length,item.bytes);assert.equal(await sha256(data),item.sha256);if(item.path==='master.docx')assert.deepEqual(data,fixture.master);else if(item.path!=='assets/NOTICE.txt')assert.deepEqual(data,new Uint8Array([137,80,78,71,4,5]));}
 return loaded;
}
function hookedZip(Base,hooks={}){
 function Wrapped(){const z=new Base(),file=z.file.bind(z),generate=z.generateAsync.bind(z);z.file=(name,data,...args)=>{if(data===undefined)return file(name);hooks.file?.(name,data,z);return file(name,data,...args);};z.generateAsync=async options=>{await hooks.beforeGenerate?.(z,file);const bytes=await generate(options);hooks.generated?.(bytes);return bytes;};return z;}
 Wrapped.loadAsync=async(...args)=>{await hooks.beforeLoad?.(...args);const z=await Base.loadAsync(...args);await hooks.afterLoad?.(z);return z;};return Wrapped;
}

test('new checkpoint manifests preserve Unicode at16/32Ki UTF16 boundaries in Node and shipped browser ZIP paths',async()=>{
 const prefix=await titlePrefix();assert.ok(prefix>0&&prefix<16383);
 for(const Zip of [JSZip,BrowserZip])for(const boundary of [16384,32768])for(const offset of [-1,0,1])for(const char of ['\u{1f642}','\u{1d11e}','\u00e9']){
  const title='A'.repeat(boundary-1+offset-prefix)+char+' CAF\u00c9 \u4e2d\u6587';const f=await titledFixture(title),before=JSON.stringify(f.project),masterHash=await sha256(f.master),b=await createCheckpoint(f.master,f.project,Zip,assets);
  assert.equal(JSON.stringify(b.manifest,null,2).slice(boundary-1+offset,boundary-1+offset+char.length),char);await verify(b,f,Zip);assert.equal(JSON.stringify(f.project),before);assert.equal(await sha256(f.master),masterHash);
 }
});

test('new checkpoint writer leaves the legacy browser-readable invalidUTF8 manifest recovery path intact',async()=>{
 const prefix=await titlePrefix(),f=await titledFixture('A'.repeat(16383-prefix)+'\u{1f642}'+'B'),b=await createCheckpoint(f.master,f.project,BrowserZip,assets),loaded=await BrowserZip.loadAsync(b.data,{checkCRC32:true}),legacy=new BrowserZip();
 for(const file of Object.values(loaded.files).filter(f=>!f.dir))legacy.file(file.name,file.name==='manifest.json'?JSON.stringify(b.manifest,null,2):await file.async('uint8array'));
 const legacyBytes=await legacy.generateAsync({type:'uint8array',compression:'DEFLATE'}),before=new Uint8Array(legacyBytes),bad=await BrowserZip.loadAsync(legacyBytes,{checkCRC32:true}),raw=await bad.file('manifest.json').async('uint8array');
 assert.throws(()=>decoder.decode(raw));assert.equal(await bad.file('manifest.json').async('string'),JSON.stringify(b.manifest,null,2));
 const recovered=await readCheckpoint(legacyBytes,BrowserZip);assert.deepEqual(recovered.project,f.project);assert.deepEqual(recovered.master,f.master);assert.deepEqual(legacyBytes,before);
});

test('new checkpoint writer refuses altered canonical or malformed manifest bytes before offer',async()=>{
 for(const defect of ['valid-json-extra','invalid-utf8','missing']){
  const f=await masterFixture(),Zip=hookedZip(BrowserZip,{beforeGenerate:async(z,file)=>{const part=z.file('manifest.json'),raw=await part.async('uint8array');if(defect==='missing')z.remove('manifest.json');else if(defect==='valid-json-extra')file('manifest.json',encoder.encode(decoder.decode(raw)+' '));else {raw[0]=0xff;file('manifest.json',raw);}}});
  await assert.rejects(createCheckpoint(f.master,f.project,Zip,assets),/manifest bytes changed/);
 }
});

test('checkpoint canonical expectations and offered bytes cannot alias retained ZIP buffers',async()=>{
 for(const Base of [JSZip,BrowserZip]){
  const f=await masterFixture(),inputs=[],original=new Uint8Array(f.master);let generated,mutated=false;
  const Zip=hookedZip(Base,{file:(name,data)=>{if(data instanceof Uint8Array)inputs.push({name,data});},generated:data=>{generated=data;},beforeLoad:()=>{if(generated&&!mutated){mutated=true;generated.fill(0);for(const x of inputs)x.data.fill(0);}}});
  const b=await createCheckpoint(f.master,f.project,Zip,assets);assert.equal(mutated,true);await verify(b,f,Base);assert.deepEqual(f.master,original);
 }
});

test('new checkpoint writer refuses source/project races while retaining last captured inputs unchanged on ordinary errors',async()=>{
 for(const master of [null,undefined,0,80000000,[],{byteLength:67108865}])await assert.rejects(createCheckpoint(master,{},JSZip,assets),/master is empty or exceeds/);
 for(const target of ['master','project']){
  const f=await masterFixture(),savedMaster=new Uint8Array(f.master),savedProject=structuredClone(f.project);let changed=false;
  const Zip=hookedZip(JSZip,{afterLoad:()=>{if(!changed){changed=true;if(target==='master')f.master[f.master.length-1]^=1;else f.project.title+=' changed';}}});
  await assert.rejects(createCheckpoint(f.master,f.project,Zip,assets),target==='master'?/inputs changed/:/differs|inputs changed/);
  if(target==='project')assert.deepEqual(f.master,savedMaster);else assert.deepEqual(f.project,savedProject);
 }
 const f=await masterFixture(),before=new Uint8Array(f.master),p=JSON.stringify(f.project);await assert.rejects(createCheckpoint(f.master,f.project,JSZip,async()=>({ok:false})),/asset/);assert.deepEqual(f.master,before);assert.equal(JSON.stringify(f.project),p);
});
