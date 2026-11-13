import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {startProjectServer} from '../scripts/project-server.mjs';
import {serializeProject} from '../src/guide-project.mjs';

// The pane's own files are not what is tested here: a stand-in snapshot answers /health and nothing else.
const snapshot={version:'0-test',buildId:'test-build',record:{},read:()=>null};
const project=title=>serializeProject({title,paragraphs:[{xml:'',text:title,style:'OSG Guide Title'}]},{savedAt:'2026-11-13T00:00:00.000Z'});

async function withServer(fn,{existing,maxBytes}={}){
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-project-server-'));
  const file=path.join(dir,'My guide.json');
  if(existing!==undefined)fs.writeFileSync(file,existing);
  const s=await startProjectServer({file,port:0,snapshot,...(maxBytes?{maxBytes}:{})});
  const call=(method,urlPath,{headers={},body,host=`127.0.0.1:${s.port}`}={})=>new Promise((resolve,reject)=>{
    const req=http.request({host:'127.0.0.1',port:s.port,path:urlPath,method,headers:{Host:host,...headers}},res=>{let d='';res.on('data',c=>d+=c);res.on('end',()=>resolve({status:res.statusCode,text:d,headers:res.headers}));});
    req.on('error',reject);if(body!==undefined)req.write(body);req.end();
  });
  try{await fn({file,dir,call,port:s.port});}finally{await s.close();fs.rmSync(dir,{recursive:true,force:true});}
}
const put=(call,body,headers={})=>call('PUT','/project',{headers:{'Content-Type':'text/plain; charset=utf-8',...headers},body});

test('a project with no file yet answers "nothing there", and the first save makes the file',async()=>{
  await withServer(async({file,dir,call})=>{
    assert.equal((await call('GET','/project')).status,204);
    assert.equal(fs.existsSync(file),false,'asking does not make a file');
    const text=project('First');
    const saved=await put(call,text);
    assert.equal(saved.status,200);
    assert.equal(fs.readFileSync(file,'utf8'),text);
    const back=await call('GET','/project');
    assert.equal(back.status,200);assert.equal(back.text,text);
    assert.deepEqual(fs.readdirSync(dir),['My guide.json'],'no temporary file and no backup are left for a new project');
  });
});

test('the project as it was found is kept once as .bak, on the first save only',async()=>{
  const found=project('As found');
  await withServer(async({file,call})=>{
    await put(call,project('Second'));
    assert.equal(fs.readFileSync(file+'.bak','utf8'),found);
    await put(call,project('Third'));
    assert.equal(fs.readFileSync(file+'.bak','utf8'),found,'later saves do not replace the backup');
    assert.equal(fs.readFileSync(file,'utf8'),project('Third'));
  },{existing:found});
});

test('something that is not a project is refused with the reason, and the file is not touched',async()=>{
  const found=project('Keep');
  await withServer(async({file,call})=>{
    for(const [body,pattern] of [['{"nope":1}',/not an OpenStage Guide project/],['not json',/not readable/],[JSON.stringify({format:'openstage-guide-project',version:99,guide:{paragraphs:[]}}),/newer/]]){
      const r=await put(call,body);
      assert.equal(r.status,400);assert.match(r.text,pattern);
    }
    assert.equal(fs.readFileSync(file,'utf8'),found);
    assert.equal(fs.existsSync(file+'.bak'),false,'a refused save does not even take the backup');
  },{existing:found});
});

test('only this page on this address may read or change the project',async()=>{
  await withServer(async({file,call,port})=>{
    const text=project('Mine');
    assert.equal((await call('PUT','/project',{host:'evil.example',headers:{'Content-Type':'text/plain'},body:text})).status,403,'another host name (DNS rebinding)');
    assert.equal((await call('GET','/project',{host:'evil.example:'+port})).status,403,'reading from another host name');
    assert.equal((await put(call,text,{Origin:'http://evil.example'})).status,403,'another web site');
    assert.equal((await put(call,text,{'Sec-Fetch-Site':'cross-site'})).status,403,'a cross-site fetch');
    assert.equal((await put(call,text,{'Sec-Fetch-Site':'same-site'})).status,403,'a sibling site');
    assert.equal((await call('PUT','/project',{headers:{'Content-Type':'application/x-www-form-urlencoded'},body:text})).status,415,'a form post');
    assert.equal((await call('POST','/project',{body:text})).status,405);
    assert.equal((await call('DELETE','/project')).status,405);
    assert.equal(fs.existsSync(file),false,'nothing refused made a file');
    // The page's own address, and no Origin or a same-origin one, are fine.
    assert.equal((await put(call,text,{Origin:`http://127.0.0.1:${port}`,'Sec-Fetch-Site':'same-origin'})).status,200);
    assert.equal((await put(call,text,{Host:`localhost:${port}`})).status,200,'localhost is the same computer');
  });
});

test('a project larger than the limit is refused and nothing is stored',async()=>{
  await withServer(async({file,call})=>{
    const r=await put(call,'x'.repeat(5000));
    assert.equal(r.status,413);
    assert.equal(fs.existsSync(file),false);
  },{maxBytes:1000});
});

test('everything else is the pane: /health answers, and the project file is not reachable by its name or a path',async()=>{
  await withServer(async({call})=>{
    const h=await call('GET','/health');
    assert.equal(h.status,200);assert.match(h.text,/OpenStage Guide/);
    assert.equal((await call('GET','/My%20guide.json')).status,404);
    assert.notEqual((await call('GET','/../package.json')).status,200);
    assert.equal((await call('GET','/project/other')).status,404);
  });
});
