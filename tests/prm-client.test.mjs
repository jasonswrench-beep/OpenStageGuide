// The client of Prism Reference Manager's local API, against a stand-in that answers as its documentation and code say (docs/local-api.md, api.rs, dispatch.rs).
// The real thing is checked by a person (TESTING-TODO.md): it needs Prism Reference Manager running, with its local connections on, and someone to choose Allow.
import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {probe,startPairing,pollPairing,call,citeSource,prmReference,cleanPort,PrmError,PRM_DEFAULT_PORT,PRM_SCOPES} from '../src/prm-client.mjs';
import {createPrmLink} from '../scripts/prm-link.mjs';
import {SOURCES} from '../src/design-models.mjs';

// A stand-in for Prism Reference Manager. `log` holds every request it got; `library` the works it holds.
async function standIn({approveAfter=2,library=[],deny=false}={}){
  const log=[],lib=[...library];
  let polls=0;
  const server=http.createServer((req,res)=>{
    let body='';
    req.on('data',c=>body+=c);
    req.on('end',()=>{
      log.push({method:req.method,url:req.url,headers:req.headers,body:body?JSON.parse(body):null});
      const send=(code,data)=>{res.writeHead(code,{'Content-Type':'application/json'});res.end(JSON.stringify(data));};
      if(req.headers.origin)return send(403,{error:'Only programs on this computer may ask.'});
      if(req.method==='POST'&&req.url==='/v1/pair')return send(202,{id:'p1',state:'pending'});
      if(req.method==='GET'&&req.url==='/v1/pair/p1'){
        polls++;
        if(deny)return send(200,{state:'denied'});
        return send(200,polls>approveAfter?{state:'approved',token:'prm_secret'}:{state:'pending'});
      }
      if(req.method==='POST'&&req.url==='/v1/call'){
        if(req.headers.authorization!=='Bearer prm_secret')return send(401,{error:'Pair this program first.'});
        const {method,args}=JSON.parse(body);
        if(method==='library.info')return send(200,{result:{name:'My library',counts:{references:lib.length}}});
        if(method==='capture.check'){
          const hit=lib.find(r=>(args.doi&&r.fields.doi===args.doi)||r.fields.title===args.title);
          return send(200,{result:hit?{found:true,by:'DOI',id:hit.id,title:hit.fields.title}:{found:false}});
        }
        if(method==='refs.add'){const r={...args.ref,id:`r${lib.length+1}`};lib.push(r);return send(200,{result:r});}
        if(method==='cite.render'){
          const r=lib.find(x=>x.id===(args.request.ref_ids?.[0]||args.request.items?.[0]?.ref_id));
          if(!r)return send(400,{error:'Choose at least one source.'});
          const who=r.creators.map(c=>c.name.family).join(' & ');
          return send(200,{result:{text:args.request.mode==='reference'?`${who} (${r.issued.year}). ${r.fields.title}.`:`(${who}, ${r.issued.year})`,style_title:'American Psychological Association 7th edition',style:args.request.style}});
        }
        return send(400,{error:'Unknown method.'});
      }
      return send(404,{error:'Unknown endpoint.'});
    });
  });
  await new Promise(r=>server.listen(0,'127.0.0.1',r));
  return {port:server.address().port,log,lib,close:()=>new Promise(r=>server.close(r))};
}

test('a port is a whole number from 1024 up, 23119 when none is given, and a reason when it is not', () => {
  assert.equal(cleanPort(''),PRM_DEFAULT_PORT);
  assert.equal(cleanPort(undefined),23119);
  assert.equal(cleanPort(' 24000 '),24000);
  for(const bad of ['80','abc','70000','1.5'])assert.throws(()=>cleanPort(bad),/whole number from 1024 to 65535/,bad);
});

test('a work of the catalogue is a reference of the library\'s own shape, authors or editors, with only what it has', () => {
  const r=prmReference(SOURCES.merrill2002);
  assert.deepEqual(r,{type:'article-journal',creators:[{role:'author',name:{family:'Merrill',given:'M. D.'}}],issued:{year:2002},
    fields:{title:'First principles of instruction',container_title:'Educational Technology Research and Development',volume:'50',issue:'3',page:'43–59',doi:'10.1007/BF02505024'}});
  const book=prmReference(SOURCES.anderson2001);
  assert.equal(book.type,'book');
  assert.deepEqual(book.creators.map(c=>c.role),['editor','editor']);
  assert.equal(book.fields.publisher,'Longman');
  assert.ok(!('doi' in book.fields)&&!('container_title' in book.fields));
  assert.equal(prmReference(SOURCES.gagne1985).fields.edition,'4th ed.');
});

test('with nothing on the port it says so, and with the API on it says it is there, signed in or not', async () => {
  const ps=await standIn();
  try{
    assert.equal(await probe({port:ps.port}),'up');
    assert.equal(ps.log[0].headers.authorization,undefined,'the probe is not signed');
  }finally{await ps.close();}
  assert.equal(await probe({port:ps.port}),'down');
  await assert.rejects(()=>startPairing({port:ps.port}),e=>e instanceof PrmError&&e.code==='unreachable'&&/must be running, with "Allow local connections" turned on/.test(e.message));
});

test('pairing asks with the program\'s name and the four permissions, waits for Allow, and is given the token once', async () => {
  const ps=await standIn({approveAfter:2});
  try{
    const {id}=await startPairing({port:ps.port});
    assert.equal(id,'p1');
    assert.deepEqual(ps.log[0].body,{name:'OpenStage Guide',scopes:['read','format','write','import']});
    assert.deepEqual(PRM_SCOPES,['read','format','write','import']);
    assert.equal(ps.log[0].headers['content-type'],'application/json');
    assert.deepEqual(await pollPairing({port:ps.port,id}),{state:'pending'});
    assert.deepEqual(await pollPairing({port:ps.port,id}),{state:'pending'});
    assert.deepEqual(await pollPairing({port:ps.port,id}),{state:'approved',token:'prm_secret'});
  }finally{await ps.close();}
  const denied=await standIn({deny:true});
  try{assert.deepEqual(await pollPairing({port:denied.port,id:'p1'}),{state:'denied'});}finally{await denied.close();}
});

test('a call is signed with the token, and one without it or with the wrong one is refused in Prism Reference Manager\'s words', async () => {
  const ps=await standIn();
  try{
    assert.deepEqual(await call({port:ps.port,token:'prm_secret',method:'library.info'}),{name:'My library',counts:{references:0}});
    assert.equal(ps.log.at(-1).headers.authorization,'Bearer prm_secret');
    assert.equal(ps.log.at(-1).headers.origin,undefined,'no web origin: Prism Reference Manager turns those away');
    await assert.rejects(()=>call({port:ps.port,token:'nope',method:'library.info'}),e=>e.code==='unauthorized'&&e.message==='Pair this program first.');
    await assert.rejects(()=>call({port:ps.port,method:'library.info'}),/not connected to Prism Reference Manager/);
  }finally{await ps.close();}
});

test('a source is added to the library when it is not there, cited as a reference-list entry and in the text, and not added twice', async () => {
  const ps=await standIn();
  try{
    const first=await citeSource({port:ps.port,token:'prm_secret',source:SOURCES.merrill2002});
    assert.deepEqual(first,{id:'r1',added:true,reference:'Merrill (2002). First principles of instruction.',inText:'(Merrill, 2002)',style:'American Psychological Association 7th edition'});
    const methods=ps.log.filter(l=>l.url==='/v1/call').map(l=>l.body.method);
    assert.deepEqual(methods,['capture.check','refs.add','cite.render','cite.render']);
    assert.deepEqual(ps.log.filter(l=>l.body?.method==='cite.render').map(l=>l.body.args.request.mode),['reference','parenthetical']);
    assert.deepEqual(ps.log.find(l=>l.body?.method==='capture.check').body.args,{doi:'10.1007/BF02505024',title:'First principles of instruction'});
    const again=await citeSource({port:ps.port,token:'prm_secret',source:SOURCES.merrill2002,style:'apa-7'});
    assert.equal(again.added,false);
    assert.equal(ps.lib.length,1,'the library holds it once');
    assert.equal(ps.log.filter(l=>l.body?.method==='refs.add').length,1);
  }finally{await ps.close();}
});

test('the connection kept for the design desk: ask, wait, allowed, cite, forget; the token stays in a file of its own, never in an answer', async () => {
  const ps=await standIn({approveAfter:1});
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-prm-'));
  const file=path.join(dir,'conn.json');
  try{
    const link=createPrmLink({file});
    let s=await link.handle({action:'status'});
    assert.deepEqual([s.ok,s.state],[true,'not-connected']);
    s=await link.handle({action:'connect',port:String(ps.port)});
    assert.deepEqual([s.ok,s.state,s.port],[true,'pending',ps.port]);
    s=await link.handle({action:'status'});
    assert.equal(s.state,'pending');
    s=await link.handle({action:'status'});
    assert.equal(s.state,'connected');
    assert.ok(!JSON.stringify(s).includes('prm_secret'),'the token is not in what the page is told');
    assert.deepEqual(JSON.parse(fs.readFileSync(file,'utf8')),{port:ps.port,token:'prm_secret'});
    s=await link.handle({action:'status'});
    assert.equal(s.state,'connected');
    const c=await link.handle({action:'cite',source:'merrill2002'});
    assert.deepEqual([c.ok,c.reference,c.inText,c.source],[true,'Merrill (2002). First principles of instruction.','(Merrill, 2002)','merrill2002']);
    assert.ok(!JSON.stringify(c).includes('prm_secret'));
    const bad=await link.handle({action:'cite',source:'nothing'});
    assert.deepEqual([bad.ok,bad.code],[false,'source']);
    assert.equal((await link.handle({action:'cite',source:'../x'})).code,'source','only a source of the catalogue');
    const gone=await link.handle({action:'forget'});
    assert.equal(gone.state,'not-connected');
    assert.equal(fs.existsSync(file),false);
    const unpaired=await link.handle({action:'cite',source:'merrill2002'});
    assert.deepEqual([unpaired.ok,unpaired.code],[false,'unauthorized']);
    assert.equal((await link.handle({action:'wat'})).code,'action');
  }finally{await ps.close();fs.rmSync(dir,{recursive:true,force:true});}
});

test('a token that Prism Reference Manager no longer knows is dropped, and the person is told; a refusal to pair is said', async () => {
  const ps=await standIn();
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-prm-'));
  const file=path.join(dir,'conn.json');
  try{
    fs.writeFileSync(file,JSON.stringify({port:ps.port,token:'prm_old'}));
    const link=createPrmLink({file});
    const s=await link.handle({action:'status'});
    assert.deepEqual([s.ok,s.state],[true,'not-connected']);
    assert.match(s.message,/no longer allows this program/);
    assert.equal(fs.existsSync(file),false);
  }finally{await ps.close();}
  const denied=await standIn({deny:true});
  try{
    const link=createPrmLink({file});
    await link.handle({action:'connect',port:String(denied.port)});
    const s=await link.handle({action:'status'});
    assert.equal(s.state,'denied');
    assert.equal(fs.existsSync(file),false);
    const off=await link.handle({action:'connect',port:'80'});
    assert.deepEqual([off.ok,off.code],[false,'port']);
  }finally{await denied.close();fs.rmSync(dir,{recursive:true,force:true});}
  const nobody=createPrmLink({file:path.join(os.tmpdir(),'osg-none.json')});
  const down=await nobody.handle({action:'connect',port:'39999'});
  assert.deepEqual([down.ok,down.code],[false,'unreachable']);
});

// ---- the server's route

import {startProjectServer} from '../scripts/project-server.mjs';
function post(port,urlPath,body){
  return new Promise((resolve,reject)=>{
    const data=JSON.stringify(body);
    const req=http.request({host:'127.0.0.1',port,path:urlPath,method:'POST',headers:{Host:`127.0.0.1:${port}`,'Content-Type':'application/json','Content-Length':Buffer.byteLength(data)}},res=>{
      const chunks=[];res.on('data',c=>chunks.push(c));res.on('end',()=>resolve({status:res.statusCode,body:Buffer.concat(chunks).toString('utf8')}));
    });
    req.on('error',reject);req.write(data);req.end();
  });
}
test('POST /api/prm answers the desk\'s Reference Manager actions, with the token kept in the data folder and out of the answer',async()=>{
  const ps=await standIn({approveAfter:0});
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-prm-server-'));
  const server=await startProjectServer({folder:path.join(dir,'guides'),port:0,snapshot:{version:'0-test',buildId:'test-build',record:{},read:()=>null},dataFolder:path.join(dir,'data')});
  try{
    const none=await post(server.port,'/api/prm',{action:'status'});
    assert.equal(none.status,200);
    assert.equal(JSON.parse(none.body).state,'not-connected');
    const asked=JSON.parse((await post(server.port,'/api/prm',{action:'connect',port:String(ps.port)})).body);
    assert.equal(asked.state,'pending');
    const done=await post(server.port,'/api/prm',{action:'status'});
    assert.equal(JSON.parse(done.body).state,'connected');
    assert.ok(!done.body.includes('prm_secret'));
    const file=path.join(dir,'data','OpenStage Guide Reference Manager.json');
    assert.equal(JSON.parse(fs.readFileSync(file,'utf8')).token,'prm_secret');
    assert.equal(fs.existsSync(path.join(dir,'guides','OpenStage Guide Reference Manager.json')),false,'not among the guides');
    const cited=JSON.parse((await post(server.port,'/api/prm',{action:'cite',source:'merrill2002'})).body);
    assert.deepEqual([cited.ok,cited.inText],[true,'(Merrill, 2002)']);
    assert.equal((await post(server.port,'/api/prm',{action:'forget'})).status,200);
    assert.equal(fs.existsSync(file),false);
  }finally{await server.close();await ps.close();fs.rmSync(dir,{recursive:true,force:true});}
});
