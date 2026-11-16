// The PDF the program makes (File > Save as PDF, and the Review screen's PDF button): scripts/pdf-maker.mjs prints the guide's web page with Edge, the server answers
// POST /api/pdf with it, and the page gives it to the author. Edge is replaced by a stand-in in most of these tests; one test uses the real Edge when this computer has it.
import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {EventEmitter} from 'node:events';
import {makePdf,findEdge,pdfIsTagged,pdfHasOutline,PdfError,PDF_MAX_HTML} from '../scripts/pdf-maker.mjs';
import {startProjectServer} from '../scripts/project-server.mjs';
import {createProjectAdapter,NO_PDF} from '../src/guide-project-adapter.mjs';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';

const TAGGED='%PDF-1.4\n1 0 obj << /Type /Catalog /MarkInfo << /Marked true >> /StructTreeRoot 2 0 R /Lang (en) >> endobj\n2 0 obj << /Type /StructTreeRoot >> endobj\n'+'x'.repeat(80);
const UNTAGGED='%PDF-1.4\n1 0 obj << /Type /Catalog >> endobj\n'+'x'.repeat(80);
const PAGE='<!doctype html><html lang="en"><head><title>T</title></head><body><main><h1>Words</h1></main></body></html>';

// A stand-in for Edge: it reads where the PDF should go from its options, writes `pdf` there, and exits.
function fakeRun(pdf,{exit=true,seen=[]}={}){
  return (exe,args,options)=>{
    seen.push({exe,args,options});
    const child=new EventEmitter();
    child.kill=()=>{child.killed=true;};
    const out=args.find(a=>a.startsWith('--print-to-pdf=')).slice('--print-to-pdf='.length);
    setImmediate(()=>{if(pdf!==null)fs.writeFileSync(out,pdf);if(exit)child.emit('exit',0);});
    return child;
  };
}
const tmpFolder=()=>fs.mkdtempSync(path.join(os.tmpdir(),'osg-pdf-test-'));
const leftOver=tmp=>fs.readdirSync(tmp).filter(n=>n.startsWith('osg-pdf-'));

test('Edge is found where Windows keeps it, and a computer without it says so',()=>{
  const seen=[];
  const found=findEdge({env:{'ProgramFiles(x86)':'C:\\PF86',ProgramFiles:'C:\\PF'},exists:f=>{seen.push(f);return f.startsWith('C:\\PF\\');}});
  assert.equal(found,path.join('C:\\PF','Microsoft','Edge','Application','msedge.exe'));
  assert.equal(seen.length,2,'the 32-bit program files are tried first, then the others');
  assert.equal(findEdge({env:{},exists:()=>true}),null,'no folder named, no Edge');
  assert.equal(findEdge({env:{ProgramFiles:'C:\\PF'},exists:()=>false}),null);
});

test('a PDF is tagged when it is marked and holds a structure tree, and not otherwise',()=>{
  assert.equal(pdfIsTagged(Buffer.from(TAGGED)),true);
  assert.equal(pdfIsTagged(Buffer.from(UNTAGGED)),false);
  assert.equal(pdfIsTagged(Buffer.from('%PDF-1.4 /StructTreeRoot but not marked')),false);
  assert.equal(pdfIsTagged(Buffer.from('%PDF-1.4 /Marked true but no tree')),false);
});

test('the PDF is made by Edge run on a page in a folder of its own, with fixed options, and the folder is gone afterwards',async()=>{
  const tmp=tmpFolder(),seen=[];
  try{
    const made=await makePdf(PAGE,{edge:'C:\\Edge\\msedge.exe',run:fakeRun(TAGGED,{seen}),tmp});
    assert.equal(made.tagged,true);assert.equal(made.bytes.subarray(0,5).toString(),'%PDF-');
    assert.equal(seen.length,1);
    const {exe,args,options}=seen[0];
    assert.equal(exe,'C:\\Edge\\msedge.exe');
    assert.equal(options.shell,false,'no shell');assert.equal(options.windowsHide,true);
    for(const flag of ['--headless=new','--no-pdf-header-footer','--disable-extensions','--no-first-run','--generate-pdf-document-outline'])assert.ok(args.includes(flag),flag);
    const page=args.at(-1);
    assert.match(page,/^file:\/\/\/.*page\.html$/,'the one address it is given is the page\'s file');
    assert.ok(!args.some(a=>/^https?:/i.test(a)),'no address on a network');
    assert.deepEqual(leftOver(tmp),[],'the folder the page and the PDF were in is removed');
    const untagged=await makePdf(PAGE,{edge:'e',run:fakeRun(UNTAGGED),tmp});
    assert.equal(untagged.tagged,false,'a PDF that is not tagged says so');
  }finally{fs.rmSync(tmp,{recursive:true,force:true});}
});

test('a page that holds a script, an empty page, a page that is too large and a computer with no Edge are refused with their reasons, and nothing is left behind',async()=>{
  const tmp=tmpFolder();
  try{
    const code=async(html,opts)=>{try{await makePdf(html,{edge:'e',run:fakeRun(TAGGED),tmp,...opts});return 'made';}catch(e){assert.ok(e instanceof PdfError);return e.code;}};
    assert.equal(await code('<p>hi</p><script>alert(1)</script>'),'refused');
    assert.equal(await code('   '),'empty');assert.equal(await code(null),'empty');
    assert.equal(await code('x'.repeat(PDF_MAX_HTML+1)),'too-large');
    assert.equal(await code(PAGE,{edge:null}),'no-edge');
    await assert.rejects(makePdf(PAGE,{edge:null,tmp}),/Microsoft Edge, which is not on this computer.*Word/);
    assert.deepEqual(leftOver(tmp),[]);
  }finally{fs.rmSync(tmp,{recursive:true,force:true});}
});

test('Edge that makes nothing, makes something that is not a PDF, cannot start or takes too long fails with its reason, kills what is slow, and leaves nothing behind',async()=>{
  const tmp=tmpFolder();
  try{
    const code=async run=>{try{await makePdf(PAGE,{edge:'e',run,tmp,timeoutMs:80});return 'made';}catch(e){return e.code;}};
    assert.equal(await code(fakeRun(null)),'no-output');
    assert.equal(await code(fakeRun('<html>not a pdf, and long enough to pass the size test of a real file..........</html>')),'bad-output');
    assert.equal(await code(()=>{throw new Error('spawn EACCES');}),'start');
    assert.equal(await code(()=>{const c=new EventEmitter();setImmediate(()=>c.emit('error',new Error('ENOENT')));c.kill=()=>{};return c;}),'start');
    let killed=false;
    assert.equal(await code(()=>{const c=new EventEmitter();c.kill=()=>{killed=true;};return c;}),'timeout');
    assert.equal(killed,true,'a browser that does not finish is stopped');
    assert.deepEqual(leftOver(tmp),[]);
  }finally{fs.rmSync(tmp,{recursive:true,force:true});}
});

const edge=findEdge();
test('with the real Edge on this computer, a guide\'s web page becomes a tagged PDF',{skip:edge?false:'Edge is not on this computer'},async()=>{
  const made=await makePdf(PAGE.replace('<main>','<main><h2>Second heading</h2><table><thead><tr><th scope="col">A</th></tr></thead><tbody><tr><td>b</td></tr></tbody></table>'),{edge});
  assert.equal(made.bytes.subarray(0,5).toString(),'%PDF-');
  assert.equal(made.tagged,true,'Edge marks the PDF and gives it a structure tree');
  const text=made.bytes.toString('latin1');
  assert.match(text,/\/Lang\s*\(en/,'it says its language');assert.match(text,/\/Title\s*[(<]/,'and its title');
  assert.equal(made.outlined,true,'and has bookmarks that follow its headings');
});

test('a PDF has bookmarks when it points at an outline, and not otherwise',()=>{
  assert.equal(pdfHasOutline(Buffer.from('%PDF-1.4 /Type /Catalog /Outlines 12 0 R /Pages 2 0 R')),true);
  assert.equal(pdfHasOutline(Buffer.from('%PDF-1.4 /Type /Catalog /Pages 2 0 R')),false);
  assert.equal(pdfHasOutline(Buffer.from('%PDF-1.4 the word /Outlines in a sentence')),false);
});

// ---- the server's answer
function call(port,method,urlPath,{headers={},body}={}){
  return new Promise((resolve,reject)=>{
    const req=http.request({host:'127.0.0.1',port,path:urlPath,method,headers:{Host:`127.0.0.1:${port}`,...headers}},res=>{
      const chunks=[];res.on('data',c=>chunks.push(c));res.on('end',()=>resolve({status:res.statusCode,headers:res.headers,body:Buffer.concat(chunks)}));
    });
    req.on('error',reject);if(body!==undefined)req.write(body);req.end();
  });
}
const JSONH={'Content-Type':'application/json'};
async function withServer(pdfMaker,fn){
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-pdf-server-'));
  const s=await startProjectServer({folder:path.join(dir,'guides'),port:0,snapshot:{version:'0-test',buildId:'test-build',record:{},read:()=>null},pdfMaker});
  try{await fn(s.port);}finally{await s.close();fs.rmSync(dir,{recursive:true,force:true});}
}

test('POST /api/pdf gives back the PDF the maker made, says whether it is tagged, and takes only this page\'s own request',async()=>{
  const got=[];
  await withServer(async html=>{got.push(html);return {bytes:Buffer.from(TAGGED),tagged:true};},async port=>{
    const ok=await call(port,'POST','/api/pdf',{headers:JSONH,body:JSON.stringify({html:PAGE})});
    assert.equal(ok.status,200);assert.equal(ok.headers['content-type'],'application/pdf');assert.equal(ok.headers['x-pdf-tagged'],'yes');
    assert.equal(ok.body.toString(),TAGGED,'the bytes, whole');assert.deepEqual(got,[PAGE]);
    assert.equal((await call(port,'POST','/api/pdf',{headers:JSONH,body:JSON.stringify({html:''})})).status,400,'nothing to print');
    assert.equal((await call(port,'POST','/api/pdf',{headers:JSONH,body:'not json'})).status,400);
    assert.equal((await call(port,'POST','/api/pdf',{headers:{'Content-Type':'text/plain'},body:PAGE})).status,415,'only in the form the program sends');
    assert.equal((await call(port,'POST','/api/pdf',{headers:{...JSONH,Origin:'http://evil.example'},body:JSON.stringify({html:PAGE})})).status,403,'not from another site');
    assert.equal((await call(port,'POST','/api/pdf',{headers:{...JSONH,'Sec-Fetch-Site':'cross-site'},body:JSON.stringify({html:PAGE})})).status,403);
    assert.equal((await call(port,'GET','/api/pdf')).status,405);
    assert.equal(got.length,1,'only the one good request reached the maker');
  });
});

test('the server tells what went wrong in words, makes one PDF at a time, and answers an untagged PDF as untagged',async()=>{
  let release;const gate=new Promise(r=>{release=r;});
  let mode='slow';
  await withServer(async()=>{
    if(mode==='slow'){await gate;return {bytes:Buffer.from(TAGGED),tagged:true};}
    if(mode==='untagged')return {bytes:Buffer.from(UNTAGGED),tagged:false};
    throw new PdfError({'no-edge':'No Edge.',refused:'Script.','too-large':'Too big.',failed:'Broke.'}[mode],mode);
  },async port=>{
    const body=JSON.stringify({html:PAGE});
    const first=call(port,'POST','/api/pdf',{headers:JSONH,body});
    await new Promise(r=>setTimeout(r,60));
    const second=await call(port,'POST','/api/pdf',{headers:JSONH,body});
    assert.equal(second.status,409);assert.match(JSON.parse(second.body).error,/already being made/);
    release();assert.equal((await first).status,200);
    mode='untagged';const un=await call(port,'POST','/api/pdf',{headers:JSONH,body});
    assert.equal(un.status,200);assert.equal(un.headers['x-pdf-tagged'],'no');
    for(const [m,status,words] of [['no-edge',501,/No Edge/],['refused',400,/Script/],['too-large',400,/Too big/],['failed',500,/Broke/]]){
      mode=m;const r=await call(port,'POST','/api/pdf',{headers:JSONH,body});
      assert.equal(r.status,status,m);assert.match(JSON.parse(r.body).error,words);assert.equal(JSON.parse(r.body).code,m);
    }
    mode='untagged';assert.equal((await call(port,'POST','/api/pdf',{headers:JSONH,body})).status,200,'and the next one is not held up by a failure');
  });
});

test('the program\'s document makes a PDF when it is given a way to, and otherwise says the program makes none',async()=>{
  const xml={DOMParserClass:DOMParser,SerializerClass:XMLSerializer};
  const plain=createProjectAdapter(createDemoAdapter(xml),{read:async()=>null,write:async()=>{},soon:()=>0,cancel:()=>{}});
  await plain.ready;await assert.rejects(plain.exportPdf(),{message:NO_PDF});
  const base64=Buffer.from(TAGGED).toString('base64');
  const doc=createProjectAdapter(createDemoAdapter(xml),{read:async()=>null,write:async()=>{},soon:()=>0,cancel:()=>{},exportPdf:async()=>base64});
  await doc.ready;assert.equal(await doc.exportPdf(),base64);
});
