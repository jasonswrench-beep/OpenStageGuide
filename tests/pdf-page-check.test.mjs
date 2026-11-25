// The page is read before it is printed (scripts/pdf-maker.mjs, preparePdfPage): anything that could make the printing reach outside the page is refused, the program's own
// pages are let through, and the policy that forbids the rest is put first in the head by the maker, whatever the page says. Nothing here needs Edge.
import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {EventEmitter} from 'node:events';
import {fileURLToPath} from 'node:url';
import {DOMParser} from '@xmldom/xmldom';
import {makePdf,preparePdfPage,PDF_PAGE_POLICY,PdfError} from '../scripts/pdf-maker.mjs';
import {startProjectServer} from '../scripts/project-server.mjs';
import {guideToHtml} from '../src/guide-html.mjs';
import {describeGuide} from '../program/ui/preview.js';
import {glanceBody} from '../src/guide-pages.mjs';
import {readGuide} from '../src/guide-model.mjs';
import {started,addModule,addLesson,addBlock} from './helpers/guide-harness.mjs';

const META=`<meta http-equiv="Content-Security-Policy" content="${PDF_PAGE_POLICY}">`;
const PNG='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const DATA=`data:image/png;base64,${PNG}`;
const wrap=(body,head='')=>`<!doctype html><html lang="en"><head><meta charset="utf-8">${head}<title>T</title></head><body><main>${body}</main></body></html>`;
const refused=(html,what,pattern)=>assert.throws(()=>preparePdfPage(html),e=>e instanceof PdfError&&e.code==='refused'&&(!pattern||pattern.test(e.message)),what||html.slice(0,100));

async function guidePage(){
  const a=await started('Web guide');
  await addModule(a,'One');await addLesson(a,'First','10');await addBlock(a,'say','Welcome everyone.');
  await a.adapter.insert(()=>glanceBody(readGuide(a.adapter.snapshot()),{withHeading:true}),{where:'after',after:1,spare:false});
  const entries=a.adapter.exportState().paragraphs;
  const nodes=describeGuide(entries,xml=>new DOMParser().parseFromString(xml,'application/xml'),{media:id=>a.adapter.mediaOf(id)});
  return guideToHtml({nodes,title:'Web guide',picture:()=>DATA});
}

test('the guide\'s own page is let through as it is, with its pictures, its contents links and its tables, and the policy is added first in its head',async()=>{
  const html=await guidePage();
  assert.match(html,/<img [^>]*src="data:image\/png;base64,/,'the page holds pictures');
  assert.match(html,/<a [^>]*href="#/,'and links to places in itself');
  assert.match(html,/<table/);
  const own=/<meta http-equiv="Content-Security-Policy" content="([^"]*)">/.exec(html);
  assert.equal(own[1],PDF_PAGE_POLICY,'the policy the maker uses is the one the page says');
  const prepared=preparePdfPage(html);
  assert.equal(prepared,html.replace('<head>',`<head>${META}`),'nothing else is changed');
  assert.equal(prepared.indexOf(META),prepared.indexOf('<head>')+'<head>'.length,'first in the head');
});

test('a page with no head, no html or no doctype gets a head of its own, with the policy first, and is otherwise as it was',()=>{
  const strip=(prepared,added)=>prepared.replace(added,'');
  let p=preparePdfPage('<p>Hi</p>');
  assert.equal(p,`<head>${META}</head><p>Hi</p>`);
  p=preparePdfPage('<!doctype html><p>Hi</p>');
  assert.equal(p,`<!doctype html><head>${META}</head><p>Hi</p>`);
  p=preparePdfPage('<!DOCTYPE html>\n<html lang="en"><body><p>Hi</p></body></html>');
  assert.equal(strip(p,`<head>${META}</head>`),'<!DOCTYPE html>\n<html lang="en"><body><p>Hi</p></body></html>');
  assert.ok(p.indexOf(`<head>${META}</head>`)>p.indexOf('<html lang="en">'),'after the html tag');
  p=preparePdfPage('\uFEFF  <!doctype html>\n<html><head>\n<title>T</title></head><body></body></html>');
  assert.ok(p.includes(`<head>${META}\n<title>`),'a head after blanks and the doctype');
});

test('a policy of the page\'s own stays, but ours is first, so it can only be narrowed',()=>{
  const prepared=preparePdfPage(wrap('<p>Hi</p>',`<meta http-equiv="Content-Security-Policy" content="default-src *">`));
  const first=prepared.indexOf(META),theirs=prepared.indexOf('default-src *');
  assert.ok(first>0&&theirs>first);
  assert.equal(prepared.indexOf(META),prepared.indexOf('<head>')+'<head>'.length);
});

test('an address in a page is a picture held inside it or a place in it, and nothing else: a share, a site, a file, a script, however it is written',()=>{
  const must=[
    '<img src="\\\\server\\share\\x.png">',            // a network share (it would be asked for, and sign-in details sent with the asking)
    '<img src="//server/share/x.png">',
    '<img src="http://example.invalid/x.png">',
    '<img src="https://example.invalid/x.png">',
    '<img src="file:///C:/x.png">',
    '<img src="C:\\x.png">',
    '<img src="x.png">',
    '<img src="">',
    '<img src=" data:image/png;base64,AAAA">',          // blanks in front
    '<img src="data:image/png;base64,AAAA ">',
    '<img src="data:image/svg+xml;base64,AAAA">',       // a picture that can carry more than a picture
    '<img src="data:text/html;base64,AAAA">',
    '<img src="data:image/png;base64,AAAA&#10;">',
    '<img src="&#104;ttp://example.invalid/x.png">',    // written as a character reference
    '<img src="javascript:alert(1)">',
    '<img src=//server/x.png>',                         // not in quotes
    '<img/src="//server/x.png">',                       // a slash for a space
    '<IMG SRC="//server/x.png">',                       // capitals
    '<img\nsrc\n=\n"//server/x.png">',                  // broken across lines
    '<img srcset="//server/x.png 1x">',
    '<video poster="//server/x.png"></video>',
    '<a href="https://example.invalid/x">x</a>',
    '<a href="//server/x">x</a>',
    '<a href="mailto:a@example.invalid">x</a>',
    '<form action="//server/x"></form>',
    '<input formaction="//server/x">',
    '<table><tr><td background="//server/x.png">x</td></tr></table>',
    '<a href="#a" ping="//server/x">x</a>',
    '<a href="#a" href="//server/x">x</a>',             // twice: both are read
    '<img src="data:image/png;base64,AAAA" src="//server/x.png">',
    '<object data="//server/x"></object>',
  ];
  for(const body of must)refused(wrap(body),body);
  // and what is let through
  for(const body of ['<img src="data:image/png;base64,AAAA" alt="">','<img src=\'data:image/gif;base64,R0lGODlhAQABAAAAACw=\'>','<a href="#main">Skip</a>','<a href="#">Top</a>','<IMG SRC="data:image/jpeg;base64,/9j/4AAQ">','<p title=\'<img src=//server/x.png>\'>an address in an attribute is only text</p>','<p title="a>b">x</p>']){
    assert.doesNotThrow(()=>preparePdfPage(wrap(body)),body);
  }
});

test('frames, objects, embeds, style sheets on other addresses, a base address, a redirect, a script and foreign markup are refused as elements, in any spelling',()=>{
  const must=[
    ['<link rel="stylesheet" href="//server/x.css">',/link element/],
    ['<link rel=stylesheet href=#x>',/link element/],
    ['<iframe src="#x"></iframe>',/iframe element/],
    ['<IFRAME SRC="#x"></IFRAME>',/iframe element/],
    ['<iframe\n>',/iframe element/],
    ['<iframe srcdoc="<img src=//server/x.png>"></iframe>',/iframe element|script handler/],
    ['<object></object>',/object element/],
    ['<embed>',/embed element/],
    ['<base href="#">',/base element/],
    ['<applet></applet>',/applet element/],
    ['<frameset></frameset>',/frameset element/],
    ['<svg><image href="//server/x.png"/></svg>',/svg element/],
    ['<math><mi href="//server/x"></mi></math>',/math element/],
    ['<script/>',/script element/],
    ['<script\n>x</script>',/script element/],
    ['<meta http-equiv="refresh" content="0;url=//server/x">',/meta element that acts/],
    ['<meta http-equiv="Refresh" content="0;url=//server/x">',/meta element that acts/],
    ['<meta http-equiv="set-cookie" content="a=b">',/meta element that acts/],
    ['<meta charset="utf-7">',/not written in UTF-8/],
  ];
  for(const [body,pattern] of must)refused(wrap(body),body,pattern);
  for(const head of ['<meta charset="UTF-8">','<meta name="viewport" content="width=device-width, initial-scale=1">','<meta name="generator" content="OpenStage Guide">',`<meta http-equiv="Content-Security-Policy" content="${PDF_PAGE_POLICY}">`,'<meta http-equiv="content-security-policy" content="default-src \'none\'">']){
    assert.doesNotThrow(()=>preparePdfPage(wrap('<p>x</p>',head)),head);
  }
});

test('a script handler is refused, and so is a style that can name an address, however it is written',()=>{
  for(const body of ['<p onclick="x()">x</p>','<p ONMOUSEOVER=x>x</p>','<p\nonfocus = "x">x</p>','<img src="data:image/png;base64,AAAA" onerror="x()">']){
    refused(wrap(body),body,/script handler/);
  }
  const styles=[
    '<p style="background:url(//server/x.png)">x</p>',
    '<p style="background:URL( //server/x.png )">x</p>',
    '<p style="background:u&#114;l(//server/x.png)">x</p>',   // written as a character reference
    '<p style="background:\\75rl(//server/x.png)">x</p>',      // written as a CSS escape
    '<p style="background:image-set(\'//server/x.png\' 1x)">x</p>',
    '<p style="background:cross-fade(\'//server/x.png\', red)">x</p>',
    '<p style="behavior:expression(alert(1))">x</p>',
    '<style>@import "//server/x.css";</style>',
    '<style>@import url(//server/x.css);</style>',
    '<style>p{background:url(\\\\server\\share\\x.png)}</style>',
    '<style>p{background:image-set("//server/x.png" 1x)}</style>',
    '<STYLE>p{background:URL(//server/x.png)}</STYLE>',
    '<style>p{}</style ><img src=//server/x.png>',            // a style that ends with a blank in its end tag, and then a picture
  ];
  for(const body of styles)refused(wrap(body),body);
  for(const body of ['<p style="font-family:&quot;Arial&quot;;color:#333;margin-left:12px">x</p>','<style>p{color:#333}ul>li{margin:0}@media print{a{color:inherit}}</style>']){
    assert.doesNotThrow(()=>preparePdfPage(wrap(body)),body);
  }
});

test('comments, CDATA, instructions, tags that are not closed and a page that is not laid out as the program\'s are refused, not guessed at',()=>{
  const must=[
    '<!-- a comment -->','<!-->','<!--->','<!-- a --!><img src=//server/x.png>-->','<![CDATA[x]]>','<?xml version="1.0"?>','<!ELEMENT x>',
    '<img src="data:image/png;base64,AAAA','<p class="a',
    'a\u0000b',
  ];
  for(const body of must)refused(wrap(body),body);
  // a tag, an unquoted value or a style that the page ends inside (nothing after it to close it)
  for(const page of ['<!doctype html><html><head></head><body><p','<!doctype html><html><head></head><body><img src=x','<!doctype html><html><head></head><body><p class=','<!doctype html><html><head><style>p{}','<!doctype html><html><head><style'])refused(page,page);
  refused('<div>x</div><head><title>T</title></head><body></body>','something before the head',/before its head/);
  refused('hello <head><title>T</title></head>','text before the head',/before its head/);
  refused('<html><body><head><title>T</title></head></body></html>','a head inside the body',/before its head/);
  refused('<p>x</p><!doctype html>','a doctype that is not first');
  refused('<!doctype html><!doctype html><p>x</p>','two doctypes');
});

test('what a browser would not read as a tag is not read as one here, and what it would is: quotes, end tags, the insides of a title or a text box',()=>{
  // inside a quoted value it is only text
  assert.doesNotThrow(()=>preparePdfPage(wrap('<p title="<img src=//server/x.png>" data-x=\'<iframe>\'>x</p>')));
  // the insides of a title, a text box and an end tag are read as tags too, which is stricter than a browser, and never less
  refused(wrap('<textarea><img src="//server/x.png"></textarea>'));
  refused('<!doctype html><html><head><title><img src="//server/x.png"></title></head><body></body></html>');
  refused(wrap('<p>x</p </p><img src=//server/x.png>>'));
  refused(wrap('</p <img src=//server/x.png>>'));
  // a "<" that is only text
  assert.doesNotThrow(()=>preparePdfPage(wrap('<p>1 < 2 and 3 <4 and a < b, a<3</p>')));
  assert.doesNotThrow(()=>preparePdfPage(wrap('<p>x</p></>')));
});

test('a large page is read in good time, and a long picture is only checked for its shape',()=>{
  const picture='A'.repeat(8_000_000);
  const body=`<p class="a">x</p>\n`.repeat(100_000)+`<img src="data:image/png;base64,${picture}">`;
  const t=Date.now();
  const prepared=preparePdfPage(wrap(body));
  assert.ok(prepared.length>8_000_000);
  assert.ok(Date.now()-t<5000,`took ${Date.now()-t} ms`);
  const t2=Date.now();
  refused(wrap(`<img src="data:image/png;base64,${picture}!">`));
  assert.ok(Date.now()-t2<5000);
});

// A stand-in for Edge that keeps the page it was given and writes a PDF.
function recordingRun(pages,{started=[]}={}){
  return (exe,args)=>{
    started.push(args);
    const child=new EventEmitter();child.kill=()=>{};
    const file=fileURLToPath(args[args.length-1]);
    pages.push(fs.readFileSync(file,'utf8'));
    const out=args.find(a=>a.startsWith('--print-to-pdf=')).slice('--print-to-pdf='.length);
    setImmediate(()=>{fs.writeFileSync(out,'%PDF-1.4\n/MarkInfo << /Marked true >> /StructTreeRoot 1 0 R\n'+'x'.repeat(80)+'\n%%EOF\n');child.emit('exit',0);});
    return child;
  };
}
const tmpFolder=()=>fs.mkdtempSync(path.join(os.tmpdir(),'osg-pdf-page-test-'));
const leftOver=tmp=>fs.readdirSync(tmp).filter(n=>n.startsWith('osg-pdf-'));

test('makePdf prints the page with the policy first, and a refused page never starts Edge or leaves a folder',async()=>{
  const tmp=tmpFolder(),pages=[],started=[];
  try{
    const html=await guidePage();
    const made=await makePdf(html,{edge:'edge.exe',run:recordingRun(pages,{started}),tmp,pollMs:10});
    assert.equal(made.tagged,true);
    assert.equal(pages.length,1);
    assert.equal(pages[0].indexOf(META),pages[0].indexOf('<head>')+'<head>'.length,'the page Edge was given has the policy first in its head');
    for(const bad of [wrap('<img src="\\\\server\\share\\x.png">'),wrap('<link rel="stylesheet" href="//server/x.css">'),wrap('<iframe src="#x"></iframe>')]){
      await assert.rejects(makePdf(bad,{edge:'edge.exe',run:recordingRun(pages,{started}),tmp,pollMs:10}),e=>e instanceof PdfError&&e.code==='refused');
    }
    assert.equal(started.length,1,'Edge was started once, for the guide, and never for a refused page');
    assert.deepEqual(leftOver(tmp),[],'no folder is left');
    // a refused page is refused on a computer with no Edge too, with the reason of the page first
    await assert.rejects(makePdf(wrap('<iframe src="#x"></iframe>'),{edge:null,tmp}),e=>e.code==='refused');
  }finally{fs.rmSync(tmp,{recursive:true,force:true});}
});

test('the server answers a refused page with 400 and the reason, before Edge is looked for',async()=>{
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-pdf-page-server-'));
  const s=await startProjectServer({folder:path.join(dir,'guides'),dataFolder:path.join(dir,'data'),port:0,snapshot:{version:'0-test',buildId:'test-build',record:{},read:()=>null}});
  const post=html=>new Promise((resolve,reject)=>{
    const body=JSON.stringify({html});
    const req=http.request({host:'127.0.0.1',port:s.port,path:'/api/pdf',method:'POST',headers:{Host:`127.0.0.1:${s.port}`,'Content-Type':'application/json','Content-Length':Buffer.byteLength(body)}},res=>{const chunks=[];res.on('data',c=>chunks.push(c));res.on('end',()=>resolve({status:res.statusCode,json:()=>JSON.parse(Buffer.concat(chunks).toString('utf8'))}));});
    req.on('error',reject);req.end(body);
  });
  try{
    for(const [bad,pattern] of [[wrap('<img src="\\\\server\\share\\x.png">'),/points outside the page/],[wrap('<link rel="stylesheet" href="//server/x.css">'),/link element/],[wrap('<iframe src="#x"></iframe>'),/iframe element/]]){
      const r=await post(bad);
      assert.equal(r.status,400);
      const j=r.json();
      assert.equal(j.code,'refused');
      assert.match(j.error,pattern);
      assert.match(j.error,/so it was not printed\.$/);
      assert.ok(!j.error.includes('server'),'the address in the page is not repeated back');
    }
  }finally{await s.close();fs.rmSync(dir,{recursive:true,force:true});}
});
