// The program's own files (program/): what they may and may not do, and that they agree with each other. Like the suite's other programs, the
// interface writes text and attributes only, reaches nothing outside this computer, and every colour and icon it names exists.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {UI_ICONS} from '../program/ui/ui-icons.js';
import {TOOL_ICONS} from '../program/ui/commands.js';

const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const dir=name=>path.join(root,'program',name);
const read=file=>fs.readFileSync(file,'utf8');
const uiFiles=fs.readdirSync(dir('ui')).filter(f=>f.endsWith('.js')).map(f=>path.join(dir('ui'),f));
// The code with its comments taken out, so words in a comment are not mistaken for words in the program.
const code=file=>read(file).replace(/\/\*[\s\S]*?\*\//g,'').replace(/^\s*\/\/.*$/gm,'');

test('the interface writes text and attributes only: no markup, no scripts made from text',()=>{
  const banned=/\b(innerHTML|outerHTML|insertAdjacentHTML|document\.write|createContextualFragment)\b|\beval\s*\(|new Function\s*\(|setAttribute\(\s*["']on|setAttribute\(\s*["']style["']|\.srcdoc\b/;
  for(const f of uiFiles)assert.doesNotMatch(code(f),banned,path.basename(f));
});

test('the interface reaches nothing but its own server: the only addresses in it are namespaces',()=>{
  const allowed=/^(www\.w3\.org|schemas\.openxmlformats\.org|schemas\.microsoft\.com)$/;
  for(const f of uiFiles){
    for(const m of code(f).matchAll(/https?:\/\/([^/"'`\s)]+)/g))assert.match(m[1],allowed,`${path.basename(f)}: ${m[0]}`);
    assert.doesNotMatch(code(f),/\b(XMLHttpRequest|WebSocket|EventSource|sendBeacon|navigator\.geolocation|document\.cookie|localStorage|sessionStorage)\b/,path.basename(f));
  }
  // fetch is only to /api and /project, on this server
  for(const f of uiFiles)for(const m of code(f).matchAll(/fetch\(\s*([^,)]+)/g))assert.match(m[1],/^(`\/api\/\$\{path\}`|"\/project")$/,`${path.basename(f)}: fetch(${m[1]}`);
});

test('every icon the interface names is one the program has, and the icon file is the one its list makes',()=>{
  const have=new Set(Object.keys(UI_ICONS));
  const named=new Set();
  for(const f of uiFiles.filter(f=>!f.endsWith('ui-icons.js')))for(const m of code(f).matchAll(/(?:uiIcon\(\s*doc\s*,\s*|\bicon:\s*)"([a-z0-9-]+)"/g))named.add(m[1]);
  for(const name of named)assert.ok(have.has(name),`${name} is in ui-icons.js`);
  const list=JSON.parse(read(dir('icons.json'))).icons;
  assert.deepEqual([...new Set(list)].sort(),[...have].sort(),'program/icons.json and ui-icons.js agree: run scripts/make-ui-icons.mjs after changing the list');
  for(const shapes of Object.values(UI_ICONS))assert.ok(Array.isArray(shapes)&&shapes.length>0&&shapes.every(s=>typeof s.tag==='string'));
  for(const name of Object.values(TOOL_ICONS))named.add(name);
  for(const name of named)assert.ok(have.has(name),`${name} is in ui-icons.js`);
  assert.ok(named.size>=20,'a good many icons are in use');
});

test('every colour name the stylesheets use is defined, and the Guide\'s pair is the one the author accepted',()=>{
  const css=read(dir('style.css')),suite=read(dir('openstage-ui.css'));
  const defined=new Set([...(css+suite).matchAll(/(--[a-z0-9-]+)\s*:/g)].map(m=>m[1]));
  const strict=new Set(),loose=new Set();
  for(const m of (css+suite).matchAll(/var\(\s*(--[a-z0-9-]+)\s*([,)])/g))(m[2]===','?loose:strict).add(m[1]);
  for(const name of strict)assert.ok(defined.has(name)||name==='--top-height',`${name} is defined`);
  assert.match(css,/^:root \{/m,'the variables are in a :root block (a slice of another file once left it out)');
  for(const [name,value] of [['--guide-action','#0e7490'],['--guide-action-hover','#155e75'],['--guide-tint','#e3f4f7'],['--guide-on-shell','#67e8f9']])assert.match(css,new RegExp(`${name}:\\s*${value};`));
  assert.doesNotMatch(css,/--assess-/,'no name left over from the program this was copied from');
  assert.ok(loose.size>=0);
});

test('the Guide\'s colours have the contrast the design guide asks for, measured here',()=>{
  const lum=hex=>{const c=[1,3,5].map(i=>parseInt(hex.slice(i,i+2),16)/255).map(v=>v<=0.03928?v/12.92:Math.pow((v+0.055)/1.055,2.4));return 0.2126*c[0]+0.7152*c[1]+0.0722*c[2];};
  const ratio=(a,b)=>{const x=lum(a),y=lum(b);return (Math.max(x,y)+0.05)/(Math.min(x,y)+0.05);};
  assert.ok(ratio('#ffffff','#0e7490')>=4.5,'white words on a filled button');
  assert.ok(ratio('#ffffff','#155e75')>=4.5,'white words under the pointer');
  assert.ok(ratio('#0e7490','#e3f4f7')>=4.5,'the blue\'s words on the chosen item');
  assert.ok(ratio('#67e8f9','#0d1b30')>=7,'the name on the navy bar');
  assert.ok(ratio('#0e7490','#ffffff')>=4.5,'a link on white');
  assert.ok(ratio('#0e7490','#eef1f5')>=4.5,'a link on the page grey');
});

test('the page has a language, a skip link, one main region, and no inline script or style',()=>{
  const html=read(dir('index.html'));
  assert.match(html,/<html lang="en">/);
  assert.match(html,/<a class="skip" href="#main">/);
  assert.equal((html.match(/<main\b/g)||[]).length,1);
  assert.doesNotMatch(html,/<script(?![^>]*\bsrc=)[^>]*>/,'every script is a file');
  assert.doesNotMatch(html,/\sstyle=|\son[a-z]+=/i,'no inline style or event attribute');
  for(const m of html.matchAll(/(?:src|href)="([^"#]+)"/g))assert.ok(m[1].startsWith('/'),`${m[1]} is on this server`);
});

test('the copies from the suite say where they came from, and the NOTICE credits them',()=>{
  for(const f of ['menubar.js','popupmenu.js','commandsearch.js','ui-icons.js'])assert.match(read(path.join(dir('ui'),f)),/OpenStage Author|Lucide|OpenAssess/,f);
  assert.match(read(path.join(root,'NOTICE')),/program\/ui/);
});
