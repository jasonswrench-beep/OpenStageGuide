// The fonts that come with the program: that the files are what the list says, are TrueType fonts that allow embedding, come with their licenses
// and credits, are drawn by the page's stylesheet, are served only by plain name, and travel inside a Word file that uses them.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import os from 'node:os';
import {fileURLToPath} from 'node:url';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {FONT_PACK,NAMED_ONLY,COMMON_FONTS,STYLES,packFont,namedOnly,namedOnlyUsed,fontNamesIn,packFontsUsed,fontChoices,obfuscateFont,newGuid} from '../src/guide-fonts.mjs';
import {docxFromPackage,EMBED_LIMITS} from '../src/guide-docx.mjs';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {startProjectServer} from '../scripts/project-server.mjs';

const ROOT=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const DIR=path.join(ROOT,'program','fonts');
const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
const bytesOf=file=>new Uint8Array(fs.readFileSync(path.join(DIR,file)));

// The table directory of a TrueType file: what tables it has, and where.
function tables(bytes){
  const v=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);
  const n=v.getUint16(4),out={};
  for(let i=0;i<n;i++){const at=12+i*16;out[String.fromCharCode(...bytes.slice(at,at+4))]={offset:v.getUint32(at+8),length:v.getUint32(at+12)};}
  return {version:v.getUint32(0),out,v};
}

test('every font the list names is a TrueType file in program/fonts, with its license beside it, and nothing else is there',()=>{
  const listed=new Set();
  for(const f of FONT_PACK){
    assert.ok(f.files.regular,`${f.family} has a regular style`);
    for(const [style,file] of Object.entries(f.files)){
      assert.ok(STYLES.includes(style));
      assert.ok(/^[A-Za-z0-9_-]+\.ttf$/.test(file),`${file} is a plain file name`);
      listed.add(file);
      const t=tables(bytesOf(file));
      assert.equal(t.version,0x00010000,`${file} is TrueType (not a PostScript-outlined OpenType, which Word cannot carry)`);
      assert.ok(t.out.glyf&&t.out.name&&t.out['OS/2'],`${file} has outlines, names and metrics`);
      assert.ok(!t.out.fvar,`${file} is a static font, not a variable one`);
      assert.equal(t.v.getUint16(t.out['OS/2'].offset+8),0,`${file} allows embedding in documents (fsType 0)`);
    }
    listed.add(f.license);
    const license=fs.readFileSync(path.join(DIR,f.license),'utf8');
    assert.match(license,/SIL OPEN FONT LICENSE Version 1\.1/i,`${f.family}'s license is the SIL Open Font License 1.1`);
    assert.ok(f.credit&&f.source.startsWith('https://github.com/'),`${f.family} has a credit and a source`);
  }
  assert.deepEqual(fs.readdirSync(DIR).sort(),[...listed].sort(),'no file is there that the list does not name, and none is missing');
});

test('the files are the makers\' own, unchanged: each font\'s own name and license text say so',()=>{
  const nameOf=(bytes,id)=>{
    const t=tables(bytes),base=t.out.name.offset,count=t.v.getUint16(base+2),str=base+t.v.getUint16(base+4);
    for(let i=0;i<count;i++){
      const r=base+6+i*12;
      if(t.v.getUint16(r+6)!==id)continue;
      const platform=t.v.getUint16(r),len=t.v.getUint16(r+8),off=t.v.getUint16(r+10);
      const raw=bytes.slice(str+off,str+off+len);
      if(platform===3||platform===0)return Buffer.from(raw).swap16().toString('utf16le');
      return Buffer.from(raw).toString('latin1');
    }
    return '';
  };
  for(const f of FONT_PACK)for(const file of Object.values(f.files)){
    const bytes=bytesOf(file);
    assert.equal(nameOf(bytes,1).replace(/ (Bold|Italic|Bold Italic)$/,'')===f.family||nameOf(bytes,16)===f.family||nameOf(bytes,1)===f.family,true,`${file} is the family ${f.family}`);
    assert.match(nameOf(bytes,13),/SIL Open Font License, Version 1\.1/,`${file} carries its license`);
  }
});

test('NOTICE credits every font and its license, and says they are unchanged',()=>{
  const notice=fs.readFileSync(path.join(ROOT,'NOTICE'),'utf8');
  for(const f of FONT_PACK){
    assert.ok(notice.includes(f.family),`NOTICE names ${f.family}`);
    assert.ok(notice.includes(f.license),`NOTICE says where ${f.family}'s license is (${f.license})`);
    assert.ok(notice.includes(f.credit.replace(/^Copyright /,'')),`NOTICE has the credit: ${f.credit}`);
  }
  assert.match(notice,/SIL Open Font License/);
  assert.match(notice,/unmodified|unchanged/i);
});

test('the stylesheet draws every font of the list, in every style it has, and no other',()=>{
  const css=fs.readFileSync(path.join(ROOT,'program','style.css'),'utf8');
  const faces=[...css.matchAll(/@font-face\s*\{([^}]*)\}/g)].map(m=>{
    const body=m[1];
    return {family:/font-family:\s*"([^"]+)"/.exec(body)?.[1],weight:/font-weight:\s*(\d+)/.exec(body)?.[1],style:/font-style:\s*(\w+)/.exec(body)?.[1],file:/url\("\/program\/fonts\/([^"]+)"\)/.exec(body)?.[1]};
  });
  const want=[];
  for(const f of FONT_PACK)for(const [style,file] of Object.entries(f.files))want.push({family:f.family,weight:style.startsWith('bold')?'700':'400',style:/italic/i.test(style)?'italic':'normal',file});
  const key=x=>`${x.family}|${x.weight}|${x.style}|${x.file}`;
  assert.deepEqual(faces.map(key).sort(),want.map(key).sort());
});

test('a font name is asked for in any of Word\'s four places, and only the program\'s own fonts are found',()=>{
  const xml='<w:rFonts w:ascii="Lexend" w:hAnsi="Lexend" w:eastAsia="Lexend" w:cs="Lexend"/><w:rFonts w:ascii="Calibri"/><w:rFonts w:hAnsi="Source Sans 3"/>';
  assert.deepEqual(fontNamesIn(xml).sort(),['Calibri','Lexend','Source Sans 3']);
  assert.deepEqual(packFontsUsed(xml).map(f=>f.family),['Lexend','Source Sans 3'],'in the order of the pack');
  assert.deepEqual(packFontsUsed('<w:rFonts w:ascii="lexend"/>'),[],'a name has to be the name');
  assert.equal(packFont('Lexend').kind,'sans');
  assert.equal(packFont('Calibri'),null);
  const choices=fontChoices();
  assert.deepEqual(choices.slice(0,FONT_PACK.length).map(c=>c.value),FONT_PACK.map(f=>f.family),'the pack comes first');
  assert.ok(choices.slice(0,FONT_PACK.length).every(c=>/comes with the program/.test(c.label)));
  assert.ok(choices.some(c=>c.value==='Calibri'));
});

// ---- the fonts that are only named

test('OpenDyslexic is offered by name only: after the pack, before the usual Word fonts, saying it is not included and where to get it',()=>{
  assert.deepEqual(NAMED_ONLY.map(f=>f.family),['OpenDyslexic']);
  const choices=fontChoices();
  assert.deepEqual(choices.map(c=>c.value),[...FONT_PACK.map(f=>f.family),'OpenDyslexic',...COMMON_FONTS],'the pack, then the name-only fonts, then the usual ones');
  const label=choices.find(c=>c.value==='OpenDyslexic').label;
  assert.match(label,/not included/);
  assert.match(label,/opendyslexic\.org/);
  assert.match(label,/only names it/);
  assert.doesNotMatch(label,/comes with the program/,'it must not say what is not true');
  assert.ok(NAMED_ONLY.every(f=>f.note&&f.source.startsWith('https://')));
  assert.equal(namedOnly('OpenDyslexic').family,'OpenDyslexic');
  assert.equal(namedOnly('Lexend'),null);
  assert.equal(packFont('OpenDyslexic'),null,'it is not in the pack');
  assert.ok(!NAMED_ONLY.some(f=>packFont(f.family)||COMMON_FONTS.includes(f.family)),'no font is on two lists');
});

test('nothing of a name-only font is shipped: no file, no stylesheet rule, no credit that implies one',()=>{
  const css=fs.readFileSync(path.join(ROOT,'program','style.css'),'utf8');
  const files=fs.readdirSync(DIR).join('\n');
  for(const f of NAMED_ONLY){
    assert.ok(!new RegExp(f.family.replace(/\s/g,''),'i').test(files),`no file of ${f.family} is in program/fonts`);
    assert.ok(!css.includes(`font-family:"${f.family}"`)&&!/@font-face[^}]*OpenDyslexic/i.test(css),`the stylesheet loads no ${f.family}`);
  }
});

test('the name-only fonts a piece of Word XML uses are found by their exact name, and never as the program\'s own',()=>{
  const xml='<w:rFonts w:ascii="OpenDyslexic" w:hAnsi="OpenDyslexic"/><w:rFonts w:ascii="Lexend"/>';
  assert.deepEqual(namedOnlyUsed(xml).map(f=>f.family),['OpenDyslexic']);
  assert.deepEqual(packFontsUsed(xml).map(f=>f.family),['Lexend'],'the pack still finds only its own');
  assert.deepEqual(namedOnlyUsed('<w:rFonts w:ascii="opendyslexic"/>'),[]);
  assert.deepEqual(namedOnlyUsed('<w:rFonts w:ascii="Calibri"/>',undefined,null),[]);
});

test('a Word file that uses OpenDyslexic only names it: no font table, no files, and the loader is never asked for it',async()=>{
  const a=await guideUsing('OpenDyslexic');
  const asked=[],named=[],embedded=[];
  const zip=await JSZip.loadAsync(await docxFromPackage(await a.sourcePackage(),JSZip,io,{entries:a.exportState().paragraphs,fonts:async families=>{asked.push(...families);return {};},onEmbedded:l=>embedded.push(...l),onNamed:l=>named.push(...l)}));
  assert.deepEqual(asked,[],'no file is asked for');
  assert.deepEqual(embedded,[]);
  assert.deepEqual(named,[{family:'OpenDyslexic'}],'and the caller is told which font is only named');
  assert.equal(zip.file('word/fontTable.xml'),null);
  assert.equal(Object.keys(zip.files).filter(n=>/^word\/fonts\//.test(n)).length,0);
  assert.doesNotMatch(await text(zip,'word/settings.xml'),/embedTrueTypeFonts/);
  assert.doesNotMatch(await text(zip,'[Content_Types].xml'),/odttf/);
  const styles=await text(zip,'word/styles.xml');
  assert.match(styles,/w:ascii="OpenDyslexic"/,'the name is in the file');
  assert.equal(io.parse(styles).getElementsByTagName('parsererror').length,0);
});

test('a guide with both kinds carries the pack\'s font and only names the other; the template does the same',async()=>{
  const a=await guideUsing('Lexend','OpenDyslexic');
  const named=[],embedded=[];
  const zip=await JSZip.loadAsync(await docxFromPackage(await a.sourcePackage(),JSZip,io,{entries:a.exportState().paragraphs,fonts:loader,onEmbedded:l=>embedded.push(...l),onNamed:l=>named.push(...l)}));
  assert.deepEqual([...(await text(zip,'word/fontTable.xml')).matchAll(/<w:font w:name="([^"]+)">/g)].map(m=>m[1]),['Lexend']);
  assert.deepEqual(embedded.map(e=>e.family),['Lexend']);
  assert.deepEqual(named,[{family:'OpenDyslexic'}]);
  const tNamed=[];
  const t=await JSZip.loadAsync(await docxFromPackage(await a.sourcePackage(),JSZip,io,{template:true,fonts:loader,onNamed:l=>tNamed.push(...l)}));
  assert.deepEqual(tNamed,[{family:'OpenDyslexic'}],'a styles template says it too, from its styles');
  assert.match(await text(t,'word/styles.xml'),/w:ascii="OpenDyslexic"/);
  assert.deepEqual([...(await text(t,'word/fontTable.xml')).matchAll(/<w:font w:name="([^"]+)">/g)].map(m=>m[1]),['Lexend']);
});

test('a file that uses no name-only font is told so with an empty list',async()=>{
  const named=[];
  await docxFromPackage(await demo().sourcePackage(),JSZip,io,{fonts:loader,onNamed:l=>named.push(l)});
  assert.deepEqual(named,[[]]);
});

test('a font is mixed with its key as Word keeps embedded fonts, and the same call gives it back',()=>{
  const font=bytesOf('Lexend-Regular.ttf');
  const guid='{00112233-4455-6677-8899-AABBCCDDEEFF}';
  const hidden=obfuscateFont(font,guid);
  // The key is the GUID's 16 bytes in reverse order, laid over the first 32 bytes twice.
  const key=[0xFF,0xEE,0xDD,0xCC,0xBB,0xAA,0x99,0x88,0x77,0x66,0x55,0x44,0x33,0x22,0x11,0x00];
  for(let i=0;i<32;i++)assert.equal(hidden[i],font[i]^key[i%16]);
  assert.deepEqual(hidden.slice(32),font.slice(32),'only the first 32 bytes change');
  assert.notDeepEqual(hidden.slice(0,32),font.slice(0,32));
  assert.deepEqual(obfuscateFont(hidden,guid),font);
  assert.throws(()=>obfuscateFont(font,'not a guid'),/GUID/);
  assert.match(newGuid(),/^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/);
  assert.notEqual(newGuid(),newGuid());
});

// ---- inside a Word file

const loader=async families=>{
  const out={};
  for(const family of families){out[family]={};for(const [style,file] of Object.entries(packFont(family).files))out[family][style]=bytesOf(file);}
  return out;
};
async function guideUsing(...families){
  const a=demo();
  const edits={overrides:{},custom:[]};
  families.forEach((f,i)=>{edits.overrides[['OSGCueBody','OSGNotes'][i]]={font:{name:f}};});
  await a.setStyleEdits(edits);
  return a;
}
const text=async(zip,n)=>zip.file(n).async('string');

test('a guide that uses one of the fonts carries it inside the Word file: font table, scrambled files, relationships, content types and the setting',async()=>{
  const a=await guideUsing('Lexend');
  const seen=[];
  const zip=await JSZip.loadAsync(await docxFromPackage(await a.sourcePackage(),JSZip,io,{entries:a.exportState().paragraphs,fonts:loader,onEmbedded:l=>seen.push(...l)}));
  const table=await text(zip,'word/fontTable.xml');
  assert.equal(io.parse(table).getElementsByTagName('parsererror').length,0);
  assert.match(table,/<w:font w:name="Lexend">/);
  assert.match(table,/<w:family w:val="swiss"\/>/);
  const keys=[...table.matchAll(/<w:(embedRegular|embedBold) r:id="(rId\d+)" w:fontKey="(\{[0-9A-F-]{36}\})"/g)];
  assert.deepEqual(keys.map(k=>k[1]),['embedRegular','embedBold'],'Lexend has a regular and a bold');
  const rels=await text(zip,'word/_rels/fontTable.xml.rels');
  for(const [,kind,rid,guid] of keys){
    const target=new RegExp(`Id="${rid}" Type="[^"]*/relationships/font" Target="(fonts/font\\d+\\.odttf)"`).exec(rels)?.[1];
    assert.ok(target,`${rid} points at a font file`);
    const stored=new Uint8Array(await zip.file(`word/${target}`).async('uint8array'));
    const original=bytesOf(packFont('Lexend').files[kind==='embedRegular'?'regular':'bold']);
    assert.notDeepEqual(stored.slice(0,32),original.slice(0,32),'the first bytes are scrambled');
    assert.deepEqual(obfuscateFont(stored,guid),original,'and the key gives the font back, byte for byte');
  }
  assert.equal(new Set(keys.map(k=>k[3])).size,keys.length,'each file has a key of its own');
  const types=await text(zip,'[Content_Types].xml');
  assert.match(types,/Extension="odttf" ContentType="application\/vnd\.openxmlformats-officedocument\.obfuscatedFont"/);
  assert.match(types,/PartName="\/word\/fontTable\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.fontTable\+xml"/);
  assert.match(await text(zip,'word/_rels/document.xml.rels'),/relationships\/fontTable" Target="fontTable\.xml"/);
  assert.match(await text(zip,'word/settings.xml'),/<w:zoom w:percent="100"\/><w:embedTrueTypeFonts\/><w:defaultTabStop/,'in the place the settings schema gives it');
  assert.equal(Object.values(zip.files).filter(f=>f.dir).length,0);
  assert.deepEqual(seen.map(s=>[s.family,s.styles]),[['Lexend',['regular','bold']]]);
  assert.equal(seen[0].bytes,bytesOf('Lexend-Regular.ttf').length+bytesOf('Lexend-Bold.ttf').length);
});

test('a family with four styles carries all four, a serif one is marked as one, and each font is carried once however many styles use it',async()=>{
  const a=await guideUsing('Source Serif 4','Source Serif 4');
  const zip=await JSZip.loadAsync(await docxFromPackage(await a.sourcePackage(),JSZip,io,{fonts:loader}));
  const table=await text(zip,'word/fontTable.xml');
  assert.equal((table.match(/<w:font /g)||[]).length,1);
  assert.match(table,/<w:family w:val="roman"\/>/);
  assert.deepEqual([...table.matchAll(/<w:(embed\w+) /g)].map(m=>m[1]),['embedRegular','embedBold','embedItalic','embedBoldItalic'],'in the order the schema gives');
  assert.equal(Object.keys(zip.files).filter(n=>/^word\/fonts\//.test(n)).length,4);
});

test('two fonts are two entries of the table',async()=>{
  const a=await guideUsing('Atkinson Hyperlegible','Lexend');
  const zip=await JSZip.loadAsync(await docxFromPackage(await a.sourcePackage(),JSZip,io,{fonts:loader}));
  assert.deepEqual([...(await text(zip,'word/fontTable.xml')).matchAll(/<w:font w:name="([^"]+)">/g)].map(m=>m[1]),['Atkinson Hyperlegible','Lexend']);
});

test('a styles template carries the fonts its styles use too',async()=>{
  const a=await guideUsing('Lexend');
  const zip=await JSZip.loadAsync(await docxFromPackage(await a.sourcePackage(),JSZip,io,{template:true,fonts:loader}));
  assert.ok(zip.file('word/fontTable.xml'));
  assert.match(await text(zip,'[Content_Types].xml'),/wordprocessingml\.template\.main\+xml/);
});

test('a file that uses none of the fonts, or asks for none to be carried, has none in it',async()=>{
  const plain=demo();
  const none=await JSZip.loadAsync(await docxFromPackage(await plain.sourcePackage(),JSZip,io,{fonts:loader}));
  assert.equal(none.file('word/fontTable.xml'),null);
  assert.doesNotMatch(await text(none,'word/settings.xml'),/embedTrueTypeFonts/);
  assert.doesNotMatch(await text(none,'[Content_Types].xml'),/odttf/);
  const a=await guideUsing('Lexend');
  const off=await JSZip.loadAsync(await docxFromPackage(await a.sourcePackage(),JSZip,io,{}));
  assert.equal(off.file('word/fontTable.xml'),null,'no loader, no fonts: the file only names the font');
  assert.match(await text(off,'word/styles.xml'),/w:ascii="Lexend"/,'and still names it');
  const missing=await JSZip.loadAsync(await docxFromPackage(await a.sourcePackage(),JSZip,io,{fonts:async()=>({})}));
  assert.equal(missing.file('word/fontTable.xml'),null,'a font that could not be found is left out, not half written');
});

test('fonts too large to carry are refused in words',async()=>{
  const a=await guideUsing('Lexend');
  const big=async()=>({Lexend:{regular:new Uint8Array(EMBED_LIMITS.fontBytes+1)}});
  await assert.rejects(docxFromPackage(await a.sourcePackage(),JSZip,io,{fonts:big}),/too large to carry inside a Word file/);
});

// ---- from the program's server

async function withServer(fn){
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'osg-fonts-'));
  const s=await startProjectServer({folder:path.join(dir,'g'),port:0,snapshot:{version:'0',buildId:'t',record:{},read:()=>null}});
  const get=(urlPath)=>new Promise((resolve,reject)=>{
    http.get({host:'127.0.0.1',port:s.port,path:urlPath,headers:{Host:`127.0.0.1:${s.port}`}},res=>{const parts=[];res.on('data',c=>parts.push(c));res.on('end',()=>resolve({status:res.statusCode,headers:res.headers,body:Buffer.concat(parts)}));}).on('error',reject);
  });
  try{await fn(get);}finally{await s.close();fs.rmSync(dir,{recursive:true,force:true});}
}

test('the server gives out the fonts and their licenses by plain name, as the right type, with the program\'s policy, and nothing else from that folder',async()=>{
  await withServer(async get=>{
    const font=await get('/program/fonts/Lexend-Regular.ttf');
    assert.equal(font.status,200);
    assert.equal(font.headers['content-type'],'font/ttf');
    assert.match(font.headers['content-security-policy'],/default-src 'self'/);
    assert.deepEqual(new Uint8Array(font.body),bytesOf('Lexend-Regular.ttf'),'the bytes are the file\'s');
    const license=await get('/program/fonts/OFL-Lexend.txt');
    assert.equal(license.status,200);assert.match(license.headers['content-type'],/text\/plain/);
    assert.match(license.body.toString(),/SIL OPEN FONT LICENSE/);
    for(const bad of ['/program/fonts/','/program/fonts/..%2Fstyle.css','/program/fonts/%2e%2e/package.json','/program/fonts/nope.ttf','/program/fonts/Lexend-Regular.exe','/program/fonts/sub/Lexend-Regular.ttf','/program/fonts/Lexend%20Regular.ttf','/program/other/Lexend-Regular.ttf']){
      assert.equal((await get(bad)).status,404,`${bad} is not served`);
    }
  });
});
