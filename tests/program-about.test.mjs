// What the About window says about the program's version, license and what it is made with, and that it agrees with the NOTICE file.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {LICENSE_LINE,MADE_WITH,ICONS_LINE,NOTICE_LINE,aboutVersion} from '../program/ui/credits.js';

const ROOT=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const notice=fs.readFileSync(path.join(ROOT,'NOTICE'),'utf8');
const pkg=JSON.parse(fs.readFileSync(path.join(ROOT,'package.json'),'utf8'));

test('the program does not call itself a prototype: not in the About line, and not in the closed section every tool page ends with',()=>{
  assert.doesNotMatch(aboutVersion({version:'0.1.0',build:'a1b2c3d4'}),/prototype/i);
  const tools=fs.readFileSync(path.join(ROOT,'public','taskpane.html'),'utf8');
  const support=/<details class="support"><summary>([^<]*)<\/summary><p>([^<]*)<\/p>/.exec(tools);
  assert.ok(support,'the closed section is there');
  assert.equal(support[1],'Support and recovery');
  assert.doesNotMatch(support[1]+support[2],/prototype|first build|test document/i);
});

test('the version line says the version and the start of the build, and says nothing it was not told',()=>{
  assert.equal(aboutVersion({version:'0.1.0',build:'a1b2c3d4'}),'Version 0.1.0, build a1b2c3d4.');
  assert.equal(aboutVersion({version:'0.1.0'}),'Version 0.1.0.','no build, no build in the line');
  assert.equal(aboutVersion({version:'0.1.0',build:'not hex!'}),'Version 0.1.0.','a build that is not a hash is left out');
  for(const bad of [null,undefined,{},{version:''},{version:'x'},{version:'1.2'},{version:'<b>1.0.0</b>'}])assert.equal(aboutVersion(bad),null,JSON.stringify(bad));
});

test('what the About window says it is made with is what NOTICE says, with the same licenses',()=>{
  assert.ok(MADE_WITH.length>=3);
  for(const c of MADE_WITH){
    for(const part of c.name.split(/,| and /).map(s=>s.trim()).filter(Boolean))assert.ok(notice.includes(part),`NOTICE names ${part}`);
    assert.ok(c.use&&c.license,`${c.name} says what it is for and its license`);
  }
  const has=(word,license)=>assert.match(notice,license,`NOTICE gives ${word} its license`);
  const by=name=>MADE_WITH.find(c=>c.name.startsWith(name));
  assert.equal(by('JSZip').license,'MIT License');has('JSZip',/JSZip[\s\S]{0,400}MIT License/);
  assert.equal(by('PDF.js'),undefined,'PDF.js is not carried by the program, so it is not said to be');
  assert.equal(by('Lucide icons').license,'ISC License');has('Lucide',/Lucide icons, ISC licen[cs]e/);
  assert.equal(by('Atkinson').license,'SIL Open Font License 1.1');has('the fonts',/SIL Open Font License, Version 1\.1/);
});

test('the license lines agree with the license, the package and NOTICE',()=>{
  assert.match(LICENSE_LINE,/Apache License 2\.0/);
  assert.equal(pkg.license,'Apache-2.0');
  assert.ok(notice.includes('Copyright 2026 Jason S. Wrench and the OpenStage Guide contributors'));
  assert.ok(LICENSE_LINE.includes('Copyright 2026 Jason S. Wrench and the OpenStage Guide contributors'));
  assert.match(ICONS_LINE,/Creative Commons Attribution-NonCommercial-ShareAlike 4\.0/);
  assert.match(notice,/Creative Commons Attribution-NonCommercial-ShareAlike 4\.0/);
  assert.match(NOTICE_LINE,/NOTICE file/);
  assert.ok(fs.existsSync(path.join(ROOT,'NOTICE')));
});

test('the About window draws the version, license and credits from this list, as text only',()=>{
  const dialogs=fs.readFileSync(path.join(ROOT,'program','ui','dialogs.js'),'utf8');
  assert.match(dialogs,/aboutVersion\(info\)/);
  assert.match(dialogs,/text: LICENSE_LINE/);
  assert.match(dialogs,/MADE_WITH\.map/);
  assert.match(fs.readFileSync(path.join(ROOT,'program','ui','app.js'),'utf8'),/aboutDialog\(doc, state\)/,'it is given what the server said');
  assert.doesNotMatch(dialogs,/innerHTML/);
});
