// What a guide is called on screen: its name, never the ending of the file it is kept in.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {guideName} from '../program/ui/guide-name.js';

const UI=path.join(path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..'),'program','ui');

test('the ending of the guide\'s file is not part of its name, whatever the case, and nothing else is cut',()=>{
  assert.equal(guideName('Reframing conflict workshop.guide-project.json'),'Reframing conflict workshop');
  assert.equal(guideName('Reframing conflict workshop.GUIDE-PROJECT.JSON'),'Reframing conflict workshop');
  assert.equal(guideName('Notes on .guide-project.json files.guide-project.json'),'Notes on .guide-project.json files','only the last ending goes');
  assert.equal(guideName('Plain name'),'Plain name');
  assert.equal(guideName('A.json'),'A.json');
  assert.equal(guideName('A.guide-project.json.bak'),'A.guide-project.json.bak','a backup copy is not a guide');
  assert.equal(guideName('.guide-project.json'),'.guide-project.json','a name that is only the ending is shown as it is, never as nothing');
  assert.equal(guideName(''),'');
  assert.equal(guideName(undefined),'');
});

test('the welcome page and the Open guide dialog show a guide by its name, and still ask for its file by the whole file name',()=>{
  const welcome=fs.readFileSync(path.join(UI,'welcome.js'),'utf8');
  const dialogs=fs.readFileSync(path.join(UI,'dialogs.js'),'utf8');
  assert.match(welcome,/class: "recent-name", text: guideName\(g\.name\)/);
  assert.match(dialogs,/class: "recent-name", text: guideName\(g\.name\)/);
  assert.match(welcome,/"data-guide": g\.name/,'the button keeps the file name to open');
  assert.match(welcome,/onOpenNamed\(g\.name\)/);
  assert.match(dialogs,/value: g\.name/,'and so does the choice in the dialog');
  assert.doesNotMatch(welcome,/class: "recent-name", text: g\.name/);
  assert.doesNotMatch(dialogs,/class: "recent-name", text: g\.name/);
});
