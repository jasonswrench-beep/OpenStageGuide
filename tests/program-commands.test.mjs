// The program's menus (program/ui/commands.js): every command in one list, built from the same tools the add-in's ribbon is.
import test from 'node:test';
import assert from 'node:assert/strict';
import {buildMenus,menuEntries,TOOL_NAMES,TOOL_ICONS} from '../program/ui/commands.js';
import {searchEntries} from '../program/ui/commandsearch.js';
import {UI_ICONS} from '../program/ui/ui-icons.js';
import {GROUPS,TOOLS,TOOL_IDS} from '../src/guide-tools.mjs';

function make(){
  const calls=[];
  let guide=false;
  const ctx={hasGuide:()=>guide,openTool:id=>calls.push(['tool',id]),updateGuide:()=>calls.push(['update']),newGuide:()=>calls.push(['new']),openGuide:()=>calls.push(['open']),
    closeGuide:()=>calls.push(['close']),search:()=>calls.push(['search']),keys:()=>calls.push(['keys']),about:()=>calls.push(['about'])};
  return {...buildMenus(ctx),calls,setGuide:v=>{guide=v;}};
}
const items=menus=>menus.flatMap(m=>m.items.filter(i=>i!=='-'));

test('the menus are File, Edit, one for each group of the ribbon, and Help, in that order',()=>{
  const {menus}=make();
  assert.deepEqual(menus.map(m=>[m.id,m.label]),[['file','File'],['edit','Edit'],['guide','Guide'],['write','Write'],['bring','Bring in'],['elearning','E-learning'],['check','Check'],['make','Make'],['help','Help']]);
  assert.equal(GROUPS.length,7,'the ribbon has seven groups: six become menus and Protect is split between File and Help');
});

test('every tool of the add-in is one item in one menu, with its own name, icon and description',()=>{
  const {menus}=make();
  const all=items(menus);
  for(const id of TOOL_IDS){
    const label=TOOL_NAMES[id];
    assert.ok(label,`${id} has a name in the menus`);
    assert.ok(TOOL_ICONS[id],`${id} has an icon`);
    const found=all.filter(i=>i.label===label);
    assert.equal(found.length,1,`${label} is in exactly one menu`);
    assert.equal(found[0].title,TOOLS[id].tip,'its description is the tool\'s own tip');
  }
  assert.deepEqual(Object.keys(TOOL_NAMES).sort(),[...TOOL_IDS].sort(),'no name for a tool that does not exist');
  assert.deepEqual(Object.keys(TOOL_ICONS).sort(),[...TOOL_IDS].sort());
  assert.equal(new Set(Object.values(TOOL_NAMES)).size,TOOL_IDS.length,'no two tools share a name');
});

test('each group of the ribbon is the menu of the same tools in the same order; Protect goes to File and Help',()=>{
  const {menus}=make();
  const names=id=>menuEntries(menus,id).filter(i=>i!=='-').map(i=>i.label);
  const group=id=>GROUPS.find(g=>g.id===id).tools.map(t=>TOOL_NAMES[t]);
  assert.deepEqual(names('guide'),group('start'));
  assert.deepEqual(names('write'),group('write'));
  assert.deepEqual(names('bring'),group('bring'));
  assert.deepEqual(names('elearning'),group('elearning'));
  assert.deepEqual(names('check'),group('check'));
  assert.deepEqual(names('make'),group('make'));
  const protect=GROUPS.find(g=>g.id==='protect').tools.map(t=>TOOL_NAMES[t]);
  assert.ok(protect.every(n=>names('file').includes(n)||names('help').includes(n)),'Save and Restore a backup are in File, Help is in Help');
  assert.deepEqual(names('file').slice(0,3),['New guide...','Open guide...','Close guide']);
});

test('a tool opens only when a guide is open; making, opening and closing know their own state',()=>{
  const m=make();
  const tools=items(m.menus).filter(i=>Object.values(TOOL_NAMES).includes(i.label));
  assert.equal(tools.length,TOOL_IDS.length);
  assert.ok(tools.every(i=>i.enabled()===false),'no tool can be used with no guide open');
  const [newItem,openItem,closeItem]=menuEntries(m.menus,'file');
  assert.equal(newItem.enabled?.()??true,true);assert.equal(openItem.enabled?.()??true,true);
  assert.equal(closeItem.enabled(),false);
  m.setGuide(true);
  assert.ok(tools.every(i=>i.enabled()===true));
  assert.equal(closeItem.enabled(),true);
});

test('choosing a tool asks for that tool, and the other commands do their own things',()=>{
  const m=make();m.setGuide(true);
  for(const id of TOOL_IDS)items(m.menus).find(i=>i.label===TOOL_NAMES[id]).run();
  assert.deepEqual(m.calls.map(c=>c[1]),TOOL_IDS);
  m.calls.length=0;
  const byLabel=l=>items(m.menus).find(i=>i.label===l);
  byLabel('New guide...').run();byLabel('Open guide...').run();byLabel('Close guide').run();byLabel('Search for a command...').run();byLabel('Keyboard keys').run();byLabel('About OpenStage Guide').run();
  assert.deepEqual(m.calls,[['new'],['open'],['close'],['search'],['keys'],['about']]);
});

test('the search and the keys have their suite keys, and the navy bar has three buttons for what is done all the time',()=>{
  const m=make();
  const help=menuEntries(m.menus,'help');
  assert.equal(help.find(i=>i.label==='Search for a command...').keys,'Ctrl+K');
  assert.equal(help.find(i=>i.label==='Keyboard keys').keys,'Ctrl+/');
  assert.deepEqual(m.quick.map(q=>q.label),['Review guide','Update guide','Make workbook']);
  m.quick[0].run();m.quick[1].run();m.quick[2].run();
  assert.deepEqual(m.calls,[['tool','reviewGuide'],['update'],['tool','participantWorkbook']]);
  assert.ok(m.quick.length<=5,'at most five buttons on the bar');
});

test('every item has an icon the program has, and the command search can find each command once',()=>{
  const {menus,quick}=make();
  for(const i of [...items(menus),...quick]){
    assert.ok(i.icon&&UI_ICONS[i.icon],`${i.label} has an icon that exists (${i.icon})`);
    assert.ok(i.title&&i.title.length>8,`${i.label} says what it does`);
  }
  const found=searchEntries(menus);
  assert.equal(found.length,items(menus).length,'each item has a function of its own, so the search lists it once');
});

test('the Edit menu: Undo and Redo, cut copy paste, add, kind, move and delete, each with its keys, dimmed when it cannot be done, and doing what the editor is told',()=>{
  const told=[];
  let state={canUndo:true,canRedo:false,hasRows:true},guide=true;
  const ctx={hasGuide:()=>guide,openTool(){},updateGuide(){},newGuide(){},openGuide(){},closeGuide(){},search(){},keys(){},about(){},
    edit:{command:name=>told.push(name),state:()=>state}};
  const {menus}=buildMenus(ctx);
  const edit=menus.find(m=>m.id==='edit');
  const by=label=>edit.items.find(i=>i!=='-'&&i.label===label);
  assert.deepEqual(edit.items.filter(i=>i!=='-').map(i=>i.label),['Undo','Redo','Cut','Copy','Paste','Add paragraph below','Kind of paragraph...','Move up','Move down','Delete paragraph']);
  assert.deepEqual([by('Undo').keys,by('Redo').keys,by('Cut').keys,by('Copy').keys,by('Paste').keys,by('Move up').keys,by('Move down').keys,by('Delete paragraph').keys,by('Add paragraph below').keys],
    ['Ctrl+Z','Ctrl+Y','Ctrl+X','Ctrl+C','Ctrl+V','Alt+Up','Alt+Down','Delete','Ctrl+Enter']);
  for(const label of ['Undo','Redo','Cut','Copy','Paste','Add paragraph below','Kind of paragraph...','Move up','Move down','Delete paragraph']){
    const item=by(label);
    assert.ok(item.icon&&UI_ICONS[item.icon],`${label} has an icon the program has`);
    assert.ok(item.title&&item.title.length>10,`${label} says what it does`);
  }
  assert.equal(by('Undo').enabled(),true);
  assert.equal(by('Redo').enabled(),false,'nothing to redo');
  state={canUndo:false,canRedo:true,hasRows:false};
  assert.deepEqual([by('Undo').enabled(),by('Redo').enabled(),by('Delete paragraph').enabled(),by('Paste').enabled()],[false,true,false,false]);
  guide=false;state={canUndo:true,canRedo:true,hasRows:true};
  assert.equal(by('Undo').enabled(),false,'no guide, nothing to edit');
  for(const label of ['Undo','Redo','Cut','Copy','Paste','Add paragraph below','Kind of paragraph...','Move up','Move down','Delete paragraph'])by(label).run();
  assert.deepEqual(told,['undo','redo','cut','copy','paste','add','kind','moveUp','moveDown','delete']);
});

test('with no editor the Edit menu is there, dimmed, and does nothing',()=>{
  const {menus}=make();
  const edit=menus.find(m=>m.id==='edit');
  for(const i of edit.items.filter(i=>i!=='-')){assert.equal(i.enabled(),false,i.label);i.run();}
});
