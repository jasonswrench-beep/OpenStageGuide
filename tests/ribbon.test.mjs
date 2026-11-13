import test from 'node:test';
import assert from 'node:assert/strict';
import {registerRibbonActions} from '../src/ribbon.mjs';
import {TOOL_IDS} from '../src/guide-tools.mjs';

function host(options={}){
  const handlers={},calls=[],errors=[];
  const office={actions:{associate:(id,fn)=>handlers[id]=fn},addin:{showAsTaskpane:async()=>{calls.push('show');if(options.showError)throw options.showError;}}};
  registerRibbonActions(office,{ready:options.ready||Promise.resolve(),go:route=>calls.push(route),onError:e=>errors.push(e)});
  return {handlers,calls,errors};
}
test('each of the twenty-one ribbon buttons opens the pane on its own tool and completes its command',async()=>{
  const h=host();
  assert.deepEqual(Object.keys(h.handlers).sort(),[...TOOL_IDS].sort(),'a function for every button, and no other');
  for(const command of TOOL_IDS){
    let completed=0;await h.handlers[command]({completed(){completed++;}});
    assert.deepEqual(h.calls.splice(0),['show',command],`${command} opens the pane on the ${command} tool`);assert.equal(completed,1);
  }
});
test('the nine commands an older ribbon runs still open the tool they always meant',async()=>{
  const h=host();
  for(const command of ['openDesigner','addModule','addActivity','participantWorkbook','courseOutline','timingAndAgenda','reviewGuide','saveBackup','restoreBackup']){
    await h.handlers[command]({completed(){}});
    assert.deepEqual(h.calls.splice(0),['show',command]);
  }
});
test('cold ribbon command waits for initialization before navigating',async()=>{
  let resolve;const ready=new Promise(r=>resolve=r),h=host({ready});
  let completed=0;const pending=h.handlers.addModule({completed(){completed++;}});
  await Promise.resolve();assert.deepEqual(h.calls,[]);assert.equal(completed,0);
  resolve();await pending;assert.deepEqual(h.calls,['show','addModule']);assert.equal(completed,1);
});
test('failed pane opening reports the error and releases the Office command',async()=>{
  const problem=new Error('Pane unavailable'),h=host({showError:problem});
  let completed=0;await h.handlers.addActivity({completed(){completed++;}});
  assert.deepEqual(h.calls,['show']);assert.deepEqual(h.errors,[problem]);assert.equal(completed,1);
});
