// The page's side of a guide file (src/guide-project-web.mjs): every read and save carries the token the program gave the page when the guide was
// opened or made, and a page that is out of date (another guide was opened in another tab or window) is told so and stops saving for good.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {webProjectStore} from '../src/guide-project-web.mjs';
import {createProjectAdapter} from '../src/guide-project-adapter.mjs';
import {createDemoAdapter} from '../src/guide-word.mjs';

const reply=(status,body='')=>({status,ok:status>=200&&status<300,text:async()=>body,clone(){return reply(status,body);}});
const STALE=JSON.stringify({error:'This guide was closed or changed in another window, so this page no longer saves it. Reload the page to go on.',code:'stale-guide'});

test('every read and save carries the guide\'s token, and nothing is sent when there is none',async()=>{
  const seen=[];
  const fetchFn=async(url,options)=>{seen.push([url,options.method||'GET',options.headers]);return reply(options.method==='PUT'?200:204);};
  let token='abc123';
  const store=webProjectStore({fetchFn,token:()=>token});
  assert.equal(await store.read(),null);
  await store.write('text');
  assert.deepEqual(seen.map(([,m,h])=>[m,h['X-Guide-Token']]),[['GET','abc123'],['PUT','abc123']]);
  token=null;seen.length=0;
  await store.read();await store.write('text');
  assert.ok(seen.every(([,,h])=>!('X-Guide-Token' in h)),'no token, no header');
  seen.length=0;
  await webProjectStore({fetchFn}).write('x');
  assert.ok(!('X-Guide-Token' in seen[0][2]),'a store given no token at all sends none (the --file way)');
});

test('an out-of-date page is told so, with an error that says it must stop; any other refusal is an ordinary error',async()=>{
  const store=status=>webProjectStore({fetchFn:async()=>reply(status,status===409?STALE:'No'),token:()=>'t'});
  await assert.rejects(store(409).write('x'),e=>e.stale===true&&/closed or changed in another window/.test(e.message));
  await assert.rejects(store(409).read(),e=>e.stale===true);
  await assert.rejects(webProjectStore({fetchFn:async()=>reply(409,JSON.stringify({error:'No guide is open.'}))}).write('x'),e=>!e.stale&&/No guide is open/.test(e.message),'another 409 is not a stale page');
  await assert.rejects(store(500).write('x'),e=>!e.stale);
  await assert.rejects(webProjectStore({fetchFn:async()=>reply(409,'not json')}).write('x'),e=>!e.stale);
});

test('a page that is told it is out of date stops saving: it says so, marks nothing to retry, and refuses further changes',async()=>{
  const states=[];
  let writes=0;
  const document=createProjectAdapter(createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer}),{
    read:async()=>null,
    write:async()=>{writes++;throw Object.assign(new Error('This guide was closed or changed in another window, so this page no longer saves it. Reload the page to go on.'),{stale:true});},
    soon:fn=>{fn();return 0;},cancel:()=>{},
    onStatus:s=>states.push(s)
  });
  await document.ready;
  await document.setCustomCues(['Pause here']);
  await document.flush();
  assert.equal(document.status().state,'error');
  assert.match(document.status().message,/closed or changed in another window.*Nothing more is saved from this page\./);
  assert.equal(document.blocked(),true);
  assert.throws(()=>document.setCustomCues(['More']),/Nothing more is saved from this page/,'further changes are refused rather than lost');
  await document.flush();
  assert.equal(writes,1,'it did not try again');
});

test('a save that fails for any other reason is still tried again with the next change',async()=>{
  let writes=0;
  const document=createProjectAdapter(createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer}),{
    read:async()=>null,
    write:async()=>{writes++;if(writes===1)throw new Error('the disk said no');},
    soon:fn=>{fn();return 0;},cancel:()=>{}
  });
  await document.ready;
  await document.setCustomCues(['One']);await document.flush();
  assert.equal(document.status().state,'error');
  assert.equal(document.blocked(),false);
  await document.setCustomCues(['Two']);await document.flush();
  assert.equal(document.status().state,'saved');
});
