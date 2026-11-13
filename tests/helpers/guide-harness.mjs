// Shared set-up for the guide pane tests: the real pane module, the real taskpane.html ids and an in-memory document.
// This is not native Word.
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {dom} from './app-dom.mjs';
import {mountGuidePane} from '../../src/guide-pane.mjs';
import {createDemoAdapter} from '../../src/guide-word.mjs';

// What the author's own Word did in the first native test (2026-10-31): the last paragraph of every insert lost its style.
export const WORD_QUIRKS={lastParagraphLosesStyle:true};

// A test can hand in its own document (a project file's, say) instead of a fresh in-memory one.
export function pane(seed,{quirks,timeoutMs,storage,onTool,adapter:given}={}){
  const ui=dom(),adapter=given||createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer,quirks});
  if(seed)adapter.seed(seed);
  const p=mountGuidePane(ui.document,{adapter,soon:fn=>fn(),timeoutMs,Zip:JSZip,xml:{DOMParserClass:DOMParser,SerializerClass:XMLSerializer},...(storage?{storage}:{}),...(onTool?{onTool}:{})}),n=id=>ui.nodes.get(id);
  // The document always ends in a plain empty paragraph after an insert that reaches its end (so the cursor can leave a block); the rows leave it out so a test sees the guide itself.
  const rows=()=>{const all=adapter.snapshot(),last=all.at(-1);return (last&&last.style==='Normal'&&!String(last.text).trim()?all.slice(0,-1):all).map(r=>`${r.style}: ${r.text.trim()}`);};
  const type=(id,v)=>{n(id).value=v;};
  // The cue radios are built by the pane, so they are found through their labels rather than by a fixed id.
  const radio=cue=>n('guide-cues').children.map(l=>l.children[0]).find(i=>i.id===`guide-cue-${cue}`);
  const pick=cue=>{radio(cue).checked=true;radio(cue).onchange();};
  const itemRadio=type=>n('guide-item-types').children.map(l=>l.children[0]).find(i=>i.id===`guide-item-type-${type}`);
  // A problem can show at the top of the pane or, for the item form, just above Insert item.
  const err=()=>[n('guide-error'),n('guide-item-problem')].filter(x=>!x.hidden).map(x=>x.textContent).join(' ');
  return {ui,adapter,p,n,rows,type,pick,radio,itemRadio,err};
}
export const submit=(a,id)=>a.n(id).onsubmit({preventDefault(){}});
export async function started(title='Reframing conflict',opts){
  const a=pane(undefined,opts);await a.p.refresh();
  a.type('guide-title',title);await submit(a,'guide-start-form');
  return a;
}

export const click=(a,id)=>a.n(id).onclick();
export const startsWith=prefix=>r=>r.startsWith(prefix);
export const moveTo=async(a,matcher)=>{a.adapter.place(a.rows().findIndex(matcher));await a.p.quietRefresh();};
export async function addModule(a,title){a.type('guide-new-title',title);await click(a,'guide-add-module');assert.equal(a.err(),'');}
export async function addLesson(a,title,minutes='10'){a.type('guide-new-title',title);a.type('guide-new-minutes',minutes);await click(a,'guide-add-lesson');assert.equal(a.err(),'');}
export async function addBlock(a,cue,text,{notes='',minutes=''}={}){a.pick(cue);a.type('guide-block-text',text);a.type('guide-block-notes',notes);a.type('guide-block-minutes',minutes);await submit(a,'guide-block-form');assert.equal(a.err(),'');}
export const guideRows=a=>a.rows().slice(a.rows().findIndex(startsWith('OSG Module')));
