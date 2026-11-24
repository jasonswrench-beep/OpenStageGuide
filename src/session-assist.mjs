// Optional help from an AI model with a session (milestone 4 of the product spec of 2026-11-22: "Optional AI assistance can propose activities and revisions from author-selected
// context. Show a preview to accept, edit, or discard, and retain the current plan on failure. Manual authoring and export are independent of an AI account.")
//
// This program has no AI model set up and sends nothing anywhere, so what is here is the part that does not depend on one: how a proposal is asked for, checked, and turned into the same
// list of changes the slide import shows (src/session-slide-import.mjs), for the author to tick, change or leave out before anything is written. A "provider" is a function the program
// would be given (a model it runs or a service the author chooses); there is none today, and asking for help says so in words. A provider that fails, answers with something that is not
// a list of proposals, or proposes what a proposal may not hold, changes nothing: the guide stays as it was and the author is told why.
import {activityBlocks} from './session-record.mjs';
import {ACTIVITY_PARTS} from './guide-cues.mjs';

export const NO_PROVIDER='No AI helper is set up in this program, so nothing was sent anywhere and nothing was changed. Everything else in the program works without one.';
const LABELS=Object.fromEntries(ACTIVITY_PARTS.map(p=>[p.id,p.label]));
// What a proposal may change: the words of the parts an author edits, never the source line (it says where an activity came from).
const EDITABLE=new Set(ACTIVITY_PARTS.map(p=>p.id).filter(id=>id!=='source'));
const LIMITS=Object.freeze({proposals:30,words:4000});
const oneLine=t=>String(t??'').replace(/\s+/g,' ').trim();

/**
 * What the helper is told: only what the author selected (the activities named in `keys`, or all), as plain text, and what is asked of it. Nothing else of the guide is in it.
 * @param {any[]} paras @param {{keys?:string[],request:string}} options
 */
export function assistContext(paras,{keys=[],request}){
  const ask=oneLine(request);
  if(!ask)throw new Error('Say what you would like help with first.');
  const chosen=activityBlocks(paras).filter(a=>!keys.length||keys.includes(a.key));
  if(!chosen.length)throw new Error('Choose at least one activity to give the helper.');
  return {request:ask.slice(0,500),activities:chosen.map(a=>({key:a.key,title:a.title,minutes:a.minutes,parts:Object.fromEntries(Object.entries(a.parts).filter(([id])=>id!=='source'))}))};
}

/** The answer of a provider, checked: a list of {key, part, to} for activities that were given and parts that may be changed. Throws, in words, what is wrong with it. */
export function checkProposals(answer,context){
  const list=Array.isArray(answer)?answer:answer&&Array.isArray(answer.proposals)?answer.proposals:null;
  if(!list)throw new Error('The helper did not answer with a list of proposals, so nothing was changed.');
  if(list.length>LIMITS.proposals)throw new Error(`The helper proposed ${list.length} changes, which is more than can be looked at together (${LIMITS.proposals}), so nothing was changed.`);
  const given=new Map(context.activities.map(a=>[a.key,a]));
  return list.map((p,i)=>{
    if(!p||typeof p!=='object')throw new Error(`Proposal ${i+1} is not in a form this program can read, so nothing was changed.`);
    const a=given.get(p.key);
    if(!a)throw new Error(`Proposal ${i+1} is for an activity that was not given to the helper, so nothing was changed.`);
    if(!EDITABLE.has(p.part))throw new Error(`Proposal ${i+1} is for "${p.part}", which a proposal may not change, so nothing was changed.`);
    const to=typeof p.to==='string'?p.to.replace(/\r\n?/g,'\n').trim():'';
    if(!to)throw new Error(`Proposal ${i+1} has no words, so nothing was changed.`);
    if(to.length>LIMITS.words)throw new Error(`Proposal ${i+1} is longer than ${LIMITS.words} letters, so nothing was changed.`);
    if(p.part==='instructions'&&!to.split('\n').some(l=>l.trim()))throw new Error(`Proposal ${i+1} leaves the instructions empty, so nothing was changed.`);
    const from=String(a.parts[p.part]??'');
    return {id:`${p.key}:${p.part}`,part:p.part,label:LABELS[p.part],from,to,conflict:false,activity:{key:a.key,title:a.title}};
  }).filter(c=>c.to!==c.from);
}

/**
 * Asks the provider, checks what it answers, and gives the changes to show. Never throws and never changes the guide: `{ok:false,message}` says why there is nothing to show.
 * @param {((context:object)=>Promise<unknown>)|null|undefined} provider
 */
export async function proposeChanges(provider,paras,options){
  let context;
  try{context=assistContext(paras,options);}catch(e){return {ok:false,message:e.message};}
  if(typeof provider!=='function')return {ok:false,message:NO_PROVIDER};
  let answer;
  try{answer=await provider(structuredClone(context));}
  catch(e){return {ok:false,message:`The helper could not be reached or failed${e&&e.message?` (${oneLine(e.message).slice(0,120)})`:''}. Your session is as it was.`};}
  try{
    const changes=checkProposals(answer,context);
    return changes.length?{ok:true,changes}:{ok:false,message:'The helper had nothing to change. Your session is as it was.'};
  }catch(e){return {ok:false,message:e.message};}
}
/** The same changes with the author's own words in place of the helper's, for those the author edited before accepting (`edits` is {id: words}). */
export const withEdits=(changes,edits={})=>changes.map(c=>Object.hasOwn(edits,c.id)&&String(edits[c.id]).trim()?{...c,to:String(edits[c.id]).replace(/\r\n?/g,'\n').trim()}:c);
