// Objectives with Bloom's levels: the six levels of the revised taxonomy (Anderson & Krathwohl, 2001), the verbs that belong to each, and the checks that
// tell an author whether an objective says something that can be seen and measured. The checks only advise: nothing here refuses an objective, and the
// author's own words are always what is kept. The citations of the taxonomy are in design-models.mjs.

export const BLOOM_LEVELS=Object.freeze([
  {id:'remember',label:'Remember',rank:1,asks:'Can learners get the fact back?',verbs:['define','describe','identify','label','list','locate','match','name','recall','recognize','select','state']},
  {id:'understand',label:'Understand',rank:2,asks:'Can learners explain it in their own words?',verbs:['classify','compare','contrast','discuss','explain','give examples of','illustrate','infer','interpret','paraphrase','predict','summarize']},
  {id:'apply',label:'Apply',rank:3,asks:'Can learners use it in a new situation?',verbs:['apply','calculate','carry out','complete','demonstrate','implement','operate','perform','practice','prepare','solve','use']},
  {id:'analyze',label:'Analyze',rank:4,asks:'Can learners take it apart and see how the parts relate?',verbs:['attribute','break down','categorize','deconstruct','diagnose','differentiate','distinguish','examine','investigate','organize','outline','relate']},
  {id:'evaluate',label:'Evaluate',rank:5,asks:'Can learners judge it against a standard?',verbs:['appraise','assess','critique','defend','judge','justify','prioritize','rate','recommend','support','test','verify']},
  {id:'create',label:'Create',rank:6,asks:'Can learners put the parts together into something new?',verbs:['build','compose','construct','create','design','develop','devise','formulate','generate','plan','produce','write']},
]);
export const levelById=id=>BLOOM_LEVELS.find(l=>l.id===id)||null;
export const levelRank=id=>levelById(id)?.rank??0;

/** Verbs that name something nobody can see a learner do, each with verbs that could stand in for it. */
export const WEAK_VERBS=Object.freeze({
  'understand':['explain','summarize','interpret'],
  'know':['define','list','identify'],
  'learn':['describe','demonstrate','apply'],
  'appreciate':['justify','defend','recommend'],
  'be aware of':['identify','describe','recognize'],
  'be familiar with':['describe','locate','use'],
  'become familiar with':['describe','locate','use'],
  'become aware of':['identify','describe','recognize'],
  'comprehend':['explain','summarize','paraphrase'],
  'grasp':['explain','illustrate','apply'],
  'realize':['identify','explain','justify'],
  'believe':['justify','defend','support'],
  'enjoy':['select','recommend','justify'],
  'think about':['examine','compare','assess'],
  'get a feel for':['describe','demonstrate','compare'],
});

// All the verbs, longest first so "give examples of" is read before "give".
const KNOWN=[...BLOOM_LEVELS.flatMap(l=>l.verbs),...Object.keys(WEAK_VERBS)].sort((a,b)=>b.length-a.length||a.localeCompare(b));
const clean=t=>String(t??'').replace(/\s+/g,' ').trim();

/** The levels a verb belongs to (some verbs, like "compare", are used at more than one). */
export const verbLevels=verb=>BLOOM_LEVELS.filter(l=>l.verbs.includes(String(verb||'').toLowerCase())).map(l=>l.id);

// The words an objective starts with before its verb: a condition ("After the demonstration, "), who ("learners will be able to").
const LEAD=/^(?:(?:by the end of|after|following|upon|at the end of|given|during|when|once)\b[^,]*,\s*)?(?:(?:the\s+)?(?:learners?|participants?|students?|trainees?|attendees?|employees?|managers?|you|they|everyone|each\s+\w+)\s+(?:will|should|must|can|shall)(?:\s+be\s+able\s+to)?|to)\s+/i;

/** The verb an objective is built on, and the rest of it: "Learners will be able to explain the steps." gives {verb:'explain', rest:'the steps.'}. */
export function verbOf(text){
  const t=clean(text).replace(LEAD,'').replace(/^be able to\s+/i,'');
  const low=t.toLowerCase();
  const hit=KNOWN.find(v=>low===v||low.startsWith(`${v} `));
  if(hit)return {verb:hit,rest:t.slice(hit.length).trim()};
  const first=low.match(/^[a-z'-]+/)?.[0]||'';
  return {verb:first,rest:t.slice(first.length).trim()};
}

/**
 * What an objective says, and what could be better about it. Never refuses; every issue is advice.
 * @param {string} text
 * @returns {{verb:string, levels:string[], level:string|null, weak:boolean, issues:{rule:string,message:string}[], suggestions:string[]}}
 */
export function checkObjective(text){
  const t=clean(text),issues=[];
  if(!t)return {verb:'',levels:[],level:null,weak:false,issues:[{rule:'OBJ-EMPTY',message:'Write what learners will be able to do.'}],suggestions:[]};
  const {verb,rest}=verbOf(t);
  const weak=Object.hasOwn(WEAK_VERBS,verb),levels=verbLevels(verb);
  let suggestions=[];
  if(weak){
    suggestions=WEAK_VERBS[verb];
    issues.push({rule:'OBJ-WEAK-VERB',message:`"${verb}" names something nobody can see a learner do. Try ${suggestions.map(s=>`"${s}"`).join(', ')}.`});
  }else if(!levels.length){
    issues.push({rule:'OBJ-NO-VERB',message:verb?`"${verb}" is not one of the action verbs of Bloom's levels. Start with what learners will do, for example "explain", "use" or "design".`:'Start with what learners will do, for example "explain", "use" or "design".'});
  }
  if(!rest.replace(/[.\s]/g,'')&&verb)issues.push({rule:'OBJ-NO-OBJECT',message:`Say what learners will ${verb}, and, if it matters, how well or under what conditions.`});
  const words=t.split(/\s+/).length;
  if(words>45)issues.push({rule:'OBJ-LONG',message:`This objective has ${words} words. One short sentence about one thing is easier to teach, to assess and to map.`});
  // Two verbs joined by "and": two objectives, which cannot be mapped to one activity and one test.
  const two=new RegExp(`\\band\\s+(${BLOOM_LEVELS.flatMap(l=>l.verbs).sort((a,b)=>b.length-a.length).join('|')})\\b`,'i').exec(rest);
  if(two)issues.push({rule:'OBJ-TWO-VERBS',message:`"${verb}" and "${two[1].toLowerCase()}" are two things learners do. Written as two objectives, each can have its own activity and its own check.`});
  // The level of an objective is its verb's; a verb used at several levels is the lowest unless the author says otherwise.
  const level=levels.length?levels[0]:null;
  return {verb,levels,level,weak,issues,suggestions};
}

/**
 * The sentence an objective is written as, from its parts: an optional condition, the verb, what, and an optional standard.
 * @param {{audience?:string,condition?:string,verb:string,what?:string,standard?:string}} p
 */
export function composeObjective({audience='learners',condition='',verb,what='',standard=''}){
  const v=clean(verb);
  if(!v)throw new Error('Choose the verb: what learners will do.');
  const cond=clean(condition).replace(/[,.\s]+$/,''),who=clean(audience)||'learners';
  const body=[v,clean(what),clean(standard)].filter(Boolean).join(' ').replace(/[.\s]+$/,'');
  const cap=s=>`${s.charAt(0).toUpperCase()}${s.slice(1)}`;
  return cond?`${cap(cond)}, ${who} will be able to ${body}.`:`${cap(who)} will be able to ${body}.`;
}
