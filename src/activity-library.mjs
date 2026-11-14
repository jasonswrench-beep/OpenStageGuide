// The activity library: ready-made activities to put into a guide, and the author's own. An entry is an Activity written as labelled
// parts (objective, materials, preparation, instructions, case study), with its minutes and any facilitator notes. Putting one into
// the guide makes the same block Add Activity makes, so everything that reads an Activity (Edit, the contents, the workbook) works on it.
//
// The ready-made ones are plain, general-purpose training activities, written for this tool. The author's own are kept in the browser's
// local storage under STORE_KEY (in the program, in its settings file beside the guides: src/program-settings.mjs): they stay on this
// computer and are there in every guide.
import {ACTIVITY_PARTS} from './guide-cues.mjs';

export const STORE_KEY='osg.activityLibrary.v1';
export const MAX_MINE=50;
const FIELD_LIMIT=4000,TITLE_LIMIT=120;
const PART_IDS=ACTIVITY_PARTS.map(p=>p.id);

const entry=e=>Object.freeze({notes:'',caseStudy:'',materials:'',preparation:'',...e});

export const STARTER_ACTIVITIES=Object.freeze([
  entry({id:'opening-round',kind:'Start',title:'Opening round',minutes:10,
    summary:'Everyone says who they are and one thing they hope to get from the session.',
    objective:'Learners introduce themselves and say what they want from the session, so the facilitator can speak to those hopes.',
    materials:'A flip chart or whiteboard and a marker.',
    preparation:'Write the two questions on the board before learners arrive. Decide who will speak first.',
    instructions:'Ask each person to say their name and their role.\nAsk each person to name one thing they hope to take from the session.\nWrite the hopes on the flip chart as people say them.\nRead the list back. Say which hopes the session will meet and which it will not.',
    notes:'Keep each turn to about thirty seconds. Go first yourself to show the length. With more than twenty people, use tables of four to six and collect one hope from each table.'}),
  entry({id:'think-pair-share',kind:'Discuss',title:'Think, pair, share',minutes:10,
    summary:'A question answered alone, then with a partner, then with the room.',
    objective:'Learners form an answer on their own, test it with a partner, and hear how others answered.',
    materials:'The question on a slide or a flip chart. A timer.',
    preparation:'Write one open question that has more than one good answer.',
    instructions:'Show the question. Ask everyone to think about it quietly for one minute and jot down an answer.\nAsk learners to turn to a partner and compare answers for three minutes.\nCall time. Ask three or four pairs to share what they said.\nSum up what the answers had in common and where they differed.',
    notes:'Do not take answers during the quiet minute. If a pair finishes early, ask them to find one way their answers differ.'}),
  entry({id:'case-discussion',kind:'Discuss',title:'Small-group case discussion',minutes:20,
    summary:'Groups read a short case and answer three questions about it.',
    objective:'Learners apply the session\'s ideas to a realistic case and explain their reasoning to others.',
    materials:'One printed copy of the case for each group. A flip chart sheet and marker for each group.',
    preparation:'Print the case. Form groups of three to five. Write the three questions on a slide.',
    instructions:'Form groups of three to five. Give each group the case.\nAsk each group to read the case and answer the three questions on the slide.\nTell groups to write their answers on the flip chart sheet. Allow twelve minutes.\nAsk each group to share one answer and the reason for it. Allow about a minute each.\nAsk what the groups chose differently, and why.',
    caseStudy:'Dana leads a team of six. At Monday\'s meeting, Sam interrupts Priya twice while she explains a delay. Priya stops talking. Dana says nothing and moves to the next item. After the meeting, two people tell Dana that Sam does this often.\nQuestions: What did Dana\'s silence say to the team? What could Dana do in the next five minutes? What could Dana do this week?',
    notes:'Replace the case with one from your learners\' own work if you can. Walk between groups and listen; step in only if a group is stuck for more than a minute.'}),
  entry({id:'role-play-triads',kind:'Practice',title:'Role play in threes',minutes:25,
    summary:'Three rounds in groups of three: one speaks, one responds, one observes.',
    objective:'Learners practise the skill in a safe setting and get specific feedback from an observer.',
    materials:'A scenario card for each group. An observer checklist for each person. A timer.',
    preparation:'Write three short scenarios that match the skill. Print the cards and the checklist. Mark the corner of each checklist with a letter so groups know who observes first.',
    instructions:'Form groups of three. Give each group a scenario card and the checklists.\nExplain the roles. One person speaks, one person responds, and one person observes with the checklist.\nRun the first round for five minutes. Call time.\nThe observer gives feedback for two minutes: one thing that worked and one thing to try next time.\nSwap roles and run two more rounds in the same way.\nBring everyone back. Ask what they noticed and what they will use.',
    notes:'Say plainly that the scenarios are practice, not a test. Keep a few spare scenarios ready for groups that finish early. If the numbers do not divide by three, make one group of four with two observers.'}),
  entry({id:'gallery-walk',kind:'Discuss',title:'Gallery walk',minutes:20,
    summary:'Groups post their ideas on the wall, then everyone walks round and reads them.',
    objective:'Learners compare ideas from several groups and add to them.',
    materials:'Flip chart paper and markers for each group. Sticky notes or a second marker color. Tape.',
    preparation:'Choose a question each group will answer. Clear wall space and put up one sheet for each group.',
    instructions:'Form groups. Give each group one question and a sheet of paper.\nAsk each group to write their answer on the sheet in large letters. Allow six minutes.\nAsk groups to post their sheets on the wall.\nAsk every group to walk to the next sheet, read it, and add a note or a question. Change sheets every two minutes until each group is back at its own.\nAsk groups to read what was added to their sheet. Take one comment from each group.',
    notes:'Play no music; people need to be able to talk. If a group is slow to start, ask them for one example before they write.'}),
  entry({id:'pulse-check',kind:'Check',title:'Pulse check',minutes:5,
    summary:'A quick show of hands or fingers to see where the room stands.',
    objective:'The facilitator learns how confident learners are, and learners see that others feel the same.',
    materials:'A slide with the question and a scale from 1 to 5.',
    preparation:'Write one question, for example: How confident are you that you could do this on Monday? Decide what you will do if most answers are low.',
    instructions:'Show the question and the scale. Say that 1 is not at all and 5 is completely.\nAsk everyone to hold up fingers for their answer at the same time.\nSay out loud what you see, for example: most of you are at 3 or 4.\nAsk one person with a low number and one with a high number what is behind their answer.',
    notes:'Ask everyone to answer together so people do not copy each other. Do not argue with a number. If most of the room is low, go back over the last point before you move on.'}),
  entry({id:'teach-back',kind:'Check',title:'Teach-back',minutes:15,
    summary:'Partners explain an idea to each other in their own words.',
    objective:'Learners show that they can explain the idea correctly, and find out what they still do not understand.',
    materials:'A one-line description of the idea for each pair, on a slide or a card.',
    preparation:'Choose the idea. Write the three points a correct explanation must include.',
    instructions:'Form pairs. Name person A and person B.\nAsk person A to explain the idea to person B in their own words, in about two minutes, without reading the card.\nAsk person B to say it back in a sentence, then say what was unclear.\nSwap and repeat with the next idea or the same one.\nBring the room back. Show the three points and ask which ones pairs left out.',
    notes:'Listen to a few pairs. If the same point is missing in many of them, teach it again briefly. Do not correct pairs one at a time in front of the room.'}),
  entry({id:'one-action',kind:'Close',title:'Reflection and one action',minutes:10,
    summary:'Each person writes one thing they will do, and by when, and tells a partner.',
    objective:'Learners decide on one specific action to take after the session and commit to it.',
    materials:'An action card or a sheet of paper for each learner. Pens.',
    preparation:'Print the card with three lines: What I will do. When I will do it. What might get in the way.',
    instructions:'Give each person a card. Ask them to think about the whole session for one minute.\nAsk them to fill in the card with one action they will take in the next two weeks.\nAsk them to tell a partner what they wrote and when they will do it.\nAsk partners to say one thing that could get in the way, and how to deal with it.\nAsk learners to keep the card where they will see it.',
    notes:'Insist on one action. Several actions usually mean none gets done. Offer to send a reminder if the course has a follow-up message.'})
]);

const lines=t=>String(t??'').split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
const oneLine=t=>String(t??'').replace(/\s+/g,' ').trim();

// An entry as the arguments of cueBlock (guide-ooxml.mjs): an Activity written as labelled parts.
export function blockFromEntry(e){
  const parts=Object.fromEntries(PART_IDS.map(id=>[id,String(e[id]??'')]));
  return {cue:'activity',text:oneLine(e.title),minutes:e.minutes?String(e.minutes):'',audience:'everyone',parts,notes:String(e.notes??'')};
}

// What the author saved: an Activity block read from the document (readBlock in guide-model.mjs) made into an entry, or a plain reason it cannot be.
export function entryFromBlock(b){
  if(!b||b.cue!=='activity')throw new Error('Put your cursor in an Activity block, then choose Save.');
  if(b.hasPicture)throw new Error('This activity has a picture, and a saved activity cannot hold one. Insert the saved copy and add the picture again.');
  const all=lines(b.text);
  const title=oneLine(all[0]).slice(0,TITLE_LIMIT);
  if(!title)throw new Error('This activity has no title to save it under. Give it a title first.');
  let parts;
  if(b.parts)parts=b.parts;
  else{
    if(all.length<2)throw new Error('This activity has no steps to save. Write what learners do under its title first, one step to a line.');
    parts={instructions:all.slice(1).join('\n')};
  }
  if(!lines(parts.instructions).length)throw new Error('This activity has no instructions to save. Write what learners do, one step to a line.');
  const out={title,minutes:Number(b.minutes)>0?Number(b.minutes):0,notes:String(b.notes??'')};
  for(const id of PART_IDS)out[id]=String(parts[id]??'');
  for(const id of [...PART_IDS,'notes'])if(out[id].length>FIELD_LIMIT)throw new Error(`The ${id==='caseStudy'?'case study':id} is longer than ${FIELD_LIMIT} characters, which is more than a saved activity holds.`);
  return out;
}

// The author's own, as read back from storage: anything that is not whole is left out rather than half-used.
export function cleanMine(raw){
  if(!Array.isArray(raw))return [];
  const out=[],seen=new Set();
  for(const r of raw){
    if(!r||typeof r!=='object')continue;
    const title=oneLine(r.title).slice(0,TITLE_LIMIT),id=typeof r.id==='string'?r.id.slice(0,60):'';
    if(!title||!id||seen.has(id)||!lines(r.instructions).length)continue;
    const e={id,kind:'Yours',title,minutes:Number(r.minutes)>0?Math.min(Number(r.minutes),600):0,summary:oneLine(r.summary).slice(0,200)||'Saved from one of your guides.'};
    for(const f of [...PART_IDS,'notes'])e[f]=String(r[f]??'').slice(0,FIELD_LIMIT);
    seen.add(id);out.push(Object.freeze(e));
    if(out.length>=MAX_MINE)break;
  }
  return out;
}

// Saves an entry among the author's own: one with the same title is replaced (saving it again updates it), a new one goes first.
export function addMine(mine,saved,{now=()=>Date.now()}={}){
  const title=oneLine(saved.title).toLowerCase(),same=mine.find(m=>oneLine(m.title).toLowerCase()===title);
  if(!same&&mine.length>=MAX_MINE)throw new Error(`You have saved ${MAX_MINE} activities, which is the most. Remove one first.`);
  const e={...saved,id:same?same.id:`mine-${now().toString(36)}-${mine.length}`,kind:'Yours',summary:'Saved from one of your guides.'};
  return {list:cleanMine([e,...mine.filter(m=>m!==same)]),replaced:!!same,id:e.id};
}
export const removeMine=(mine,id)=>mine.filter(m=>m.id!==id);

// Words to match a search against: the title, the kind and the summary, in any case.
export function matches(e,query){
  const q=oneLine(query).toLowerCase();
  if(!q)return true;
  const hay=`${e.title} ${e.kind} ${e.summary}`.toLowerCase();
  return q.split(' ').every(w=>hay.includes(w));
}
