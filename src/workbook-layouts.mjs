// The five layouts of the participant workbook (milestone 2 of the product spec of 2026-11-22). A layout is a choice about what the workbook adds around the guide's own learner
// content, not a different look (the five looks, "Clearline" and the rest, are how the pages are drawn; a layout is what is on them): an agenda and a response space after each
// activity (Facilitated Workshop), an observation checklist and a revision plan (Skills Practice), an evidence table and decision prompts (Case and Discussion), reflection prompts,
// a practice log and review checkpoints (Reflection and Transfer), and rating items with labelled ends and an evidence summary (Assessment and Evidence).
//
// Everything here is for learners: it is written from the activities' learner parts only (their title, grouping and output), never the facilitator's, and the page text is plain
// words in the styles the workbook already carries. Score keys and scoring words are left out unless the author asks for them ("Exclude answer keys and scoring instructions from
// participant copies unless the author explicitly includes them").
import {readGuide,readBlock,blockSpan,kindOf,cueFromText} from './guide-model.mjs';
import {makeContext,heading} from './guide-ooxml.mjs';
import {formTable,ratingGrid,writingLines,formText,checkList} from './guide-forms-ooxml.mjs';

export const WORKBOOK_LAYOUTS=Object.freeze([
  {id:'',label:'The guide as it is',fit:'Any guide',distinct:'Nothing is added to what the guide holds for learners.',note:'Your notes:'},
  {id:'facilitated',label:'Facilitated Workshop',fit:'In-person or virtual cohorts',distinct:'An agenda of the activities, and a place to write after each one.',note:'Your notes:'},
  {id:'skills',label:'Skills Practice',fit:'Practice and coached feedback',distinct:'An observation checklist and a revision plan.',note:'What I tried:'},
  {id:'case',label:'Case and Discussion',fit:'Case analysis and classroom discussion',distinct:'An evidence table and decision prompts.',note:'What the evidence says:'},
  {id:'reflection',label:'Reflection and Transfer',fit:'Spaced follow-up and field practice',distinct:'Reflection prompts, a practice log and review checkpoints.',note:'What I noticed:'},
  {id:'assessment',label:'Assessment and Evidence',fit:'Surveys, assessments and training evaluation',distinct:'Rating items with labelled ends, and an evidence summary.',note:'My evidence:'}
]);
export const layoutById=id=>WORKBOOK_LAYOUTS.find(l=>l.id===(id||''))||null;

const oneLine=t=>String(t??'').replace(/\s+/g,' ').trim();
const MAX_ROWS=30;

/**
 * The activities learners meet, in order, from the guide's paragraphs: each with its lesson, its title, and only the parts written for learners. An activity for the facilitator
 * alone is left out. (Pure: `paras` is the guide's {text, style} list, or the workbook's reading of its source.)
 */
export function participantPlan(paras){
  const out=[];
  for(const mod of readGuide(paras).modules)for(const lesson of mod.lessons){
    for(let i=lesson.index+1;i<=lesson.end;i++){
      if(kindOf(paras[i])!=='cue')continue;
      const cue=cueFromText(paras[i].text);
      if(!cue||cue.id!=='activity')continue;
      const span=blockSpan(paras,i);
      if(!span||span.start!==i)continue;
      i=span.end;
      const b=readBlock(paras,span.start);
      if(!b||b.audience==='facilitator')continue;
      const title=oneLine(String(b.text).split('\n')[0]);
      if(!title)continue;
      out.push({lesson:lesson.title,title,output:oneLine(b.parts?.output),grouping:oneLine(b.parts?.grouping)});
    }
  }
  return out;
}

const page=(title)=>heading(1,title,true);
const section=(title)=>heading(2,title);

function agenda(plan,ctx){
  if(!plan.length)return '';
  const rows=plan.slice(0,MAX_ROWS).map(a=>[a.title,a.grouping||'',a.output||'']);
  return page('Agenda')+formText('What we will do, in order. Each line says how we work and what we make.')
    +formTable({columns:[{name:'Activity',weight:3},{name:'How we work',weight:2},{name:'What we make',weight:3}],rows,ctx,caption:'Agenda'})
    +(plan.length>MAX_ROWS?formText(`The first ${MAX_ROWS} of ${plan.length} activities are listed.`):'');
}
const skills=ctx=>
  page('Observation checklist')+formText('Use this to watch one round of practice. Write the criterion, mark it when you see it, and write what you saw.')
  +formTable({columns:[{name:'Criterion',weight:3,kind:'blank'},{name:'Seen',weight:1,kind:'check'},{name:'What I saw',weight:4,kind:'blank'}],blankRows:5,ctx,caption:'Observation checklist'})
  +section('Revision plan')+formText('After the feedback, decide what to change.')
  +writingLines({label:'What I will change:',count:3})+writingLines({label:'The evidence I will look for:',count:2})+writingLines({label:'When I will try it again:',count:1});
const caseDiscussion=ctx=>
  page('Evidence table')+formText('List the evidence from the case before you decide. Write where each piece came from and what it suggests.')
  +formTable({columns:[{name:'Evidence',weight:3,kind:'blank'},{name:'Where it came from',weight:2,kind:'blank'},{name:'What it suggests',weight:3,kind:'blank'}],blankRows:6,ctx,caption:'Evidence table'})
  +section('Decision prompts')
  +writingLines({label:'The decision to make:',count:2})+writingLines({label:'The option I choose, and why:',count:3})+writingLines({label:'What would change my mind:',count:2});
const reflection=ctx=>
  page('Reflection')+formText('Take a few minutes after each practice to write down what you noticed.')
  +writingLines({label:'Three things I learned:',count:3})+writingLines({label:'Two things I will try:',count:2})+writingLines({label:'One question I still have:',count:1})
  +section('Practice log')
  +formTable({columns:[{name:'Date',weight:1,kind:'blank'},{name:'What I tried',weight:3,kind:'blank'},{name:'What happened',weight:3,kind:'blank'},{name:'What I will change',weight:3,kind:'blank'}],blankRows:5,ctx,caption:'Practice log'})
  +section('Review checkpoints')
  +checkList({items:['After one week: I tried it and wrote down what happened.','After one month: I can show the result to someone else.','After three months: I looked back at my practice log.'],columns:['Date'],ctx,caption:'Review checkpoints'});
function assessment(objectives,{scoreKey},ctx){
  const ends=['1 Strongly disagree','2','3','4','5 Strongly agree'];
  const rows=objectives.slice(0,MAX_ROWS).map(o=>[o.text,'','']);
  return page('Rating items')+formText('Mark one box in each row. 1 means Strongly disagree and 5 means Strongly agree.')
    +ratingGrid({statements:['I can explain the main idea to a colleague.','I can do this without help.','I will use this in my work.'],labels:ends,ctx,caption:'Rating items'})
    +section('How I see it')+formText('Mark the box that is closest to how you see it, from the word on the left to the word on the right.')
    +formTable({columns:[{name:'Left end',weight:2},{name:'1',weight:1,kind:'check'},{name:'2',weight:1,kind:'check'},{name:'3',weight:1,kind:'check'},{name:'4',weight:1,kind:'check'},{name:'5',weight:1,kind:'check'},{name:'Right end',weight:2}],
      rows:[['Confusing','','','','','','Clear'],['Not useful','','','','','','Useful'],['Slow','','','','','','Fast']],ctx,caption:'Semantic differential'})
    +section('Evidence summary')+formText('For each objective, write what shows that you can do it, and how sure you are, from 1 (not sure) to 5 (very sure).')
    +formTable({columns:[{name:'Objective',weight:3},{name:'What shows I can do it',weight:3,kind:'blank'},{name:'How sure I am (1 to 5)',weight:1,kind:'blank'}],rows,blankRows:rows.length?0:4,ctx,caption:'Evidence summary'})
    +(scoreKey?formText('How to score the rating items: add the numbers you marked and divide by the number of rows. A higher number means stronger agreement.'):'');
}

/**
 * What a layout adds: `front` goes after the learning objectives, `back` at the end, and `note` is the label of the lines to write on after an activity.
 * @param {string} id one of WORKBOOK_LAYOUTS @param {{plan:ReturnType<typeof participantPlan>,objectives?:{text:string}[],scoreKey?:boolean}} model
 */
export function layoutPages(id,{plan=[],objectives=[],scoreKey=false}={},ctx=makeContext()){
  const layout=layoutById(id);
  if(!layout)throw new Error('Choose one of the five workbook layouts, or the guide as it is.');
  const out={front:'',back:'',note:layout.note,allowed:[]};
  switch(layout.id){
    case 'facilitated':out.front=agenda(plan,ctx);break;
    case 'skills':out.back=skills(ctx);break;
    case 'case':out.back=caseDiscussion(ctx);break;
    case 'reflection':out.back=reflection(ctx);break;
    case 'assessment':out.back=assessment(objectives,{scoreKey},ctx);break;
    default:break;
  }
  // `allowed` is every paragraph the layout writes. They are its own plain words and the learner parts of the guide (an activity's title, grouping and output, an objective), so the
  // leak check may see each of them: a page called "Observation checklist" is not a leak of a material that happens to be called the same.
  out.allowed=paragraphTexts(out.front+out.back);
  return out;
}
const XML_ENTITIES={'&amp;':'&','&lt;':'<','&gt;':'>','&quot;':'"','&apos;':"'"};
const paragraphTexts=xml=>[...String(xml).matchAll(/<w:p>[\s\S]*?<\/w:p>/g)]
  .map(m=>[...m[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(t=>t[1].replace(/&(?:amp|lt|gt|quot|apos);/g,e=>XML_ENTITIES[e])).join('').trim())
  .filter(Boolean);
