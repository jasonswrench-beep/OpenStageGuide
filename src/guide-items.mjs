// The assessment item types and the three assessment levels. Plain data: the pane builds its form from it, the builders
// in guide-assess-ooxml.mjs build the Word content from it, and the model reads items back by it.
// The type names follow OpenAssess (multipleChoice, multipleResponse, trueFalse, fillInBlank, shortAnswer, essay) so a
// bank question and a guide question mean the same thing.

export const ITEM_TYPES=Object.freeze([
  {id:'multiple-choice',label:'Multiple choice',keyLabel:'multiple choice',group:'check',options:true,pick:'one',answerLine:'Your answer:',
   help:'A question with a few answers to choose from, and one correct answer.'},
  {id:'multiple-response',label:'Multiple response',keyLabel:'multiple response',group:'check',options:true,pick:'many',answerLine:'Your answers:',instruction:'Choose all that apply.',
   help:'A question with a few answers to choose from, and more than one correct answer.'},
  {id:'true-false',label:'True or false',keyLabel:'true or false',group:'check',truefalse:true,answerLine:'True or false:',
   help:'A statement the learner marks true or false.'},
  {id:'fill-blank',label:'Fill in the blank',keyLabel:'fill in the blank',group:'check',answer:'Accepted answers (separate with semicolons)',answerLine:'Your answer:',
   stemHint:'Mark each blank in the sentence with [blank].',help:'A sentence with a missing word or phrase.'},
  // The key is a number, and optionally how close an answer may be (plus or minus) and its units. The learner writes the number on a line.
  {id:'numeric',label:'Numeric',keyLabel:'numeric',group:'check',numeric:true,answer:'Correct number',answerLine:'Your answer:',
   help:'A question answered with a number. The key gives the right number, and how close an answer can be and still count.'},
  {id:'short-answer',label:'Short answer',keyLabel:'short answer',group:'check',answer:'What a good answer includes',answerLine:'Your answer:',lines:3,
   help:'A question answered in a sentence or two.'},
  // Matching and rating statements are tables with a header row: the pairs and the statements are one per line in the form.
  {id:'matching',label:'Matching',keyLabel:'matching',group:'check',matching:true,instruction:'Write the letter of the best match for each item.',
   stemHint:'For example: Match each term with its meaning.',help:'Items in a table, each matched with one of a lettered list of choices.'},
  {id:'ordering',label:'Ordering',keyLabel:'ordering',group:'check',ordering:true,answerLine:'Your order, by letter:',instruction:'Put the items in order by writing their letters, first to last, on the line.',
   stemHint:'For example: Put the steps of a good reframe in order.',help:'Items the learner puts in order. They are listed mixed up; the key gives the right order.'},
  {id:'reflection',label:'Reflection or essay',keyLabel:'reflection',group:'check',answer:'Model answer, or what to look for',answerLine:'Your response:',lines:6,
   help:'A prompt answered in a paragraph or more.'},
  {id:'performance',label:'Performance task',keyLabel:'performance checklist',group:'check',performance:true,answerLine:'Your response or notes:',lines:6,
   stemHint:'Describe what the learner is to do.',help:'A task the learner carries out, with a checklist the facilitator scores it by.'},
  // The rubric is a table in the key: a row for each criterion, a column for each level, each cell saying what that level looks like.
  {id:'rubric',label:'Performance task with rubric',keyLabel:'performance rubric',group:'check',rubric:true,answerLine:'Your response or notes:',lines:6,
   stemHint:'Describe what the learner is to do.',help:'A task the learner carries out, with a rubric table the facilitator scores it by: each criterion described at each level. The rubric goes in the key; the learner gets the task and lines to write on.'},
  {id:'rating-statements',label:'Likert scale',keyLabel:null,group:'survey',rating:true,noKey:true,noNumber:true,
   stemHint:'A heading or instruction, for example: Trainer. Mark one box for each statement.',help:'Statements in a table with a check box under each rating. There is no answer key.'},
  {id:'semantic-differential',label:'Semantic differential',keyLabel:null,group:'survey',semantic:true,noKey:true,noNumber:true,
   stemHint:'For example: Mark the point between each pair of words that is closest to how the session felt.',help:'Pairs of opposite words with a scale of check boxes between them. There is no answer key.'},
  {id:'action-plan',label:'Action plan',keyLabel:null,group:'survey',action:true,noKey:true,noNumber:true,
   stemHint:'For example: What will you do differently after this training?',help:'Learners write what they will do: the opportunity, what could get in the way, what they will do and by when. There is no answer key.'},
  {id:'open-comment',label:'Open comment',keyLabel:null,group:'survey',answerLine:'Your comments:',lines:4,noKey:true,noNumber:true,
   help:'A survey question answered in the learner\'s own words. There is no answer key.'}
]);

// The rating scales for rating statements. "custom" takes the author's own labels.
export const RATING_SCALES=Object.freeze([
  {id:'agree',label:'Five-point agreement',labels:Object.freeze(['Strongly disagree','Disagree','Neutral','Agree','Strongly agree'])},
  {id:'success',label:'Five-point success',labels:Object.freeze(['Not at all','Slightly','Moderately','Mostly','Completely'])}
]);
export const scaleById=id=>RATING_SCALES.find(s=>s.id===id);
export const LIMITS=Object.freeze({pairs:10,extras:5,statements:20,minScale:2,maxScale:7,orderItems:8,minOrderItems:3,criteria:12,semanticPairs:12,actions:5,minLevels:2,maxLevels:5,rubricRows:8});
// The levels a new rubric starts with, one to a line as the form takes them: the level, then optional points after a bar.
export const DEFAULT_RUBRIC_LEVELS='Exceeds | 3\nMeets | 2\nApproaching | 1';
// A rubric line is the criterion, then a description of it for each level, separated by bars.
export const parseRubricRows=text=>parseLines(text).map(line=>line.split('|').map(s=>s.trim()));
export const SEMANTIC_POINTS=Object.freeze([5,7]);
export const ACTION_PROMPTS=Object.freeze(['Opportunity:','What could get in the way:','What I will do:','By when:']);

// The items of an ordering question, mixed up the same way every time: alphabetical, or backwards if alphabetical is already
// the right order. Reading an item back needs only the right order, because the mixed-up order follows from it.
export const scrambleOrder=items=>{
  const sorted=[...items].sort((a,b)=>a.localeCompare(b,'en',{sensitivity:'base'})||(a<b?-1:a>b?1:0));
  return sorted.every((s,i)=>s===items[i])?sorted.reverse():sorted;
};
// A checklist line is the criterion and, after a bar, optional points.
export const parseCriteria=text=>parseLines(text).map(line=>{
  const at=line.lastIndexOf('|');
  if(at<0)return {text:line,points:''};
  const points=line.slice(at+1).trim();
  return /^\d{1,3}(\.\d{1,2})?$/.test(points)?{text:line.slice(0,at).trim(),points}:{text:line,points:''};
});

// What the author types in the form's text boxes, as lists. A pair is the item and its match, split by a bar or a tab,
// so a line pasted from a spreadsheet works too.
export const parseLines=text=>String(text??'').split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
export const parsePairs=text=>parseLines(text).map(line=>{
  const at=line.search(/\t|\|/);
  return at<0?{item:line,match:''}:{item:line.slice(0,at).trim(),match:line.slice(at+1).replace(/^\s*[|\t]\s*/,'').trim()};
});
export const parseLabels=text=>String(text??'').split(/[,\n]/).map(l=>l.trim()).filter(Boolean);
// The labels a rating item uses, from its scale choice. Used by the builder and by reading an item back.
export const scaleLabels=item=>{
  const preset=scaleById(item.scale);
  return preset?[...preset.labels]:(item.labels||[]).map(l=>String(l??'').trim()).filter(Boolean);
};
// A matching item's choices: every match and every extra choice, once each, in alphabetical order so the order of the
// items does not give the answers away.
export const matchingChoices=(pairs,extras=[])=>{
  const seen=new Map();
  for(const text of [...pairs.map(p=>p.match),...extras]){
    const t=String(text??'').trim(),key=t.toLowerCase();
    if(t&&!seen.has(key))seen.set(key,t);
  }
  return [...seen.values()].sort((a,b)=>a.localeCompare(b,'en',{sensitivity:'base'})||(a<b?-1:a>b?1:0));
};
// A numeric item's key: the number, then (plus or minus the tolerance) and the units when there are any, with no closing full
// stop so the number is never mistaken for one with a decimal point. Reading the key back is the reverse.
const PLAIN_NUMBER=/^-?\d+(?:\.\d+)?$/;
export function numericKey({answer,tolerance,units}){
  const a=String(answer??'').trim(),t=String(tolerance??'').trim(),u=String(units??'').replace(/\s+/g,' ').trim();
  if(!a)throw new Error('Type the correct number for the key.');
  if(!PLAIN_NUMBER.test(a))throw new Error('The correct number must be a number without commas or spaces, like 42 or 3.5.');
  if(t&&!/^\d+(?:\.\d+)?$/.test(t))throw new Error('How close an answer can be must be a number that is not negative, like 0.5.');
  if(u.length>30)throw new Error('The units can be up to 30 characters, like kg or miles per hour.');
  return `${a}${t?` (plus or minus ${t})`:''}${u?` ${u}`:''}`;
}
export function readNumericKey(text){
  const m=/^(-?\d+(?:\.\d+)?)(?: \(plus or minus (\d+(?:\.\d+)?)\))?(?: (.+?))?\.?$/.exec(String(text??'').trim());
  return m?{answer:m[1],tolerance:m[2]??'',units:m[3]??''}:{answer:String(text??'').trim(),tolerance:'',units:''};
}
export const itemTypeById=id=>ITEM_TYPES.find(t=>t.id===id);
export const itemTypeByKeyLabel=label=>ITEM_TYPES.find(t=>t.keyLabel&&t.keyLabel===String(label||'').trim().toLowerCase());

export const LEVELS=Object.freeze([
  {id:'level-one',title:'Level One Assessment',button:'Level One: reaction survey',blurb:'Learners rate the training right after it. No right answers, no answer key.',
   intro:'Dear Participant: Please complete the following short survey about your experience with today\'s training program. We are always looking to improve our training programs, so your input is appreciated.',
   info:['Program title:','Date or dates:','Trainer or facilitator:','Your name (optional):']},
  {id:'level-two',title:'Level Two Assessment',button:'Level Two: knowledge check (quiz)',blurb:'The quiz. Questions with right answers, and an answer key for you.',
   intro:'Answer each question. Choose the best answer, or write your answer on the line.',
   info:['Your name:']},
  {id:'level-three',title:'Level Three Assessment',button:'Level Three: follow-up survey',blurb:'Learners report back some weeks later on what they used. No answer key.',
   intro:'Dear Participant: Recently you attended a training program (enter the topic here). As a follow-up to that program, please complete the following short survey about your experience. We are always looking to improve our training programs, so your input is appreciated.',
   info:['Program title:','Date or dates:','Trainer or facilitator:','Your name (optional):']}
]);
export const levelById=id=>LEVELS.find(l=>l.id===id);

export const MAX_OPTIONS=6;
