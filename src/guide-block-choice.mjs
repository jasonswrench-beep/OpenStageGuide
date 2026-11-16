// Adding a block in place: Insert > Block... on the guide's own page. The author chooses a kind of block (Say, Ask, Activity, Debrief...), and a
// block with its icon and label goes in after the block the cursor is in, ready to type its words in the page itself. Nothing here needs the tools
// beside the page: the same builders those tools use (cueBlock) write the block, and this file only lists what can be chosen and builds the
// one thing the tools do not: an Activity written as labelled parts, with the parts' labels in place and their words still to type.
// A block left with no words is found by the Review (BLOCK-EMPTY, emptyBlocks in guide-model.mjs), so a block added and forgotten is not forgotten for long.
import {CUES,CUE_GROUPS,ACTIVITY_PARTS,cueById} from './guide-cues.mjs';
import {cueBlock,para,run} from './guide-ooxml.mjs';
import {STYLE_IDS} from './guide-styles.mjs';

// A picture block needs a picture file and a handout block may cite a form: Insert > Picture and the Handout tool are the way to those.
const NOT_HERE=new Set(['picture']);
export const ACTIVITY_PARTS_CHOICE='activity-parts';

/** What the chooser lists: every cue the guide has (except those that need a file), in the order of the Graphic Cues page, each with the group it is in. */
export function blockChoices(){
  const out=[];
  for(const group of CUE_GROUPS){
    for(const cue of CUES.filter(c=>c.group===group.id&&!NOT_HERE.has(c.id))){
      out.push({id:cue.id,label:cue.label,group:group.id,groupLabel:group.label,help:cue.help,audience:cue.audience});
      // Right after Activity: the same block with its parts (objective, materials, preparation, instructions, case study) to type in.
      if(cue.id==='activity')out.push({id:ACTIVITY_PARTS_CHOICE,label:'Activity with its parts',group:group.id,groupLabel:group.label,help:'An activity written as labelled parts: Objective, Materials, Preparation, Instructions and Case study, each ready to type in. A part left empty is left out when you finish.',audience:'everyone'});
    }
  }
  return out;
}
export const choiceById=id=>blockChoices().find(c=>c.id===id)||null;

/** The minutes of a block as the form gives them: nothing, or a whole number of 1 to 999. */
export function blockMinutes(value){
  const text=String(value??'').trim();
  if(!text)return 0;
  if(!/^\d{1,3}$/.test(text)||Number(text)<1)throw new Error('Give the minutes as a whole number from 1 to 999, or leave them out.');
  return Number(text);
}

/**
 * The function that builds the chosen block, for the editor to insert: it takes the insert's context and gives the block's paragraphs.
 * The block has no words yet (the author types them in the page); `minutes`, when given, are written as "(N min)" after them.
 */
export function buildChoice(id,{minutes=''}={}){
  const choice=choiceById(id);
  if(!choice)throw new Error('Choose a kind of block first.');
  const mins=blockMinutes(minutes);
  if(id===ACTIVITY_PARTS_CHOICE){
    return ctx=>cueBlock({cue:'activity',text:'',minutes:mins||undefined},ctx)+ACTIVITY_PARTS.map(p=>
      para(p.for==='facilitator'?STYLE_IDS.notes:STYLE_IDS.body,run(`${p.label}:`,STYLE_IDS.cueLabel)+run(' '))).join('');
  }
  return ctx=>cueBlock({cue:cueById(id),text:'',minutes:mins||undefined},ctx);
}
