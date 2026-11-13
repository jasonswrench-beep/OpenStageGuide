import {assessmentDraftFromDefinition} from '../../src/assessment-file.mjs';
export function assessmentFileDefinition(options=5,count=2,min=1){
 const items=Array.from({length:count},(_,i)=>({id:`item-${i+1}`,type:i%2?'semantic-differential':'likert',prompt:`Authored prompt ${i+1} — 日本語 😀\nSecond line`,reverse:i%2===1,options:Array.from({length:options},(_,j)=>({id:`option-${i+1}-${j+1}`,label:`Anchor ${j+1} é / ${j?'More':'Less'}`,code:min+j})),...(i%2?{target:'The authored course',leftEndpoint:'Unclear',rightEndpoint:'Clear',highScoreMeaning:'Greater authored clarity'}:{})}));
 return {schemaVersion:1,id:'assessment-private-test',version:2,title:'Authored PRIVATE definition é 😀',evidenceStatus:'not-established',items,subscales:[{id:'ordered-mean',name:'Authored ordered mean',itemIds:items.map(i=>i.id).reverse(),aggregation:'mean',missing:{mode:'minimum-answered',minimumAnswered:1}},{id:'single-sum',name:'Authored complete sum',itemIds:[items[0].id],aggregation:'sum',missing:{mode:'complete-case'}}]};
}
export function assessmentFileDraft({unfinished=false}={}){
 const draft=structuredClone(assessmentDraftFromDefinition(assessmentFileDefinition()));
 draft.itemBuffers[0].buffer.fields['item-target']='Retained inactive semantic target';
 draft.itemBuffers[0].buffer.fields['item-high']='Retained inactive explanation';
 draft.activeItem=structuredClone(draft.itemBuffers[0]);
 draft.ruleBuffers[1].buffer.fields['rule-k']='099';
 if(unfinished){draft.title='';draft.version='2.';draft.activeItem={id:'draft-item-500',buffer:structuredClone(draft.activeItem.buffer)};draft.activeItem.buffer.fields['item-prompt']='Unfinished raw\u0000 text\r\n日本語 😀';draft.activeItem.buffer.fields['anchor-min']='-';draft.activeItem.buffer.anchors[0].code='1.';draft.activeRule={id:'draft-rule-501',buffer:{fields:{'rule-name':'Unfinished rule','rule-method':'mean','rule-missing':'minimum-answered','rule-k':'-'},members:draft.items.map(i=>({id:i.id,checked:false})).reverse()}};}
 return draft;
}
