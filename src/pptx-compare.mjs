import {validateSlideSource,serializeReview} from './pptx-review.mjs';
const content=s=>JSON.stringify([s.title,s.text,s.notes]);
function indexContent(slides){const index=new Map();for(const slide of slides){const key=content(slide);if(!index.has(key))index.set(key,[]);index.get(key).push(slide.id);}return index;}
// Preserve relative order of common IDs: inserts/removals alone are not moves.
function stableOrder(oldIds,newIds){
  const matrix=Array.from({length:oldIds.length+1},()=>new Uint16Array(newIds.length+1));
  for(let i=oldIds.length-1;i>=0;i--)for(let j=newIds.length-1;j>=0;j--)matrix[i][j]=oldIds[i]===newIds[j]?matrix[i+1][j+1]+1:Math.max(matrix[i+1][j],matrix[i][j+1]);
  const stable=new Set();let i=0,j=0;while(i<oldIds.length&&j<newIds.length){if(oldIds[i]===newIds[j]){stable.add(oldIds[i]);i++;j++;}else if(matrix[i+1][j]>=matrix[i][j+1])i++;else j++;}return stable;
}
export function comparePptx(original,incoming,review,editor){
  validateSlideSource(original);validateSlideSource(incoming);serializeReview(original,review,editor);
  const old=new Map(original.slides.map(s=>[s.id,s])),next=new Map(incoming.slides.map(s=>[s.id,s]));
  const oldContent=indexContent(original.slides),newContent=indexContent(incoming.slides);
  const local=new Map(review.slides.map(s=>[s.id,s]));local.set(editor.id,editor);
  const stable=stableOrder(original.slides.filter(s=>next.has(s.id)).map(s=>s.id),incoming.slides.filter(s=>old.has(s.id)).map(s=>s.id));
  const rows=[];
  for(const after of incoming.slides){
    const before=old.get(after.id),authored=before?local.get(before.id):null;
    const candidates=(oldContent.get(content(after))||[]).filter(id=>id!==after.id);
    const fields=before?['title','text','notes'].map(key=>({key,incomingChanged:after[key]!==before[key],localChanged:authored[key]!==before[key],conflict:after[key]!==before[key]&&authored[key]!==before[key]&&after[key]!==authored[key]})):[];
    const changed=fields.some(f=>f.incomingChanged),moved=!!before&&!stable.has(after.id);
    rows.push({key:`incoming:${after.id}`,kind:before?'match-candidate':'added-candidate',id:after.id,originalOrder:before?.order??null,incomingOrder:after.order,
      changed,moved,fields,conflict:fields.some(f=>f.conflict),contentCandidates:candidates,
      warnings:candidates.length?[before?'This ID may have been reused; identical source content exists under other original IDs.':'Content also appears under original IDs; identity needs manual review.']:[],
      original:before?structuredClone(before):null,authored:authored?structuredClone(authored):null,incoming:structuredClone(after)});
  }
  for(const before of original.slides.filter(s=>!next.has(s.id))){const candidates=newContent.get(content(before))||[];rows.push({key:`missing:${before.id}`,kind:'missing-candidate',id:before.id,originalOrder:before.order,incomingOrder:null,changed:false,moved:false,fields:[],conflict:false,contentCandidates:[...candidates],warnings:candidates.length?['Original content appears under incoming IDs; absence is not proof of deletion.']:[],original:structuredClone(before),authored:structuredClone(local.get(before.id)),incoming:null});}
  const count=predicate=>rows.filter(predicate).length;
  return {format:'openstage-pptx-comparison',version:1,originalHash:original.source.sha256,incomingHash:incoming.source.sha256,sameSnapshot:original.source.sha256===incoming.source.sha256,
    warning:'Comparison proposals only. Numeric slide IDs are presentation-local candidates, not proof of deck lineage. Images/layout/media changes are not detected. No baseline or authored text is updated.',
    counts:{added:count(r=>r.kind==='added-candidate'),missing:count(r=>r.kind==='missing-candidate'),changed:count(r=>r.changed),moved:count(r=>r.moved),unchanged:count(r=>r.kind==='match-candidate'&&!r.changed&&!r.moved),conflicts:count(r=>r.conflict),ambiguous:count(r=>r.contentCandidates.length>0)},rows};
}
