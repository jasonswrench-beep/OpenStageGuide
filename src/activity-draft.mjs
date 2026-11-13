export const ACTIVITY_DRAFT_FIELDS=Object.freeze(['title','icon','minutes','facilitator','participant','answers','lines']);
const FLAGS=['badInput','customError','patternMismatch','rangeOverflow','rangeUnderflow','stepMismatch','tooLong','tooShort','typeMismatch','valueMissing'];
// Raw values and validity, never library/model normalization. A number input
// can have value==='' while retaining an unfinished native badInput buffer.
export function rawControl(input){return {value:input.value,invalid:Object.fromEntries(FLAGS.map(key=>[key,!!input.validity?.[key]]))};}
export const sameRaw=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
export const rawInvalid=control=>Object.values(control.invalid).some(Boolean);
export function wouldReplaceRaw(current,defaults,candidate,currentOrigin,targetOrigin){
 const invalid=Object.entries(current.fields).some(([key,c])=>Object.entries(c.invalid).some(([flag,on])=>on&&!defaults.fields[key].invalid[flag]))||(current.materials||[]).some(row=>Object.values(row).some(rawInvalid));
 const values=x=>({fields:Object.fromEntries(ACTIVITY_DRAFT_FIELDS.map(k=>[k,x.fields[k].value])),materials:x.materials===null?null:x.materials.map(row=>Object.fromEntries(Object.entries(row).map(([k,v])=>[k,v.value])))});
 const active=!!currentOrigin||invalid||!sameRaw(values(current),values(defaults));
 return active&&(invalid||!sameRaw(values(current),values(candidate))||!!currentOrigin&&!sameRaw(currentOrigin,targetOrigin));
}
