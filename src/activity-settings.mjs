// Only the numeric duration in the original, unambiguous cue is replaced.
// Activity text, pictures, formatting outside that number and IDs are untouched.
const clean=s=>s.replace(/[\r\n\u200b]/g,'').trim();
export function createActivitySettingsAdapter({run,sync=ctx=>ctx.sync()}){
  async function inspect(ctx,project,id){
    const item=project.modules.flatMap(m=>m.activities).find(a=>a.id===id);
    if(!item)throw new Error('The selected activity no longer exists.');
    const controls=ctx.document.contentControls.getByTag(`osg:activity:${id}`);controls.load('items');await ctx.sync();
    if(controls.items.length!==1)throw new Error('The selected activity is missing or duplicated.');
    const paragraphs=controls.items[0].getRange('Whole').paragraphs;paragraphs.load('items/text');await ctx.sync();
    if(paragraphs.items.length!==1||clean(paragraphs.items[0].text)!==item.title)throw new Error('The activity title is not a standalone paragraph. Preserve this copy before editing settings.');
    const cue=paragraphs.items[0].getPreviousOrNullObject();cue.load('text,isNullObject');await ctx.sync();
    if(cue.isNullObject||clean(cue.text)!==`${item.icon} · ${item.minutes} min`)throw new Error('The activity timing cue was changed or moved in Word. Settings were not applied. Preserve this copy before repairing the cue.');
    const parent=cue.parentContentControlOrNullObject,table=cue.parentTableOrNullObject;
    parent.load('isNullObject');table.load('isNullObject');
    const children=cue.getRange('Content').contentControls;
    children.load('items');await ctx.sync();
    // Older guides put their icon in this same paragraph. Searching/replacing
    // only the numeric text leaves those original inline pictures untouched.
    if(!parent.isNullObject||!table.isNullObject||children.items.length)throw new Error('The timing cue contains managed content or is nested. Preserve this copy before editing settings.');
    return {cue,item};
  }
  return {
    check:(project,id)=>run(ctx=>inspect(ctx,project,id)),
    write:(base,next,id)=>run(async ctx=>{
      const {cue,item}=await inspect(ctx,base,id),updated=next.modules.flatMap(m=>m.activities).find(a=>a.id===id);
      if(item.minutes===updated.minutes)return;
      const matches=cue.search(String(item.minutes),{matchCase:true});matches.load('items');await ctx.sync();
      if(matches.items.length!==1)throw new Error('The timing number is ambiguous. Settings were not applied.');
      matches.items[0].insertText(String(updated.minutes),'Replace');await sync(ctx);
      // Never claim completion after a dropped or ambiguous Word batch.
      await inspect(ctx,next,id);
    })
  };
}
