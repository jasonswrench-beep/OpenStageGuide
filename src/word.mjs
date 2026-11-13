import { ICONS,managedTags } from './model.mjs';
import {createDocumentStore} from './document-store.mjs';
import {createAcceptanceFaults,ACCEPTANCE_TITLE} from './acceptance-faults.mjs';
import {readCompressedDocument} from './office-file.mjs';
import {createActivitySettingsAdapter} from './activity-settings.mjs';
import {moduleAsset} from './diagnostics.mjs';
import {hasFrontMatter,frontMatterFields} from './frontmatter.mjs';
import {assertBlankMaster,inspectCapturedFrontMatter} from './checkpoint.mjs';
import {assertBlankBodyOoxml} from './frontmatter-package.mjs';
const acceptance=createAcceptanceFaults(()=>Office.context.document.settings);
export function supported() { return Office.context.requirements.isSetSupported('WordApi','1.3'); }
async function scan() {
  return Word.run(async ctx=>{
    const controls=ctx.document.contentControls;controls.load('items/tag,items/text');await ctx.sync();
    return controls.items.map(c=>({tag:c.tag,text:c.text}));
  });
}
// Tag presence does not prove placement. Native range comparisons reject
// reordered, nested, overlapping or unrelated managed fields. Heading anchors
// must be standalone paragraphs, outside tables and enclosing controls.
async function verifyPlacement(project){
  await Word.run(async ctx=>{
    const all=ctx.document.contentControls;all.load('items/tag,items/text');await ctx.sync();
    const controls=[...managedTags(project)].map(tag=>{
      const found=all.items.filter(c=>c.tag===tag);
      if(found.length!==1)throw new Error('A module placement field is missing or duplicated. Preserve this document and use recovery.');
      return found[0];
    });
    const ranges=controls.map(c=>c.getRange('Whole'));
    const checks=ranges.slice(1).map((r,i)=>ranges[i].compareLocationWith(r));
    const body=ctx.document.body.getRange('Whole'),bodyChecks=ranges.map(r=>body.compareLocationWith(r));
    const parents=controls.map(c=>{
      const parent=c.parentContentControlOrNullObject;parent.load('isNullObject');
      const table=c.getRange('Whole').parentTableOrNullObject;table.load('isNullObject');
      return {parent,table};
    });
    const headings=controls.filter(c=>c.tag.startsWith('osg:module:')||hasFrontMatter(project)&&(c.tag.startsWith('osg:section-heading:')||c.tag.startsWith('osg:course:'))).map(c=>{
      const paragraphs=c.getRange('Whole').paragraphs;paragraphs.load('items/text');
      return {c,paragraphs};
    });
    await ctx.sync();
    if(checks.some(r=>!['Before','AdjacentBefore'].includes(r.value)))throw new Error('Managed fields are outside their expected module order. Preserve this document; placement and export are paused.');
    if(bodyChecks.some(r=>!['Contains','ContainsStart','ContainsEnd','Equal'].includes(r.value))||parents.some(({parent,table})=>!parent.isNullObject||!table.isNullObject))throw new Error('Managed fields are nested or outside the main document paragraphs. Preserve this document before repairing its structure.');
    const clean=s=>s.replace(/[\r\n\u200b]/g,'').trim();
    if(headings.some(({c,paragraphs})=>paragraphs.items.length!==1||clean(paragraphs.items[0].text)!==clean(c.text)))throw new Error('A module heading is not a standalone insertion boundary. Preserve this document before repairing its structure.');
  });
}
async function field(ctx,text,type,id,style,append=text=>ctx.document.body.insertParagraph(text,'End')) {
  const para=append(text||'\u200b');para.styleBuiltIn=style;
  const control=para.getRange('Content').insertContentControl();
  control.tag=`osg:${type}:${id}`;control.title=`OpenStage Guide ${type}`;control.cannotDelete=true;
  // The invisible marker keeps optional empty controls from acquiring Word's instructional placeholder.
  await acceptance.sync(ctx,'write');return para;
}
async function iconData(label) {
  const file=ICONS.find(([x])=>x===label)?.[1]||'Interaction.png';
  const response=await fetch(moduleAsset('assets/icons/'+encodeURIComponent(file),import.meta.url));if(!response.ok)throw new Error('The activity icon could not be loaded.');
  const blob=await response.blob();
  return new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(reader.result.split(',')[1]);reader.onerror=reject;reader.readAsDataURL(blob);});
}
async function write(next,kind,item) {
  let picture;
  if(kind==='activity')picture=await iconData(item.icon);
    await Word.run(async ctx=>{
      let append=text=>ctx.document.body.insertParagraph(text,'End');
      if(kind==='activity'){
        const index=next.modules.findIndex(m=>m.activities.some(a=>a.id===item.id));
        if(index<0)throw new Error('The activity destination no longer exists.');
        const following=next.modules[index+1];
        if(following){
          const controls=ctx.document.contentControls.getByTag(`osg:module:${following.id}`);
          controls.load('items');await ctx.sync();
          if(controls.items.length!==1)throw new Error('The following module is missing or duplicated. Preserve this document and use recovery.');
          // Word may persist paragraph controls as block controls. Use the
          // control's explicit Before location so insertion stays outside it.
          // Keep one fixed boundary for every new paragraph.
          const boundary=controls.items[0];
          append=text=>boundary.insertParagraph(text,'Before');
        }
      }
      if(kind==='course'){
        if(hasFrontMatter(next)){const body=ctx.document.body.getOoxml();await ctx.sync();assertBlankBodyOoxml(body.value);}
        await field(ctx,next.title,'course',next.id,'Title');
        for(const f of frontMatterFields(next))await field(ctx,f.text,f.type,f.id,f.style);
      }
      if(kind==='module'){
        await field(ctx,item.title,'module',item.id,'Heading1');
        const label=ctx.document.body.insertParagraph('Learning objective','End');label.styleBuiltIn='Heading2';
        await field(ctx,item.objective,'objective',item.id,'Normal');
      }
      if(kind==='activity'){
        // Match the approximately one-inch teaching cues in the author's textbook.
        // Office image dimensions are points (72 points = one inch).
        const cue=append('');cue.styleBuiltIn='Normal';cue.spaceAfter=6;
        const image=cue.getRange('Start').insertInlinePictureFromBase64(picture,'Start');image.width=72;image.height=72;image.altTextDescription=item.icon+' icon';
        const heading=append(`${item.icon} · ${item.minutes} min`);heading.styleBuiltIn='Normal';
        await field(ctx,item.title,'activity',item.id,'Heading2',append);
        for(const [type,title] of [['facilitator','Facilitator instructions'],['participant','Participant instructions'],['answers','Facilitator answers']]){
          const label=append(title);label.styleBuiltIn='Normal';label.font.bold=true;
          await field(ctx,item[type],type,item.id,'Normal',append);
        }
      }
      await acceptance.sync(ctx,'write');
    });
}
const activitySettings=createActivitySettingsAdapter({run:fn=>Word.run(fn),sync:ctx=>acceptance.sync(ctx,'write')});
const zip=()=>{if(!globalThis.JSZip)throw new Error('The package verifier is unavailable.');return globalThis.JSZip;};
const verifyBlank=async journal=>assertBlankMaster(await readCompressedDocument(Office.context.document),zip(),journal);
const verifyFrontMatter=async (project,options)=>inspectCapturedFrontMatter(await readCompressedDocument(Office.context.document),project,zip(),options);
const store=createDocumentStore({settings:()=>acceptance.settings(),scan,write,verifyPlacement,checkActivitySettings:activitySettings.check,writeActivitySettings:activitySettings.write,verifyBlank,verifyFrontMatter});
export const {loadProject,readProject,persist,insert,recover,isBlocked,updateActivitySettings,updateCourseDetails,updateActivityMaterials,updateObjectiveCoverage,updateDeliveryPlan}=store;
export const captureMaster=()=>store.captureMaster(()=>readCompressedDocument(Office.context.document));
export const acceptanceEligible=()=>acceptance.eligible();
export async function armAcceptance(point){
  acceptance.clear();
  const p=await store.readProject(),rows=await scan();
  if(p?.title!==ACCEPTANCE_TITLE||!rows.some(r=>r.tag===`osg:course:${p.id}`&&r.text.trim()===ACCEPTANCE_TITLE))throw new Error('Disposable course marker does not match Word text.');
  acceptance.arm(point);
}
export const clearAcceptance=()=>acceptance.clear();
export async function acceptanceSnapshot(){
 return store.inspect(async()=>{
  if(!acceptance.eligible())throw new Error('Not a marked acceptance document.');
  const s=Office.context.document.settings;
  await new Promise((ok,no)=>s.refreshAsync(r=>r.status==='succeeded'?ok():no(new Error('Native settings refresh failed'))));
  return {project:s.get('openstage.guide.project.v1'),pending:s.get('openstage.guide.pending.v1'),fields:await scan(),blocked:store.isBlocked(),events:acceptance.events()};
 });
}
export async function jump(id,type='module') {
  if(!['module','activity'].includes(type))throw new Error('Unsupported outline destination.');
  await store.readProject();
  await Word.run(async ctx=>{const c=ctx.document.contentControls.getByTag(`osg:${type}:${id}`);c.load('items');await ctx.sync();if(c.items.length!==1)throw new Error('This outline item is missing or duplicated in Word. Refresh the course before continuing.');c.items[0].getRange().select();await ctx.sync();});
}
