// How the guide pane talks to the document. Two adapters share one interface:
//   createWordAdapter  : the real thing, through Office.js
//   createDemoAdapter  : an in-memory document, for the browser preview and for tests
// Interface: paragraphs(), title(), setTitle(t), customCues(), setCustomCues(names), cursorIndex(), insert(build,{where,after,select,spare,remove}),
//            go(i,text), plainLineAfter(i) (a plain line below a block, cursor in it), moveRange({start,end,guard,after}), deleteRange({start,end,guard}), applyEdits(edits), reviewParts(),
//            refreshFields(), keepPaneOpen(), onSelectionChanged(fn),
//            sourcePackage() (the whole document as a flat Office package), openDocument(base64,name) (a new Word window),
//            slideMemory(), setSlideMemory(text) (what the guide remembers of the presentations its slides came from),
//            versionMemory(), setVersionMemory(text) (the delivery versions the author planned),
//            offerDownload(base64,name), pickFiles({kind}) (the file window: Word's pane cannot open a file chooser).
//            `build(ctx)` returns the body XML to insert.
import {makeContext,buildPackage,timeText,esc,SPARE_PARAGRAPH} from './guide-ooxml.mjs';
import {REF_TAG_PREFIX} from './guide-forms.mjs';
import {NS,STYLE_IDS,STYLE_NAMES,stylesPart,numberingPart,BULLET_NUM_ID} from './guide-styles.mjs';
import {iconUrl,ICON_FILES} from './guide-cues.mjs';
import {cleanStyleEdits,applyStyleEdits,listStyles} from './guide-styles-edit.mjs';
import {cleanParagraphXml} from './guide-xml-clean.mjs';
import {readStyles,textOf} from './guide-check.mjs';
import {explainWordError} from './word-errors.mjs';
import {pickFilesWithDialog,pickWithBrowserInputs} from './file-pick.mjs';
import {readPptx} from './pptx.mjs';
import {readCompressedDocument} from './office-file.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG='http://schemas.microsoft.com/office/2006/xmlPackage';
const AUTO_SHOW='Office.AutoShowTaskpaneWithDocument';
// The author's own cue names are one custom document property (a JSON list of up to three names).
export const CUSTOM_CUES_KEY='osg.customCues';
// The guide's look is one custom document property too (see guide-looks.mjs).
export const LOOK_PROPERTY='osg.look';
// What the guide remembers of the presentations its slides came from is one document setting (see slides-sync.mjs).
const SLIDE_MEMORY_KEY='osg.slides.v1';
const VERSIONS_KEY='osg.versions.v1';
// Word only reopens a pane with its document when the manifest has a ShowTaskpane action named Office.AutoShowTaskpaneWithDocument. Version
// 1.1.0.0 of manifest.xml gave the Open button that action, and in the author's Word (2026-11-01) the button then did nothing at all, so the
// Open button runs the openDesigner function again (version 1.1.0.1) and nothing in the manifest has the action. While it is false the pane does
// not ask to reopen (a document that asks makes Word say "We can't find the task pane to open"), and it takes the request out of a document that
// carries it. Reopening with the pane is therefore not available; it stays false until a manifest with the action is proven in Word.
const AUTO_OPEN_SUPPORTED=false;

// Text compared by its letters and digits only, so a tab, a space or the placeholder character a picture can leave
// behind never makes two paragraphs look different.
const norm=s=>String(s??'').replace(/[^\p{L}\p{N}]+/gu,'').toLowerCase();
const OSG_STYLE_NAMES=new Map(Object.entries(STYLE_IDS).filter(([k])=>STYLE_NAMES[k]).map(([k,id])=>[id,STYLE_NAMES[k]]));

// What the inserted paragraphs should read as: their words, and their style when it is one of the guide's own.
// (Built-in styles such as Heading 1 are named in the user's language, so they are not checked.)
export function expectedParagraphs(body,DOMParserClass=globalThis.DOMParser){
  const doc=new DOMParserClass().parseFromString(`<w:document ${NS}><w:body>${body}</w:body></w:document>`,'application/xml');
  // Every paragraph in document order, including those inside tables, which Word also counts as paragraphs.
  return Array.from(doc.getElementsByTagNameNS(W,'body')[0].getElementsByTagNameNS(W,'p')).map(p=>{
    const ppr=Array.from(p.childNodes).find(n=>n.nodeType===1&&n.localName==='pPr');
    const id=ppr&&Array.from(ppr.childNodes).find(n=>n.nodeType===1&&n.localName==='pStyle')?.getAttributeNS(W,'val');
    return {style:OSG_STYLE_NAMES.get(id)||null,text:textOf(p)};
  });
}

const CHANGED='That part of the guide has changed since you chose it. Put your cursor in it again, then try again.';
// Word's insertOoxml drops the formatting of the last paragraph of a package and can join it to the next one (see SPARE_PARAGRAPH).
// A package Word made for a range ends with the document's section properties, so the spare empty paragraph goes just before them.
export function withSpareParagraph(pkg){
  const body=pkg.lastIndexOf('</w:body>');
  if(body<0)return pkg;
  const sect=pkg.lastIndexOf('<w:sectPr',body),tail=sect>=0?pkg.slice(sect,body):'';
  // Section properties that close a paragraph (inside its pPr) are not the body's own, so the spare goes at the very end then.
  const at=sect>=0&&tail.includes('</w:sectPr>')&&!/<\/w:p>|<\/w:pPr>/.test(tail)?sect:body;
  return pkg.slice(0,at)+SPARE_PARAGRAPH+pkg.slice(at);
}

const toBase64=buffer=>{
  const bytes=new Uint8Array(buffer);let s='';
  for(let i=0;i<bytes.length;i+=0x8000)s+=String.fromCharCode.apply(null,bytes.subarray(i,i+0x8000));
  return btoa(s);
};
const iconCache=new Map();
export async function loadIcons(names,fetchFn=globalThis.fetch){
  const out={};
  for(const name of names){
    if(!iconCache.has(name)){
      let response;
      try{response=await fetchFn(iconUrl(name));}catch{response=null;}
      if(!response||!response.ok)throw new Error(`The picture for the ${name} cue could not be loaded. Check that the OpenStage Guide server is running, then try again.`);
      iconCache.set(name,toBase64(await response.arrayBuffer()));
    }
    out[name]=iconCache.get(name);
  }
  return out;
}

export function createWordAdapter({Word:WordApi,Office:OfficeApi,fetchFn,DOMParserClass=globalThis.DOMParser}={}){
  // Office.js adds Word to the page after the pane has loaded, so look both up when they are needed.
  const officeApi=()=>OfficeApi||globalThis.Office;
  const run=fn=>{
    const api=WordApi||globalThis.Word;
    if(!api?.run)return Promise.reject(new Error('Word is not ready yet. Close the OpenStage Guide pane, open it again from the ribbon, then try again.'));
    // Word's own error codes ("InvalidArgument") become a sentence about what to do; the pane's own messages pass through unchanged.
    return Promise.resolve(api.run(fn)).catch(e=>{throw explainWordError(e);});
  };
  // The active look's theme: styles an insert has to define are written from it.
  let theme;
  const readProperty=key=>run(async c=>{
    const item=c.document.properties.customProperties.getItemOrNullObject(key);item.load('value');await c.sync();
    return item.isNullObject?null:String(item.value);
  });
  const writeProperty=(key,value)=>run(async c=>{c.document.properties.customProperties.add(key,value);await c.sync();});
  return {
    kind:'word',
    setTheme(t){theme=t||undefined;},
    async paragraphs(){
      return run(async c=>{
        const ps=c.document.body.paragraphs;ps.load('items/text,items/style,items/styleBuiltIn');await c.sync();
        return ps.items.map(p=>({text:p.text,style:p.style,builtIn:p.styleBuiltIn}));
      });
    },
    async title(){
      return run(async c=>{const pr=c.document.properties;pr.load('title');await c.sync();return pr.title||'';});
    },
    async setTitle(title){
      return run(async c=>{c.document.properties.title=title;await c.sync();});
    },
    // The author's own cue names, kept as one custom document property so they travel with the file.
    async customCues(){
      const raw=await readProperty(CUSTOM_CUES_KEY);
      if(raw===null)return [];
      try{const v=JSON.parse(raw);return Array.isArray(v)?v.map(String):[];}catch{return [];}
    },
    async setCustomCues(names){return writeProperty(CUSTOM_CUES_KEY,JSON.stringify(names));},
    // The look the guide was given (a JSON text, see guide-looks.mjs), or null when it has none.
    async getLook(){return readProperty(LOOK_PROPERTY);},
    async setLook(text){return writeProperty(LOOK_PROPERTY,text);},
    // What the guide remembers of the presentations its slides came from (see slides-sync.mjs): kept in the document's settings, which
    // travel with the file. Not being able to read or write it never stops anything; it only means slides are matched by title.
    async slideMemory(){
      try{const v=officeApi()?.context?.document?.settings?.get(SLIDE_MEMORY_KEY);return typeof v==='string'?v:null;}catch{return null;}
    },
    async setSlideMemory(text){
      const settings=officeApi()?.context?.document?.settings;
      if(!settings)return false;
      settings.set(SLIDE_MEMORY_KEY,text);
      await new Promise((resolve,reject)=>settings.saveAsync(r=>r.status==='succeeded'?resolve():reject(new Error('The document did not save what the guide remembers about its slides.'))));
      return true;
    },
    // The delivery versions the author planned (see guide-versions.mjs): a JSON text kept in the document's settings, which travel with the file.
    async versionMemory(){
      try{const v=officeApi()?.context?.document?.settings?.get(VERSIONS_KEY);return typeof v==='string'?v:null;}catch{return null;}
    },
    async setVersionMemory(text){
      const settings=officeApi()?.context?.document?.settings;
      if(!settings)throw new Error('This version of Word cannot keep your delivery versions in the document.');
      settings.set(VERSIONS_KEY,text);
      await new Promise((resolve,reject)=>settings.saveAsync(r=>r.status==='succeeded'?resolve():reject(new Error('The document did not keep your delivery versions.'))));
      return true;
    },
    // Choosing files: a file chooser cannot be opened from inside the pane, so a small window of its own does it (see file-pick.mjs).
    // kind is 'picture', 'slides' or 'files' (any file a file box in the pane asked for: accept, multiple, label);
    // resolves {deck, files} or {cancelled:true}.
    async pickFiles({kind,accept,multiple,label}){
      return pickFilesWithDialog({Office:officeApi(),url:new URL('picker.html',globalThis.location.href).href,kind,accept,multiple,label});
    },
    // Changes the styles the document already has. Each group of settings goes in its own step, so one that Word refuses does
    // not stop the rest. A style the document does not have yet is skipped: the next insert makes it from the look. Returns
    // {applied:[names], skipped:[names], refused:[{style, what, why}]}.
    async applyLook(specs){
      return run(async c=>{
        const styles=c.document.getStyles(),found=specs.map(spec=>({spec,style:styles.getByNameOrNullObject(spec.name)}));
        found.forEach(f=>f.style.load('isNullObject'));await c.sync();
        const applied=[],skipped=[],refused=[];
        for(const {spec,style} of found){
          if(style.isNullObject){skipped.push(spec.name);continue;}
          const steps=[];
          if(Object.keys(spec.font).length)steps.push(['font',()=>{for(const [k,v] of Object.entries(spec.font))style.font[k]=v;}]);
          if(Object.keys(spec.paragraph).length)steps.push(['spacing and indents',()=>{for(const [k,v] of Object.entries(spec.paragraph))style.paragraphFormat[k]=v;}]);
          if(spec.shading)steps.push(['shading',()=>{style.shading.backgroundPatternColor=spec.shading;}]);
          for(const b of spec.borders)steps.push([`${b.location.toLowerCase()} border`,()=>{const border=style.borders.getByLocation(b.location);border.type='Single';border.color=b.color;border.width=b.width;}]);
          let any=false;
          for(const [what,change] of steps){
            try{change();await c.sync();any=true;}
            catch(e){refused.push({style:spec.name,what,why:String(e?.code||e?.message||'refused')});}
          }
          if(any||!steps.length)applied.push(spec.name);
        }
        return {applied,skipped,refused};
      });
    },
    // Puts the running header and footer on the document's first section (a package, or null to take one away; undefined leaves it).
    async setHeaderFooter({header,footer}){
      return run(async c=>{
        const section=c.document.sections.getFirst(),done=[],failed=[];
        for(const [what,content,get] of [['header',header,'getHeader'],['footer',footer,'getFooter']]){
          if(content===undefined)continue;
          try{
            const body=section[get]('Primary');
            if(content===null)body.clear();else body.insertOoxml(content,'Replace');
            await c.sync();done.push(what);
          }catch(e){failed.push({what,why:String(e?.code||e?.message||'refused')});}
        }
        return {done,failed};
      });
    },
    // Index, in body.paragraphs order, of the paragraph that holds the start of the selection. Every paragraph has its
    // own start point, so matching start points picks exactly one even when the cursor sits on a paragraph boundary.
    async cursorIndex(){
      return run(async c=>{
        const here=c.document.getSelection().paragraphs.getFirst().getRange('Start'),ps=c.document.body.paragraphs;
        ps.load('items');await c.sync();
        const where=ps.items.map(p=>p.getRange('Start').compareLocationWith(here));
        await c.sync();
        const at=where.findIndex(w=>w.value==='Equal');
        return at<0?ps.items.length-1:at;
      });
    },
    // Inserts after paragraph number `after`, or at the top (where:'start'). `remove` {start,end,guard} also deletes
    // the paragraphs of the block being replaced, after checking that the first one still reads as `guard`.
    // Word does not always do what insertOoxml asks, so this reads the result back, puts the guide's styles on any
    // paragraph that lost its own, removes empty paragraphs the insert left behind, and stops with a plain message
    // if the words did not arrive as separate paragraphs.
    async insert(build,{where='after',after,select='last',spare=true,remove=null}={}){
      const ctx=makeContext(),body=build(ctx),icons=await loadIcons([...ctx.used],fetchFn);
      const pkg=buildPackage(body,ctx,icons,{spare,theme}),want=expectedParagraphs(body,DOMParserClass),N=want.length;
      await run(async c=>{
        const all=c.document.body.paragraphs;all.load('items/text');await c.sync();
        const before=all.items.length;
        if(remove&&norm(all.items[remove.start]?.text)!==norm(remove.guard))throw new Error('That block has changed since you chose it. Select it again, then choose Edit this block.');
        let start0=0;
        if(where==='start')c.document.body.insertOoxml(pkg,'Start');
        else{
          const anchor=all.items[after];
          if(!anchor)throw new Error('That part of the document is no longer there. Choose Refresh and try again.');
          const table=anchor.parentTableOrNullObject;table.load('isNullObject');await c.sync();
          start0=after+1;
          if(!table.isNullObject){
            // From inside a table the block goes after the whole table, so the paragraphs of the rest of the table come first.
            const rest=all.items.slice(after+1,after+1+400).map(p=>p.parentTableOrNullObject);
            rest.forEach(t=>t.load('isNullObject'));await c.sync();
            const out=rest.findIndex(t=>t.isNullObject);
            start0=after+1+(out<0?rest.length:out);
          }
          (table.isNullObject?anchor.getRange('Whole'):table.getRange('Whole')).insertOoxml(pkg,'After');
        }
        await c.sync();
        const now=c.document.body.paragraphs;now.load('items/text,items/style');await c.sync();
        const items=now.items,text=i=>items[i]?.text;
        let first=-1;
        for(let i=Math.max(0,start0-1);i<=Math.min(items.length-1,start0+3)&&first<0;i++)if(norm(text(i))===norm(want[0].text))first=i;
        const diagnostic=`[Diagnostic: inserted ${N} paragraphs; document had ${before} and now has ${items.length}; first found at ${first}; expected near ${start0}]`;
        if(first<0)throw new Error(`Word did not insert this where expected, so the guide may look wrong. Click in the document and press Ctrl+Z, then try again. ${diagnostic}`);
        // Each of the guide's own paragraphs must have arrived on its own, in order. They are found one after another
        // rather than at fixed positions, because Word may count a table's rows as paragraphs and the package may hold tables.
        const fix=[],found=[];
        let pos=first;
        for(let j=0;j<N;j++){
          if(!want[j].style)continue;
          let at=-1;
          for(let i=pos;i<items.length&&i<=pos+N+12&&at<0;i++)if(norm(text(i))===norm(want[j].text))at=i;
          if(at<0)throw new Error(`Word joined part of this block to another paragraph, so the guide may look wrong. Click in the document and press Ctrl+Z, then try again. ${diagnostic}`);
          found.push({j,at});pos=at+1;
          if(items[at].style!==want[j].style)fix.push({at,style:want[j].style});
        }
        const lastOwn=found.length?found.at(-1).at:first;
        const kill=[];
        // Anything empty between the insertion point and the first paragraph is Word's, not the author's.
        for(let i=start0;i<first;i++)if(!norm(text(i)))kill.push(i);
        // After the block, the paragraph count says how many empty paragraphs Word left. With tables in the content the count
        // is less certain, so then only the one spare paragraph is cleared.
        const extra=items.length-before-N-kill.length;
        const allowed=/<w:tbl>/.test(body)?Math.min(extra,1):extra;
        let trailing=0;
        if(extra>0)
          // Word cannot delete the document's final paragraph, so an empty one there is left alone.
          for(let i=lastOwn+1;i<items.length-1&&trailing<allowed&&!norm(text(i));i++){kill.push(i);trailing++;}
        // A block being replaced goes whole. A table in it must be deleted as a table: deleting a cell's paragraph only
        // empties the cell. The paragraphs of one table are a single run, so the table is deleted once per run.
        const gone=[];
        if(remove){
          const doomed=[];for(let i=remove.start;i<=remove.end;i++)doomed.push(i);
          const tables=doomed.map(i=>items[i].parentTableOrNullObject);
          tables.forEach(t=>t.load('isNullObject'));
          await c.sync();
          let inTable=false;
          doomed.forEach((i,n)=>{
            if(tables[n].isNullObject){inTable=false;kill.push(i);}
            else{if(!inTable)tables[n].delete();inTable=true;gone.push(i);}
          });
        }
        for(const f of fix)items[f.at].style=f.style;
        kill.sort((a,b)=>b-a).forEach(i=>items[i].delete());
        try{await c.sync();}
        catch(e){throw new Error(`Word could not finish tidying this block${fix.length?' or apply the guide\'s styles to it':''}, so it may look wrong. Click in the document and press Ctrl+Z, then try again. ${diagnostic}`);}
        const at=(select==='first'?first:lastOwn);
        const target=at-kill.filter(i=>i<at).length-gone.filter(i=>i<at).length;
        const fresh=c.document.body.paragraphs;fresh.load('items/text,items/style');await c.sync();
        // When the insert ends the document, the last paragraph is the block's own (shaded notes, a cue line). Word then has nothing plain
        // below it: pressing Enter only makes another notes line, and the author cannot get out (seen in the author's recording, 2026-11-08).
        // So the document always ends in a plain, empty paragraph the cursor can go to.
        const end=fresh.items.at(-1);
        if(end){
          if(/^OSG /i.test(end.style||'')){
            if(norm(end.text))c.document.body.insertParagraph('','End').style='Normal';
            else end.style='Normal';
          }
          await c.sync();
        }
        fresh.items[target]?.select('End');
        await c.sync();
      });
    },
    // A way out of a block. In the shaded notes Enter only starts another line of notes, so the author cannot get below a block that
    // has nothing plain under it (seen in the author's recording, 2026-11-08: slides put in before a Transition). This makes sure there is
    // a plain, empty paragraph directly after paragraph `index` (the block's last) and puts the cursor in it. One that is already there
    // is used rather than another made. Returns {at, created}.
    async plainLineAfter(index){
      return run(async c=>{
        const all=c.document.body.paragraphs;all.load('items/text,items/style');await c.sync();
        const here=all.items[index];
        if(!here)throw new Error('That part of the document is no longer there. Choose Refresh and try again.');
        const next=all.items[index+1];
        if(next&&!norm(next.text)&&next.style==='Normal'){next.select('Start');await c.sync();return {at:index+1,created:false};}
        const table=here.parentTableOrNullObject;table.load('isNullObject');await c.sync();
        const line=(table.isNullObject?here:table).insertParagraph('','After');
        line.style='Normal';
        line.select('Start');
        await c.sync();
        return {at:index+1,created:true};
      });
    },
    // Moves the cursor to paragraph `index`. If the paragraph there no longer reads as `expect`, the document has
    // changed since the list was made, so it looks for the same words nearby rather than going somewhere wrong.
    async go(index,expect){
      return run(async c=>{
        const ps=c.document.body.paragraphs;ps.load('items/text');await c.sync();
        let at=index;
        if(expect!==undefined&&norm(expect)){
          if(!ps.items[at]||norm(ps.items[at].text)!==norm(expect)){
            const hits=ps.items.map((p,i)=>norm(p.text)===norm(expect)?i:-1).filter(i=>i>=0).sort((a,b)=>Math.abs(a-index)-Math.abs(b-index));
            if(!hits.length)throw new Error('That part of the document has changed since this list was made. Choose Review again.');
            at=hits[0];
          }
        }
        if(!ps.items[at])throw new Error('That part of the document is no longer there. Choose Refresh and try again.');
        ps.items[at].select('Start');await c.sync();
      });
    },
    // Deletes paragraphs start..end (a whole lesson or module), after checking that the first still reads as `guard`. A table in the
    // block is deleted as a table (deleting a cell's paragraph only empties the cell). Word will not remove the document's last
    // paragraph, so when the block runs to the end that paragraph is emptied and given the Normal style instead.
    async deleteRange({start,end,guard}){
      return run(async c=>{
        const all=c.document.body.paragraphs;all.load('items/text');await c.sync();
        const items=all.items;
        if(!items[start]||!items[end]||end<start||norm(items[start].text)!==norm(guard))throw new Error(CHANGED);
        const doomed=[];for(let i=start;i<=end;i++)doomed.push(i);
        const tables=doomed.map(i=>items[i].parentTableOrNullObject);
        tables.forEach(t=>t.load('isNullObject'));await c.sync();
        let inTable=false;const kill=[];
        doomed.forEach((i,n)=>{
          if(tables[n].isNullObject){inTable=false;kill.push(i);}
          else{if(!inTable)tables[n].delete();inTable=true;}
        });
        kill.sort((a,b)=>b-a).forEach(i=>{
          if(i===items.length-1){items[i].insertText('','Replace');items[i].style='Normal';}
          else items[i].delete();
        });
        try{await c.sync();}catch(e){throw new Error('Word could not delete this part of the guide. Nothing may have changed; if it looks wrong, click in the document and press Ctrl+Z.');}
        const fresh=c.document.body.paragraphs;fresh.load('items');await c.sync();
        fresh.items[Math.max(0,Math.min(start,fresh.items.length-1))]?.select('Start');
        await c.sync();
      });
    },
    // Moves paragraphs start..end (a whole lesson or module) to after paragraph `after` (-1 is the top of the document). Word has no
    // cut and paste for this, so the block's own OOXML is copied to the new place, read back to see that every paragraph arrived with
    // its words and style, and only then is the original deleted. If any check fails the original is left alone.
    async moveRange({start,end,guard,after}){
      const n=end-start+1;
      const src=await run(async c=>{
        const all=c.document.body.paragraphs;all.load('items/text,items/style');await c.sync();
        const items=all.items;
        if(!items[start]||!items[end]||end<start||norm(items[start].text)!==norm(guard))throw new Error(CHANGED);
        if(after>=start&&after<=end)throw new Error('A part of the guide cannot be moved into itself.');
        if(after>=items.length)throw new Error('That part of the document is no longer there. Choose Refresh and try again.');
        const first=items[start].parentTableOrNullObject,last=items[end].parentTableOrNullObject;
        first.load('isNullObject');last.load('isNullObject');await c.sync();
        if(!first.isNullObject||!last.isNullObject)throw new Error('This part of the guide starts or ends inside a table, so it cannot be moved from here. Move it in Word instead.');
        const o=items[start].getRange('Whole').expandTo(items[end].getRange('Whole')).getOoxml();await c.sync();
        return {pkg:o.value,want:items.slice(start,end+1).map(p=>({text:p.text,style:p.style}))};
      });
      const pkg=withSpareParagraph(src.pkg);
      const copy=await run(async c=>{
        const all=c.document.body.paragraphs;all.load('items/text');await c.sync();
        const before=all.items.length;
        let start0=0;
        if(after<0)c.document.body.insertOoxml(pkg,'Start');
        else{
          const anchor=all.items[after];
          const table=anchor.parentTableOrNullObject;table.load('isNullObject');await c.sync();
          start0=after+1;
          if(!table.isNullObject){
            const rest=all.items.slice(after+1,after+1+400).map(p=>p.parentTableOrNullObject);
            rest.forEach(t=>t.load('isNullObject'));await c.sync();
            const out=rest.findIndex(t=>t.isNullObject);
            start0=after+1+(out<0?rest.length:out);
          }
          (table.isNullObject?anchor.getRange('Whole'):table.getRange('Whole')).insertOoxml(pkg,'After');
        }
        await c.sync();
        const now=c.document.body.paragraphs;now.load('items/text,items/style');await c.sync();
        const items=now.items;
        const diagnostic=`[Diagnostic: moved ${n} paragraphs; document had ${before} and now has ${items.length}; expected the copy near ${start0}]`;
        let first=-1;
        for(let i=Math.max(0,start0-1);i<=Math.min(items.length-1,start0+3)&&first<0;i++)if(norm(items[i].text)===norm(src.want[0].text))first=i;
        const arrived=first>=0&&src.want.every((w,j)=>items[first+j]&&norm(items[first+j].text)===norm(w.text));
        if(!arrived)throw new Error(`Word did not bring every paragraph across, so nothing was moved and the original is still where it was. If an extra copy appeared, click in the document and press Ctrl+Z. ${diagnostic}`);
        const fix=[];
        src.want.forEach((w,j)=>{if(w.style&&items[first+j].style!==w.style)fix.push({at:first+j,style:w.style});});
        // Anything empty before the copy, and the empty paragraphs after it that the spare paragraph and Word left, are Word's, not the author's.
        const kill=[];
        for(let i=start0;i<first;i++)if(!norm(items[i].text))kill.push(i);
        const extra=items.length-before-n-kill.length;
        let trailing=0;
        for(let i=first+n;i<items.length-1&&trailing<extra&&!norm(items[i].text);i++){kill.push(i);trailing++;}
        for(const f of fix)items[f.at].style=f.style;
        kill.sort((a,b)=>b-a).forEach(i=>items[i].delete());
        try{await c.sync();}
        catch(e){throw new Error(`Word could not finish tidying the moved part, so the guide may look wrong. Nothing was removed from the old place. Click in the document and press Ctrl+Z, then try again. ${diagnostic}`);}
        return {first:first-kill.filter(i=>i<first).length};
      });
      // The copy is in place. When it went above the original, the original has moved down by the length of the block.
      const oldStart=after<start?start+n:start;
      try{await this.deleteRange({start:oldStart,end:oldStart+n-1,guard});}
      catch(e){throw new Error(`The part was copied to its new place, but the old one could not be removed, so it now appears twice. Click in the document and press Ctrl+Z to undo the move. ${(e&&e.message)||''}`);}
      const at=after<start?copy.first:copy.first-n;
      await this.go(at,guard);
      return {at};
    },
    // Ask Word to open this pane again when the document is next opened. Word only does that when the add-in's manifest
    // has a ShowTaskpane action whose TaskpaneId is Office.AutoShowTaskpaneWithDocument. This add-in's buttons run
    // functions instead, so asking made Word say "We can't find the task pane to open" on every reopen (seen 2026-10-31),
    // and giving the Open button that action stopped it opening (seen 2026-11-01). So this does not ask by default, and it
    // removes the request from a document that carries it. A caller can still pass manifestSupportsIt:true.
    async keepPaneOpen({manifestSupportsIt=AUTO_OPEN_SUPPORTED}={}){
      try{
        const settings=officeApi()?.context?.document?.settings;
        if(!settings)return false;
        const asked=settings.get(AUTO_SHOW)===true;
        if(manifestSupportsIt===asked)return false;
        if(manifestSupportsIt)settings.set(AUTO_SHOW,true);else settings.remove(AUTO_SHOW);
        await new Promise((resolve,reject)=>settings.saveAsync(r=>r.status==='succeeded'?resolve():reject(new Error('not saved'))));
        return manifestSupportsIt;
      }catch{return false;}
    },
    // Tell the pane when the cursor moves or the document changes under it (typing, Undo, Redo). Word can refuse to
    // hook the event while a document is still opening, and says so only through the callback, so a refusal is retried.
    onSelectionChanged(handler,{retries=5,delayMs=1000}={}){
      try{
        const office=officeApi(),document=office?.context?.document;
        if(!document?.addHandlerAsync)return false;
        const attach=left=>document.addHandlerAsync(office.EventType.DocumentSelectionChanged,()=>handler(),result=>{
          if(result?.status==='failed'&&left>0)setTimeout(()=>attach(left-1),delayMs);
        });
        attach(retries);
        return true;
      }catch{return false;}
    },
    // Heading edits replace the paragraph's text. Time edits keep the icon and replace the words after it.
    async applyEdits(edits){
      return run(async c=>{
        const ps=c.document.body.paragraphs;ps.load('items/text');await c.sync();
        // Each edit is sent to Word on its own, so that when Word refuses one the others are still made, and the message can say exactly
        // what to type by hand to finish it.
        const refused=[];
        const quote=t=>{const x=String(t??'').replace(/[\r\n\u0007\ufffc]+/g,' ').replace(/^[^A-Za-z0-9]+/,'').replace(/\s+/g,' ').trim();return `"${x.length>50?`${x.slice(0,50)}...`:x}"`;};
        const byHand=e=>{
          const now=quote(ps.items[e.index]?.text);
          if(e.kind==='heading'||e.kind==='caption')return `change ${e.kind==='caption'?'the caption':'the heading'} ${now} to read "${e.text}"`;
          if(e.kind==='question')return `change "${e.from}" to "${e.to}" in the question ${now}`;
          return `change the time line ${now} to read "${timeText(e.minutes,e.cumulative)}" (select the words after the small picture and type over them)`;
        };
        for(const e of edits){
          const p=ps.items[e.index];
          if(!p)continue;
          try{
            if(e.kind==='heading'||e.kind==='caption'){p.insertText(e.text,'Replace');await c.sync();continue;}
            // A question's number is replaced where it stands, so the rest of the question keeps its formatting.
            if(e.kind==='question'){const hit=p.search(e.from,{matchCase:true});hit.load('items');await c.sync();hit.items[0]?.insertText(e.to,'Replace');await c.sync();continue;}
            // The time block's icon is a picture at the start of the paragraph. Word refuses to give the range of an inline picture
            // (InlinePicture.getRange: InvalidArgument, seen in the author's Word 2026-11-03), so the words are found by searching for
            // "Time:" and replaced up to the end of the paragraph, which leaves the picture and the space after it alone.
            const text=timeText(e.minutes,e.cumulative),hit=p.search('Time:',{matchCase:true});
            hit.load('items');await c.sync();
            if(hit.items.length)hit.items[0].expandTo(p.getRange('Content').getRange('End')).insertText(text,'Replace');
            else p.insertText(text,'Replace');
            await c.sync();
          }catch(err){
            refused.push({manual:byHand(e),code:[err?.code,err?.debugInfo?.errorLocation&&`in ${err.debugInfo.errorLocation}`].filter(Boolean).join(' ')||String(err?.message||err)});
          }
        }
        if(refused.length){
          const codes=[...new Set(refused.map(r=>r.code))].join('; '),shown=refused.slice(0,4).map(r=>r.manual),more=refused.length>4?`, and ${refused.length-4} more like them`:'';
          throw new Error(`Word would not make ${refused.length} of ${edits.length} changes to numbers and times, and made the rest. To finish by hand: ${shown.join('; ')}${more}. (Word's code: ${codes}.)`);
        }
      });
    },
    // The document as Word would save it, so the review sees what a saved file would hold.
    async reviewParts(){
      const [ooxml,title]=await Promise.all([
        run(async c=>{const o=c.document.body.getOoxml();await c.sync();return o.value;}),
        this.title()
      ]);
      const doc=new DOMParserClass().parseFromString(ooxml,'application/xml');
      const part=name=>Array.from(doc.getElementsByTagNameNS(PKG,'part')).find(p=>p.getAttributeNS(PKG,'name')===name);
      const inner=name=>Array.from(part(name)?.getElementsByTagNameNS(PKG,'xmlData')[0]?.childNodes||[]).find(n=>n.nodeType===1)||null;
      const root=inner('/word/document.xml');
      if(!root)throw new Error('Word did not return the document for review. Try again.');
      return {root,stylesRoot:inner('/word/styles.xml'),title,language:officeApi()?.context?.contentLanguage||''};
    },
    // The whole document as one flat Office package (its text, styles, lists and pictures), for making the learner workbook.
    async sourcePackage(){
      return run(async c=>{const o=c.document.body.getOoxml();await c.sync();return o.value;});
    },
    // Opens a finished .docx (as base64) in a new Word window. Word can refuse (an older Word, or a file it dislikes).
    async openDocument(base64){
      return run(async c=>{const d=c.application.createDocument(base64);d.open();await c.sync();return true;});
    },
    // The whole open document as a .docx (bytes), exactly as Word would save it now, unsaved changes included. For a backup.
    async captureFile(){
      const document=officeApi()?.context?.document;
      if(!document?.getFileAsync)throw new Error('Word is not ready yet. Close the OpenStage Guide pane, open it again from the ribbon, then try again.');
      return readCompressedDocument(document);
    },
    // Word's own PDF of the open document, as base64, read a slice at a time. Word may refuse, or make a PDF without tags.
    async exportPdf(){
      const office=officeApi(),document=office?.context?.document;
      if(!document?.getFileAsync)throw new Error('Word is not ready yet. Close the OpenStage Guide pane, open it again from the ribbon, then try again.');
      const call=action=>new Promise((resolve,reject)=>{try{action(r=>r.status==='succeeded'?resolve(r.value):reject(new Error(r.error?.message||'Word could not make the PDF.')));}catch(e){reject(e);}});
      const file=await call(cb=>document.getFileAsync(office.FileType?.Pdf||'pdf',{sliceSize:65536},cb));
      try{
        if(!Number.isSafeInteger(file.size)||file.size<=0||file.size>64*1024*1024||file.sliceCount<1||file.sliceCount>1024)throw new Error('Word made a PDF this add-in cannot read (empty, or over 64 MB).');
        const bytes=new Uint8Array(file.size);let offset=0;
        for(let i=0;i<file.sliceCount;i++){
          const slice=await call(cb=>file.getSliceAsync(i,cb));
          if(!slice.data||offset+slice.data.length>bytes.length)throw new Error('Word returned an incomplete PDF. Nothing was saved.');
          bytes.set(slice.data,offset);offset+=slice.data.length;
        }
        if(offset!==bytes.length)throw new Error('Word returned an incomplete PDF. Nothing was saved.');
        return toBase64(bytes);
      }finally{await call(cb=>file.closeAsync(cb)).catch(()=>{});}
    },
    // The same file as an ordinary download, for when Word cannot open it (or for a PDF).
    offerDownload(base64,fileName,type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'){
      const bytes=Uint8Array.from(atob(base64),ch=>ch.charCodeAt(0));
      const url=URL.createObjectURL(new Blob([bytes],{type}));
      const a=document.createElement('a');a.href=url;a.download=fileName;document.body.append(a);a.click();a.remove();
      setTimeout(()=>URL.revokeObjectURL(url),60000);
    },
    // Bring the contents table up to date, where this Word can.
    async refreshFields(){
      try{
        if(!officeApi()?.context?.requirements?.isSetSupported('WordApi','1.5'))return false;
        await run(async c=>{
          const f=c.document.body.fields;f.load('items/code');await c.sync();
          for(const x of f.items)if(/\bTOC\b/.test(x.code))x.updateResult();
          await c.sync();
        });
        return true;
      }catch{return false;}
    }
  };
}

// An in-memory document with the same interface, for the browser preview and for tests.
// quirks.lastParagraphLosesStyle reproduces what real Word did in the first native test: the last paragraph an insert
// brings in keeps its words but loses its style, and joins the paragraph that follows.
// The author's pictures in the in-memory document: id -> {mime, base64}. What a picture is is checked here, once, however it got here.
const PICTURE_MIMES={'image/png':'png','image/jpeg':'jpg','image/gif':'gif'};
const PICTURE_B64=/^[A-Za-z0-9+/]*={0,2}$/;
const iconRid=icon=>`rIdIcon_${icon.replace(/[^a-z0-9]/gi,'_')}`;

// `icons(names)` gives the real pictures of the cue icons the guide uses (base64 by name), for the file a Word document is made from;
// without it (the tests) a one-pixel stand-in is used.
export function createDemoAdapter({DOMParserClass=globalThis.DOMParser,SerializerClass=globalThis.XMLSerializer,quirks={},icons=null}={}){
  const parseXml=text=>new DOMParserClass().parseFromString(text,'application/xml');
  // The styles part: the program's own definitions with the author's edits (guide-styles-edit.mjs) laid over them, and the names those give.
  let styleEdits={overrides:{},custom:[]};
  const styleIo={parse:parseXml,serialize:el=>new SerializerClass().serializeToString(el)};
  const currentStyles=()=>applyStyleEdits(stylesPart(),styleEdits,styleIo);
  let styles=readStyles(parseXml(stylesPart()));
  const refreshStyles=()=>{styles=readStyles(parseXml(currentStyles()));};
  let media=Object.create(null),paras=[],cursor=-1,docTitle='',autoShow=false,selectionHandler=null,ownCues=[],activeTheme,lookText=null,slideMemory=null,versionMemory=null,formsText=null,designText=null,sessionsText=null,demoRefusals=null,picker=null;
  const applied=[],headerFooter=[];
  const history=[],future=[],opened=[],downloaded=[];
  let capturedFile=null;
  // The steps Undo goes back through, each a copy of the paragraphs as they were. A change made in a group (typing in one paragraph) is one
  // step for as long as it goes on in the same group; any other change starts a new step. A new step ends what Redo could bring back.
  const HISTORY_LIMIT=300,GROUP_MS=2000;
  let lastGroup=null,lastAt=0;
  const copyState=()=>({paras:paras.map(p=>({...p})),cursor,docTitle,styleEdits:JSON.parse(JSON.stringify(styleEdits))});
  const remember=group=>{
    const now=Date.now();
    future.length=0;
    if(group&&group===lastGroup&&now-lastAt<GROUP_MS&&history.length){lastAt=now;return;}
    history.push(copyState());
    if(history.length>HISTORY_LIMIT)history.shift();
    lastGroup=group||null;lastAt=now;
  };
  // What a paragraph carries that its words and style do not say: the names of its bookmarks, and the form citations in it (a content
  // control tagged osg-ref:<id> and the words it holds). Forms and citations are read from these (guide-forms.mjs).
  const marksOf=el=>{
    const bookmarks=Array.from(el.getElementsByTagNameNS(W,'bookmarkStart')).map(b=>b.getAttributeNS(W,'name')).filter(Boolean);
    const refs=[];
    for(const sdt of Array.from(el.getElementsByTagNameNS(W,'sdt'))){
      const kids=Array.from(sdt.childNodes).filter(n=>n.nodeType===1);
      const tag=kids.find(n=>n.localName==='sdtPr')&&Array.from(kids.find(n=>n.localName==='sdtPr').childNodes).find(n=>n.nodeType===1&&n.localName==='tag')?.getAttributeNS(W,'val');
      if(!tag||!tag.startsWith(REF_TAG_PREFIX))continue;
      const content=kids.find(n=>n.localName==='sdtContent');
      refs.push({id:tag.slice(REF_TAG_PREFIX.length),text:content?textOf(content):''});
    }
    return bookmarks.length||refs.length?{bookmarks,refs}:undefined;
  };
  // Rewrites the words of the nth citation of a form inside a paragraph's XML (the first run's look is kept).
  const retextCitation=(xml,id,nth,text)=>{
    let seen=0;
    return xml.replace(/<w:sdt(?:\s[^>]*)?>[\s\S]*?<\/w:sdt>/g,block=>{
      if(!block.includes(`w:val="${REF_TAG_PREFIX}${id}"`)||seen++!==nth)return block;
      const at=block.indexOf('<w:sdtContent>');
      if(at<0)return block;
      const rPr=/<w:rPr>[\s\S]*?<\/w:rPr>/.exec(block.slice(at))?.[0]||'';
      return block.replace(/<w:sdtContent>[\s\S]*<\/w:sdtContent>/,`<w:sdtContent><w:r>${rPr}<w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:sdtContent>`);
    });
  };
  // Reads a paragraph's words and marks again from its XML, after an edit changed the XML.
  const refresh=p=>{
    const el=parseXml(p.xml).documentElement;
    p.text=textOf(el);
    const m=marksOf(el);
    if(m)p.marks=m;else delete p.marks;
  };
  const entry=(el,{marks=true}={})=>{
    const id=Array.from(el.childNodes).filter(n=>n.nodeType===1&&n.localName==='pPr').flatMap(n=>Array.from(n.childNodes)).find(n=>n.nodeType===1&&n.localName==='pStyle')?.getAttributeNS(W,'val')||'Normal';
    const m=marks?marksOf(el):undefined;
    return {xml:new SerializerClass().serializeToString(el),text:textOf(el),style:styles.get(id)?.name||id,builtIn:id==='Title'?'Title':undefined,...(m?{marks:m}:{})};
  };
  // A table is read the way Word reports it: its cells' paragraphs, one after another. The whole table's XML rides on
  // the first of them (the rest carry none), so the document can still be reviewed as Word would save it, and
  // taking the paragraphs away takes the table with them.
  const parseBody=body=>Array.from(parseXml(`<w:document ${NS}><w:body>${body}</w:body></w:document>`).getElementsByTagNameNS(W,'body')[0].childNodes).filter(n=>n.nodeType===1&&(n.localName==='p'||n.localName==='tbl')).flatMap(n=>{
    if(n.localName==='p')return [entry(n)];
    // A table's later cells carry no XML of their own, so no marks are read from a table: forms and citations belong outside tables.
    const cells=Array.from(n.getElementsByTagNameNS(W,'p')).map(c=>entry(c,{marks:false}));
    cells.forEach((c,i)=>{c.xml=i===0?new SerializerClass().serializeToString(n):'';});
    return cells;
  });
  const retext=(xml,text)=>xml.replace(/(<w:t[^>]*>)[^<]*(<\/w:t>)/,`$1${esc(text)}$2`);
  // The pictures a block was made with (each with the id its XML names) are kept under ids that no other block's pictures have, so two
  // blocks that each had "their first picture" do not share one. What is not a picture of a kind Word holds is refused before anything changes.
  // The pictures are checked and kept first (`registerPictures` says which new id each took); the ids in the XML are changed after, in as many pieces of XML as share them.
  const keepPictures=(body,pics)=>{
    if(!pics.length)return body;
    return renamePictures(body,registerPictures(pics));
  };
  const renamePictures=(text,renamed)=>{
    let out=text;
    for(const r of renamed)out=out.split(`r:embed="${r.from}"`).join(`r:embed="${r.id}"`);
    return out;
  };
  const registerPictures=pics=>{
    const renamed=[];
    for(const p of pics){
      if(!p||typeof p.rid!=='string'||!/^rIdPic[A-Za-z0-9_-]{1,60}$/.test(p.rid)||!PICTURE_MIMES[p.mime]||typeof p.base64!=='string'||!p.base64||!PICTURE_B64.test(p.base64)||p.base64.length%4!==0)throw new Error('A picture could not be added: it is not a PNG, JPG or GIF.');
      let id;
      do id=`rIdPic${Math.random().toString(36).slice(2,10)}`;while(media[id]||renamed.some(r=>r.id===id));
      renamed.push({id,from:p.rid,mime:p.mime,base64:p.base64});
    }
    for(const r of renamed)media[r.id]={mime:r.mime,base64:r.base64};
    return renamed;
  };
  // The pictures the paragraphs still use (a picture whose paragraph was deleted is dropped when the document is saved).
  const usedMedia=()=>{
    const used=new Set();
    for(const p of paras)for(const m of String(p.xml||'').matchAll(/r:embed="(rIdPic[A-Za-z0-9_-]+)"/g))used.add(m[1]);
    return Object.fromEntries([...used].filter(id=>media[id]).map(id=>[id,{...media[id]}]));
  };
  return {
    kind:'demo',
    async paragraphs(){return paras.map(({text,style,builtIn,marks})=>({text,style,builtIn,...(marks?{marks}:{})}));},
    async title(){return docTitle;},
    async setTitle(t){docTitle=t;},
    async customCues(){return [...ownCues];},
    async setCustomCues(names){ownCues=[...names];},
    // Looks: the demo document keeps what it was given, so a test can see exactly what Word would have been asked for.
    setTheme(t){activeTheme=t||undefined;},
    async slideMemory(){return slideMemory;},
    async formSettings(){return formsText;},
    async setFormSettings(text){formsText=text;return true;},
    // The design desk's notes (src/design-model.mjs): the needs analysis, the objectives, the storyboard and the evaluation plan, as the text the project file keeps.
    async design(){return designText;},
    async setDesign(text){designText=text===undefined?null:text;return true;},
    // What was made from the guide and the card decks being edited (src/session-record.mjs), as the text the project file keeps.
    async sessions(){return sessionsText;},
    async setSessions(text){sessionsText=text===undefined?null:text;return true;},
    async setSlideMemory(text){slideMemory=text;return true;},
    async versionMemory(){return versionMemory;},
    async setVersionMemory(text){versionMemory=text;return true;},
    async getLook(){return lookText;},
    async setLook(text){lookText=text;},
    // The preview uses plain file inputs; a test (or a quirk) can stand in for the author's choice.
    setPicker(fn){picker=fn;},
    async pickFiles(opts){
      if(picker)return picker(opts);
      return pickWithBrowserInputs(opts.kind,{accept:opts.accept,multiple:opts.multiple,readDeck:async f=>readPptx(await f.arrayBuffer(),f.name,globalThis.JSZip,{details:true})});
    },
    // A look is a set of settings for styles, by name. Here it is laid over the styles as edits (one step of Undo), so the guide, the Review and
    // the Word file made from it all see it, and the Styles panel shows it.
    async applyLook(specs){
      const refused=(demoRefusals?.(specs))||[];
      applied.push(specs);
      const byName=new Map(listStyles(currentStyles(),styleIo).map(x=>[x.name.toLowerCase(),x.id]));
      const ALIGN={Centered:'center',Left:'left',Right:'right',Justified:'justify'};
      const overrides={...styleEdits.overrides};
      const done=[],skipped=[];
      for(const spec of specs){
        const id=byName.get(String(spec.name).toLowerCase());
        if(!id){skipped.push(spec.name);continue;}
        const mine={};
        const f=spec.font||{};
        const font={};
        for(const k of ['name','size','color','bold','italic'])if(f[k]!==undefined)font[k]=f[k];
        if(Object.keys(font).length)mine.font=font;
        const p=spec.paragraph||{};
        const para={};
        for(const k of ['spaceBefore','spaceAfter','leftIndent','firstLineIndent','keepWithNext','keepTogether'])if(p[k]!==undefined)para[k]=p[k];
        if(p.lineSpacing!==undefined)para.lineSpacing=Math.round(p.lineSpacing/12*100)/100;
        if(p.alignment&&ALIGN[p.alignment])para.alignment=ALIGN[p.alignment];
        if(Object.keys(para).length)mine.paragraph=para;
        if(spec.shading)mine.shading=spec.shading;
        const borders=(spec.borders||[]).map(b=>({location:String(b.location).toLowerCase(),color:b.color,width:(({Pt025:0.25,Pt050:0.5,Pt075:0.75,Pt100:1,Pt150:1.5,Pt225:2.25,Pt300:3,Pt450:4.5,Pt600:6})[b.width])??(Number(b.width)||1)}));
        if(borders.length)mine.borders=borders;
        if(Object.keys(mine).length){overrides[id]={...(overrides[id]||{}),...mine};}
        done.push(spec.name);
      }
      const next=cleanStyleEdits({...styleEdits,overrides});
      applyStyleEdits(stylesPart(),next,styleIo);
      remember();
      styleEdits=next;
      refreshStyles();
      return {applied:done,skipped,refused};
    },
    // The styles, as the Styles panel shows them, and what the author has changed. A style of the author's own is only taken away when no
    // paragraph uses it (the paragraphs would be left with a style that is not there).
    listStyles(){return listStyles(currentStyles(),styleIo,new Set(styleEdits.custom.map(c=>c.id)));},
    styleEdits(){return JSON.parse(JSON.stringify(styleEdits));},
    styleUse(){
      // How many paragraphs use each style, as the style of the paragraph or as the style of some of its words.
      const use={};
      for(const p of paras){
        const ids=new Set([...String(p.xml||'').matchAll(/<w:(?:pStyle|rStyle) w:val="([^"]+)"/g)].map(m=>m[1]));
        for(const id of ids)use[id]=(use[id]||0)+1;
      }
      return use;
    },
    async setStyleEdits(edits,{group}={}){
      const next=cleanStyleEdits(edits);
      applyStyleEdits(stylesPart(),next,styleIo);
      const keep=new Set(next.custom.map(c=>c.id));
      const use=this.styleUse();
      for(const c of styleEdits.custom)if(!keep.has(c.id)&&use[c.id])throw new Error(`The style ${c.name} is used by ${use[c.id]} paragraph${use[c.id]===1?'':'s'}, so it stays until none of them uses it.`);
      remember(group);
      styleEdits=next;
      refreshStyles();
      return true;
    },
    async setHeaderFooter(hf){headerFooter.push(hf);return {done:['header','footer'].filter(k=>hf[k]!==undefined),failed:[]};},
    async cursorIndex(){return cursor<0?paras.length-1:cursor;},
    async insert(build,{where='after',after,select='last',spare=true,remove=null,pictures=null}={}){
      const made=makeContext();
      let body=build(made);
      body=keepPictures(body,pictures||made.pictures||[]);
      // The steps of an activity are a bulleted list: in the guide they share the one bulleted list, so the program's Word file shows bullets, not letters.
      for(const n of made.bullets||[])body=body.split(`<w:numId w:val="${n}"/>`).join(`<w:numId w:val="${BULLET_NUM_ID}"/>`);
      if(remove&&norm(paras[remove.start]?.text)!==norm(remove.guard))throw new Error('That block has changed since you chose it. Select it again, then choose Edit this block.');
      const at=where==='start'||!paras.length?0:after+1;
      if(where!=='start'&&paras.length&&!paras[after])throw new Error('That part of the document is no longer there. Choose Refresh and try again.');
      // The last paragraph in the package is the spare one when there is a spare; it is empty, so nothing is lost with it.
      const added=parseBody(body+(spare?SPARE_PARAGRAPH:'')),last=added.pop();
      if(!spare){
        if(!quirks.lastParagraphLosesStyle)added.push(last);
        else if(last.text.trim()){
          const next=paras[at];
          if(next)next.text=last.text+next.text;
          else added.push({...last,style:'Normal'});
        }
      }
      remember();
      paras.splice(at,0,...added);
      const gone=remove?remove.end-remove.start+1:0;
      if(remove)paras.splice(remove.start,gone);
      // The real document always ends in a plain empty paragraph after an insert that reaches its end (see insert() above).
      const tail=paras.at(-1);
      if(tail&&/^OSG /i.test(tail.style||'')){
        if(norm(tail.text))paras.push({xml:'',text:'',style:'Normal'});
        else tail.style='Normal';
      }
      cursor=(select==='first'?at:at+added.length-1)-gone;
    },
    async plainLineAfter(index){
      if(!paras[index])throw new Error('That part of the document is no longer there. Choose Refresh and try again.');
      const next=paras[index+1];
      if(next&&!norm(next.text)&&next.style==='Normal'){cursor=index+1;return {at:index+1,created:false};}
      remember();
      paras.splice(index+1,0,{xml:'<w:p xmlns:w="'+W+'"/>',text:'',style:'Normal'});
      cursor=index+1;
      return {at:index+1,created:true};
    },
    async go(index,expect){
      let at=index;
      if(expect!==undefined&&norm(expect)&&(!paras[at]||norm(paras[at].text)!==norm(expect))){
        const hits=paras.map((p,i)=>norm(p.text)===norm(expect)?i:-1).filter(i=>i>=0).sort((a,b)=>Math.abs(a-index)-Math.abs(b-index));
        if(!hits.length)throw new Error('That part of the document has changed since this list was made. Choose Review again.');
        at=hits[0];
      }
      if(!paras[at])throw new Error('That part of the document is no longer there. Choose Refresh and try again.');
      cursor=at;
    },
    // Delete or move paragraphs start..end after checking that the first still reads as `guard`. `after` is the paragraph the block goes
    // after (-1 is the top of the document).
    async deleteRange({start,end,guard}){
      if(!paras[start]||!paras[end]||end<start||norm(paras[start].text)!==norm(guard))throw new Error(CHANGED);
      remember();
      paras.splice(start,end-start+1);
      cursor=Math.max(0,Math.min(start,paras.length-1));
    },
    async moveRange({start,end,guard,after,group}){
      if(!paras[start]||!paras[end]||end<start||norm(paras[start].text)!==norm(guard))throw new Error(CHANGED);
      if(after>=start&&after<=end)throw new Error('A part of the guide cannot be moved into itself.');
      if(after>=paras.length)throw new Error('That part of the document is no longer there. Choose Refresh and try again.');
      remember(group);
      const block=paras.splice(start,end-start+1),at=(after>end?after-block.length:after)+1;
      paras.splice(at,0,...block);
      cursor=at;
      return {at};
    },
    async keepPaneOpen({manifestSupportsIt=false}={}){if(manifestSupportsIt===autoShow)return false;autoShow=manifestSupportsIt;return manifestSupportsIt;},
    onSelectionChanged(handler){selectionHandler=handler;return true;},
    async applyEdits(edits,{group}={}){
      remember(group);
      for(const e of edits){
        const p=paras[e.index];if(!p)continue;
        if(e.kind==='question'){p.text=p.text.replace(e.from,e.to);p.xml=p.xml.replace(e.from,e.to);continue;}
        if(e.kind==='form'||e.kind==='caption'){p.xml=retext(p.xml,e.text);refresh(p);continue;}
        if(e.kind==='citation'){p.xml=retextCitation(p.xml,e.id,e.nth,e.to);refresh(p);continue;}
        const text=e.kind==='heading'?e.text:` ${timeText(e.minutes,e.cumulative)}`;
        p.xml=retext(p.xml,text);p.text=e.kind==='heading'?e.text:text;
      }
    },
    async reviewParts(){
      return {root:parseXml(`<w:document ${NS}><w:body>${paras.map(p=>p.xml).join('')}</w:body></w:document>`),stylesRoot:parseXml(currentStyles()),title:docTitle,language:'en-US'};
    },
    // A flat Office package like the one Word returns for the whole body: the content, the styles, the lists the content
    // uses and a small picture for every icon it refers to.
    async sourcePackage(){
      const body=paras.map(p=>p.xml).join('');
      const ids=[...new Set([...body.matchAll(/r:embed="([^"]+)"/g)].map(m=>m[1]))];
      const lists=Math.max(0,...[...body.matchAll(/<w:numId w:val="(\d+)"/g)].map(m=>Number(m[1])).filter(n=>n!==BULLET_NUM_ID));
      const bulleted=body.includes(`<w:numId w:val="${BULLET_NUM_ID}"/>`);
      const part=(name,type,inner)=>`<pkg:part pkg:name="${name}" pkg:contentType="${type}"><pkg:xmlData>${inner}</pkg:xmlData></pkg:part>`;
      const png='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
      // The cue icons the guide uses are the real pictures when the program can load them; the author's own pictures are always the real ones.
      const iconNames=ICON_FILES.filter(n=>ids.includes(iconRid(n)));
      let real={};
      if(icons&&iconNames.length){try{real=await icons(iconNames);}catch{real={};}}
      const files=ids.map((id,i)=>{
        const pic=media[id];
        if(pic)return {id,name:`image${i+1}.${PICTURE_MIMES[pic.mime]}`,type:pic.mime,data:pic.base64};
        const icon=iconNames.find(n=>iconRid(n)===id);
        return {id,name:`image${i+1}.png`,type:'image/png',data:(icon&&real[icon])||png};
      });
      const rels=`<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${lists||bulleted?'<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>':''}${files.map(f=>`<Relationship Id="${f.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${f.name}"/>`).join('')}`;
      return `<pkg:package xmlns:pkg="${PKG}">${part('/word/document.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',`<w:document ${NS}><w:body>${body}</w:body></w:document>`)}${part('/word/_rels/document.xml.rels','application/vnd.openxmlformats-package.relationships+xml',`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`)}${part('/word/styles.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',currentStyles())}${lists||bulleted?part('/word/numbering.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',numberingPart(lists,new Set(),bulleted)):''}${files.map(f=>`<pkg:part pkg:name="/word/media/${f.name}" pkg:contentType="${f.type}" pkg:compression="store"><pkg:binaryData>${f.data}</pkg:binaryData></pkg:part>`).join('')}</pkg:package>`;
    },
    // The preview has no second window: it keeps what it was asked to open, for the tests to look at.
    async openDocument(base64,fileName){
      if(quirks.cannotOpenDocument)throw new Error('This Word cannot open a new document from here.');
      opened.push({base64,fileName});return true;
    },
    // A small PDF, tagged unless the preview is told Word would not tag it.
    async exportPdf(){
      if(quirks.cannotExportPdf)throw new Error('Word could not make the PDF.');
      const catalogue=quirks.untaggedPdf?'<< /Type /Catalog >>':'<< /Type /Catalog /MarkInfo << /Marked true >> /StructTreeRoot 5 0 R /Lang (en-US) >>';
      return btoa(`%PDF-1.7\n1 0 obj ${catalogue} endobj\n2 0 obj << /Title (${docTitle||'Untitled'}) >> endobj\ntrailer << /Root 1 0 R /Info 2 0 R >>\n%%EOF`);
    },
    offerDownload(base64,fileName,type){downloaded.push({base64,fileName,type});},
    openedDocuments(){return opened;},
    downloads(){return downloaded;},
    // Rewrites the contents table the way Word's field update would: one entry per Heading 1/2, Module and Lesson.
    async refreshFields(){
      const level=s=>/^(heading 1|osg module|osg assessment)$/i.test(s||'')?1:/^(heading 2|osg lesson)$/i.test(s||'')?2:0;
      const isToc=p=>/^toc [1-9]$/i.test(p.style||'')||/update this table to list the modules and lessons/i.test(p.text||'');
      const first=paras.findIndex(isToc);
      if(first<0)return true;
      const entries=paras.filter(p=>level(p.style)).map(p=>({xml:'',text:`${p.text.trim()}\t1`,style:`TOC ${level(p.style)}`}));
      const kept=paras.filter(p=>!isToc(p)),removed=paras.length-kept.length;
      paras=[...kept.slice(0,first),...entries,...kept.slice(first)];
      if(cursor>=first)cursor+=entries.length-removed;
      return true;
    },
    // For tests and the preview only. The demo document has no .docx of its own: a test hands over the bytes a backup should see.
    setCapturedFile(bytes){capturedFile=bytes;},
    async captureFile(){
      if(!capturedFile)throw new Error('The browser preview has no Word document to back up. Open the guide from Word.');
      return capturedFile;
    },
    snapshot(){return paras.map(p=>({text:p.text,style:p.style}));},
    cursor(){return cursor;},
    seed(items){paras=items.map(p=>({xml:'',text:p.text,style:p.style}));cursor=paras.length-1;},
    // For the program's project file (guide-project.mjs): everything this document holds as plain data, and a way to put it back.
    // importState takes what cleanGuide returned. The Word-format XML of every paragraph is checked first, so a damaged project
    // is refused whole and the document is left as it was.
    // A picture the guide holds, by the id its paragraphs name, or null.
    mediaOf(id){return media[id]||null;},
    exportState(){
      const copy=v=>JSON.parse(JSON.stringify(v));
      return {title:docTitle,paragraphs:paras.map(({xml,text,style,builtIn})=>({xml,text,style,...(builtIn?{builtIn}:{})})),cursor,ownCues:[...ownCues],look:lookText,slideMemory,versionMemory,forms:formsText,design:designText,sessions:sessionsText,applied:copy(applied),headerFooter:copy(headerFooter),media:usedMedia(),styles:copy(styleEdits)};
    },
    importState(guide){
      // A guide file is data a person may be sent: what could make Word reach outside the file (fields other than the program's own, links,
      // objects, addresses) is taken out of every paragraph here, before the guide holds it (src/guide-xml-clean.mjs).
      const cleaned=guide.paragraphs.map(p=>p.xml?cleanParagraphXml(p.xml,styleIo).xml:p.xml);
      const marksList=guide.paragraphs.map((p,i)=>{
        if(!p.xml)return undefined;
        let d=null;
        try{d=parseXml(cleaned[i]);}catch{d=null;}
        // Every paragraph is a Word paragraph or table. A browser's own parser also flags unbalanced tags with a parsererror.
        if(!d||!d.documentElement||d.getElementsByTagName('parsererror').length||!['p','tbl'].includes(d.documentElement.localName))throw new Error(`Paragraph ${i+1} of the project is damaged, so the project was not opened.`);
        return d.documentElement.localName==='p'?marksOf(d.documentElement):undefined;
      });
      paras=guide.paragraphs.map((p,i)=>({...p,xml:cleaned[i],...(marksList[i]?{marks:marksList[i]}:{})}));cursor=guide.cursor;docTitle=guide.title;ownCues=[...guide.ownCues];
      lookText=guide.look;slideMemory=guide.slideMemory;versionMemory=guide.versionMemory;formsText=guide.forms??null;designText=guide.design??null;sessionsText=guide.sessions??null;media=Object.create(null);Object.assign(media,guide.media||{});styleEdits=cleanStyleEdits(guide.styles);refreshStyles();
      applied.length=0;applied.push(...guide.applied);headerFooter.length=0;headerFooter.push(...guide.headerFooter);history.length=0;future.length=0;lastGroup=null;
    },
    autoShow(){return autoShow;},
    setAutoShow(v){autoShow=!!v;},
    // What the pane asked of the styles and the header and footer; refuse(fn) makes Word "refuse" the settings fn returns.
    lookState(){return {theme:activeTheme,applied,headerFooter,lookText};},
    refuse(fn){demoRefusals=fn;},
    // What the author does in Word itself: click somewhere, type over a paragraph's words, press Ctrl+Z.
    place(index){cursor=index;},
    typeOver(index,text){remember();paras[index].text=text;paras[index].xml=retext(paras[index].xml,text);},
    undo(){
      const was=history.pop();
      if(was){future.push(copyState());paras=was.paras;cursor=was.cursor;docTitle=was.docTitle;if(was.styleEdits){styleEdits=was.styleEdits;refreshStyles();}}
      lastGroup=null;
      return !!was;
    },
    redo(){
      const next=future.pop();
      if(next){history.push(copyState());paras=next.paras;cursor=next.cursor;docTitle=next.docTitle;if(next.styleEdits){styleEdits=next.styleEdits;refreshStyles();}}
      lastGroup=null;
      return !!next;
    },
    historyState(){return {undo:history.length,redo:future.length};},
    // The editor's one way to change paragraphs: replace `count` paragraphs from `start` by the paragraphs the given Word-format XML makes.
    // One edit is count 1 with one paragraph; Enter is 1 with 2; Backspace joining is 2 with 1; a new line is 0 with 1. `guard` is the words the
    // first replaced paragraph had when the editor last looked, so a stale row is refused and never applied to the wrong paragraph. Tables are
    // not touched (their cells carry no XML of their own). `group` makes a run of edits one step of Undo. `select` is where the cursor goes.
    // Replace the table that starts at paragraph `start` by another table (the same table with a cell's words changed, or a row more or less).
    // The table's first entry holds all its XML and the entries after it that have none are its other cells, so all of them are replaced
    // by the cells of the new table. `guard` is the words of the first cell the editor last saw.
    async replaceTable({start,guard,xml,select,group,pictures=null}){
      const first=paras[start];
      if(!Number.isInteger(start)||!first||!/^<w:tbl[\s>]/.test(first.xml||'')||norm(first.text)!==norm(guard))throw new Error(CHANGED);
      let d=null;
      try{d=parseXml(xml);}catch{d=null;}
      if(typeof xml!=='string'||!d||!d.documentElement||d.getElementsByTagName('parsererror').length||d.documentElement.localName!=='tbl')throw new Error('That table could not be made.');
      // Pictures that came with the table (one put in a cell) are kept under ids of their own, as they are for a block.
      if(pictures?.length)xml=renamePictures(xml,registerPictures(pictures));
      const added=parseBody(xml);
      if(!added.length||!/^<w:tbl[\s>]/.test(added[0].xml||''))throw new Error('That table could not be made.');
      let end=start+1;
      while(end<paras.length&&!paras[end].xml)end++;
      remember(group);
      paras.splice(start,end-start,...added);
      cursor=Math.max(0,Math.min(select===undefined?start:select,paras.length-1));
      return {at:start,count:added.length};
    },
    async replaceParagraphs({start,count,guard,xmls,select,group,pictures=null}){
      if(!Number.isInteger(start)||!Number.isInteger(count)||start<0||count<0||start+count>paras.length||!Array.isArray(xmls))throw new Error(CHANGED);
      const old=paras.slice(start,start+count);
      if(count&&norm(old[0].text)!==norm(guard))throw new Error(CHANGED);
      if(old.some(p=>!p.xml||/^<w:tbl[\s>]/.test(p.xml)))throw new Error('A table is changed with the tools beside the guide.');
      for(const x of xmls){
        let d=null;
        try{d=parseXml(x);}catch{d=null;}
        if(!d||!d.documentElement||d.getElementsByTagName('parsererror').length||d.documentElement.localName!=='p')throw new Error('That paragraph could not be made.');
      }
      // Pictures that came with the paragraphs (copied from another guide) are kept under ids of their own; the ids are changed in every paragraph that names them.
      if(pictures?.length){const renamed=registerPictures(pictures);xmls=xmls.map(x=>renamePictures(x,renamed));}
      const added=xmls.length?parseBody(xmls.join('')):[];
      if(added.length!==xmls.length)throw new Error('That paragraph could not be made.');
      remember(group);
      paras.splice(start,count,...added);
      const at=select===undefined?start:select;
      cursor=Math.max(0,Math.min(at,paras.length-1));
      return {at:cursor};
    },
    fireSelection(){return selectionHandler?.();}
  };
}
