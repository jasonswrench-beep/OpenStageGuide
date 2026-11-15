// The top of the pane while a tool is open: the tool's name and what it is for, and "More tools", a short list of the tools that
// belong with it. It never lists all twenty-one (src/guide-tools.mjs says which belong together). It is only this strip; what
// the tool shows below it is the pane's own work (guide-pane.mjs for the guide's tools, app.mjs for the older panels).
import {toolFor,relatedTools,programName} from './guide-tools.mjs';

// Tools that name themselves on their own screen (Assessments, Review guide, a legacy panel) do not need the strip to say it again.
export const namesItself=tool=>!!(tool.ownHeading||tool.titlebar||tool.kind==='legacy');

export function mountToolStrip(doc,{open}){
  const $=id=>doc.getElementById(id);
  let current=null;
  function show(id){
    const tool=toolFor(id);
    if(!tool){hide();return;}
    current=id;
    const quiet=namesItself(tool);
    $('tool-strip').hidden=false;
    $('tool-title').hidden=quiet;
    // The name the tool has in its menu, so the menu, "More tools" and the heading all say the same.
    $('tool-title').textContent=programName(id);
    $('tool-note').hidden=quiet;
    $('tool-note').textContent=tool.note;
    const more=relatedTools(id);
    // "More tools" is a disclosure that opens on request and closes again each time a tool is opened, so a long list never sits in the way.
    $('tool-more').hidden=!more.length;
    $('tool-more').open=false;
    $('tool-more-list').replaceChildren(...more.map(t=>{
      const li=doc.createElement('li'),b=doc.createElement('button');
      const name=programName(t.id);
      b.type='button';b.id=`tool-more-${t.id}`;b.dataset.tool=t.id;b.textContent=name;
      b.setAttribute('aria-label',`${name}: ${t.tip}`);
      b.onclick=()=>open(t.id);
      li.append(b);return li;
    }));
  }
  function hide(){current=null;$('tool-strip').hidden=true;}
  return {show,hide,current:()=>current};
}
