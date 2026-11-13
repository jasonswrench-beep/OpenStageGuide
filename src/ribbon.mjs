import {NAVIGATION_ACTIONS,navigationTarget} from './guide-model.mjs';
import {TOOL_IDS} from './guide-tools.mjs';
// Every ribbon button runs the function named for its tool (src/guide-tools.mjs), and the pane opens that tool and no other: until
// 1.1.0.2 five of the nine buttons all opened the same long Guide page.
export const ROUTES=Object.freeze(Object.fromEntries(TOOL_IDS.map(id=>[id,id])));

// All commands operate on the same pane and its unsaved form values.
export function registerRibbonActions(office,{ready,go,onError}) {
  for(const [command,route] of Object.entries(ROUTES)) {
    office.actions.associate(command,async event=>{
      try {
        await ready;
        await office.addin.showAsTaskpane();
        await go(route);
      } catch(e) {onError(e);}
      finally {event?.completed();}
    });
  }
}

// The keyboard shortcuts that move about the guide. They work from the document, with the pane open or not, and say where they
// went (or why they could not go) in the pane's status line, so the pane need not be open for the cursor to move.
export function registerNavigationActions(office,{ready=Promise.resolve(),adapter,say,onError}){
  for(const [action,kind] of Object.entries(NAVIGATION_ACTIONS)){
    office.actions.associate(action,async event=>{
      try{
        await ready;
        const paras=await adapter.paragraphs(),at=await adapter.cursorIndex(),target=navigationTarget(paras,at,kind);
        if(target.index===undefined)say(target.message);
        else{await adapter.go(target.index,target.text);say(`Moved to ${target.text}.`);}
      }catch(e){onError(e);}
      finally{event?.completed();}
    });
  }
}
