// Local acceptance seam. Inert unless a disposable package carries an exact
// marker bound to its course ID AND the tester explicitly arms one boundary.
export const ACCEPTANCE_KEY='openstage.guide.acceptance.v1';
export const ACCEPTANCE_RUN='2026-10-25-native-boundaries';
export const ACCEPTANCE_TITLE='OpenStage Controlled Failure Acceptance';
export function marked(settings){
  const m=settings.get(ACCEPTANCE_KEY),p=settings.get('openstage.guide.project.v1');
  return !!(m?.disposable===true&&m.run===ACCEPTANCE_RUN&&m.projectId===p?.id&&p?.title===ACCEPTANCE_TITLE);
}
export function createAcceptanceFaults(settings){
  let armed=null;const events=[];
  const record=(event,extra={})=>{if(marked(settings())){events.push({event,...extra});if(events.length>100)events.shift();}};
  function active(kind){
    if(!armed)return false;
    if(!marked(settings())||settings().get('openstage.guide.project.v1')?.id!==armed.id){
      events.push({event:'scope-lost',kind:armed.kind});armed=null;
      throw new Error('Acceptance document identity changed. Synthetic experiment aborted before the native operation.');
    }
    return armed.kind===kind;
  }
  function hit(kind){if(!active(kind))return null;if(--armed.remaining!==0)return null;const point=armed;armed=null;record('synthetic-fault',{kind,when:point.when});return point;}
  const failure=()=>({status:'failed',error:{message:'SYNTHETIC acceptance boundary; not a Word-reported error'}});
  return {
    eligible:()=>marked(settings()),events:()=>structuredClone(events),
    clear(){if(armed)record('unused-arm-cleared');armed=null;},
    arm({kind,when,nth=1}){
      armed=null; // A new run never inherits an unconsumed earlier plan.
      if(!marked(settings()))throw new Error('Acceptance faults require a marked disposable document.');
      if(!['save','write-sync'].includes(kind)||!['before','after'].includes(when)||!Number.isInteger(nth)||nth<1)throw new Error('Invalid acceptance boundary.');
      armed={kind,when,remaining:nth,id:settings().get('openstage.guide.project.v1').id};
      record('armed',{kind,when,nth});
    },
    settings(){
      const s=settings();
      return {get:k=>s.get(k),set:(k,v)=>s.set(k,v),refreshAsync:cb=>s.refreshAsync(cb),saveAsync:cb=>{
        const point=hit('save');
        if(point?.when==='before'){record('real-save-skipped');queueMicrotask(()=>cb(failure()));return;}
        s.saveAsync(result=>{record('real-save-callback',{status:result.status});cb(point?.when==='after'&&result.status==='succeeded'?failure():result);});
      }};
    },
    async sync(ctx,kind='scan'){
      const point=kind==='write'?hit('write-sync'):null;
      if(point?.when==='before')throw new Error('SYNTHETIC acceptance exception before native write synchronization');
      await ctx.sync();
      if(point?.when==='after')throw new Error('SYNTHETIC acceptance exception after native write synchronization succeeded');
    }
  };
}
