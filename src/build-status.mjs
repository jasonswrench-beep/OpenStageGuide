const HASH=/^[a-f0-9]{64}$/;
const summaries=Object.freeze({
  unchecked:'Not checked against the server.',
  matched:'Matched the server at last check; recheck before testing.',
  stale:'Loaded code differs from the server. Preserve work and reload normally before testing.',
  mixed:'Observed code identities differ. Preserve work and reload normally before testing; see support details.',
  unknown:'Build not verified; see support details.'
});

export function buildStatusView(buildId,status){
  const identity=typeof buildId==='string'&&HASH.test(buildId)?buildId:null;
  let state=status===undefined||status===null?'unchecked':typeof status==='string'&&Object.hasOwn(summaries,status)?status:'unknown';
  if(!identity&&state==='matched')state='unknown';
  const label=identity?`Loaded code ${identity.slice(0,16)}`:'Loaded build identity unavailable';
  return {status:state,text:`${label} · ${summaries[state]}`};
}

export function renderBuildStatus(loadedNode,checkNode,buildId,status,busy=false){
  const view=buildStatusView(buildId,status);
  loadedNode.textContent=view.text;
  loadedNode.dataset.status=view.status;
  checkNode.disabled=!!busy;
  return view;
}
