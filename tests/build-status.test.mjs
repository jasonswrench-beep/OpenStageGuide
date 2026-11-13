import test from 'node:test';
import assert from 'node:assert/strict';
import {buildStatusView,renderBuildStatus} from '../src/build-status.mjs';
import {inspectBuild} from '../src/diagnostics.mjs';

const low='1'.repeat(64),high='f'.repeat(64);
const entry=id=>`https://localhost:3843/build/${id}/src/app.mjs`;
const registry=id=>['app','model','word','document-store','diagnostics'].map(name=>`https://localhost:3843/build/${id}/src/${name}.mjs`);

test('unchecked header distinguishes loaded identity from a server check',()=>{
  for(const status of [undefined,null]){
    const view=buildStatusView(low,status);
    assert.equal(view.status,'unchecked');assert.match(view.text,/Loaded code 1111111111111111/);
    assert.match(view.text,/Not checked against the server/);assert.doesNotMatch(view.text,/Matched|Older|newer|reopen/);
  }
  assert.match(buildStatusView(null).text,/identity unavailable/);
});

test('matched header qualifies the last check and shows only the loaded prefix',()=>{
  const view=buildStatusView(high,'matched');
  assert.equal(view.status,'matched');assert.match(view.text,/Loaded code ffffffffffffffff/);
  assert.match(view.text,/at last check; recheck before testing/);assert.ok(!view.text.includes(high));
});

for(const [loaded,served,label] of [[low,high,'lexically higher server'],[high,low,'rollback or alternate server']]){
  test(`stale header reports a mismatch without chronology for ${label}`,async()=>{
    const requests=[];
    const report=await inspectBuild({entryURL:entry(loaded),registry:registry(loaded),fetcher:async path=>{
      requests.push(path);
      if(path!=='/health')throw new Error('Loaded snapshot unavailable in this controlled case');
      return {ok:true,headers:{get:()=>null},text:async()=>JSON.stringify({app:'OpenStage Guide',buildId:served})};
    }});
    assert.equal(report.status,'stale');assert.equal(report.loadedBuildId,loaded);assert.equal(report.servedBuildId,served);
    assert.equal(requests.length,2);
    const view=buildStatusView(report.loadedBuildId,report.status);
    assert.equal(view.status,'stale');assert.match(view.text,/differs from the server/);
    assert.match(view.text,/Preserve work and reload normally/);assert.doesNotMatch(view.text,/older|newer|reopen/i);
  });
}

test('mixed module identities warn without a current-build claim',async()=>{
  let requests=0;
  const report=await inspectBuild({entryURL:entry(low),registry:[...registry(low),entry(high)],fetcher:async()=>{requests++;throw new Error('Unexpected request');}});
  assert.equal(report.status,'mixed');assert.equal(requests,0);
  const view=buildStatusView(report.loadedBuildId,report.status);
  assert.equal(view.status,'mixed');assert.match(view.text,/Observed code identities differ/);
  assert.match(view.text,/Preserve work and reload normally/);assert.doesNotMatch(view.text,/Matched|older|newer|reopen/i);
});

test('unavailable and future states do not disclose arbitrary input or claim a match',()=>{
  for(const status of ['unknown','future_PRIVATE_STATUS',{},Symbol('PRIVATE_STATUS')]){
    const view=buildStatusView(null,status);
    assert.equal(view.status,'unknown');assert.match(view.text,/identity unavailable/);
    assert.match(view.text,/Build not verified/);assert.doesNotMatch(view.text,/PRIVATE|Matched/);
  }
  assert.equal(buildStatusView('PRIVATE_INVALID_ID','matched').status,'unknown');
  assert.doesNotMatch(buildStatusView('PRIVATE_INVALID_ID','matched').text,/PRIVATE|Matched/);
});

test('rendered header updates states and busy recheck control without retaining a stale badge',()=>{
  const loadedNode={textContent:'',dataset:{kept:'retained'}},checkNode={disabled:false};
  renderBuildStatus(loadedNode,checkNode,low,undefined);
  assert.equal(loadedNode.dataset.status,'unchecked');assert.equal(checkNode.disabled,false);
  renderBuildStatus(loadedNode,checkNode,low,'matched',true);
  assert.equal(loadedNode.dataset.status,'matched');assert.match(loadedNode.textContent,/at last check/);assert.equal(checkNode.disabled,true);
  renderBuildStatus(loadedNode,checkNode,low,'stale',false);
  assert.equal(loadedNode.dataset.status,'stale');assert.match(loadedNode.textContent,/differs/);assert.equal(checkNode.disabled,false);
  renderBuildStatus(loadedNode,checkNode,low,'mixed');
  assert.equal(loadedNode.dataset.status,'mixed');
  renderBuildStatus(loadedNode,checkNode,null,'unknown');
  assert.equal(loadedNode.dataset.status,'unknown');assert.doesNotMatch(loadedNode.textContent,/Matched|1111111111111111/);
  assert.equal(loadedNode.dataset.kept,'retained');assert.equal(checkNode.disabled,false);
});
