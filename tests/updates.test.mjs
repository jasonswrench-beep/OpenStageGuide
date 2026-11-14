// Check for updates (scripts/updates.mjs): the comparison, the answers, the links, the one request GitHub is asked, and the switch that turns it off.
// Nothing here reaches the network: the request is a stand-in.
import test from 'node:test';
import assert from 'node:assert/strict';
import {EventEmitter} from 'node:events';
import {versionKey,compareVersions,spoken,releasePage,isReleasePage,answerFrom,fetchReleases,checkForUpdates,REPOSITORY,RELEASES_PAGE,RELEASES_API,USER_AGENT,MAX_ANSWER_BYTES,TIMEOUT_MS} from '../scripts/updates.mjs';
import {releaseLink,describeAnswer,RELEASES_PATH} from '../program/ui/updates.js';

const list=(...tags)=>JSON.stringify(tags.map(t=>typeof t==='string'?{tag_name:t}:t));

test('versions sort as releases do: alphas by number, then betas, then candidates, then the release; anything else is not a version',()=>{
  const order=['v0.1.0-alpha.9','v0.1.0-alpha.10','v0.1.0-beta.1','v0.1.0-rc.1','v0.1.0','v0.1.1','v0.2.0-alpha.1','v1.0.0'];
  for(let i=1;i<order.length;i++)assert.equal(compareVersions(versionKey(order[i]),versionKey(order[i-1])),1,`${order[i]} is after ${order[i-1]}`);
  assert.equal(compareVersions(versionKey('0.1.0'),versionKey('v0.1.0')),0,'the leading v does not matter');
  assert.equal(compareVersions(versionKey('v0.1.0-alpha.3'),versionKey('v0.1.0-alpha.3')),0);
  for(const bad of ['','latest','v1','v1.2','v1.2.3.4','v1.2.3-gamma.1','v1.2.3-alpha','v1.2.3-alpha.x','1.2.3-alpha.1; rm','v9999999.0.0',undefined,null])assert.equal(versionKey(bad),undefined,String(bad));
  assert.equal(spoken('v0.1.0-alpha.26'),'0.1.0 alpha 26');
  assert.equal(spoken('v1.0.0'),'1.0.0');
});

test('the answer: a newer release is told with its page; the same release is current; a copy with no release recorded is told the newest; drafts and odd tags are ignored',()=>{
  const all=list('v0.1.0-alpha.1','v0.1.0-alpha.2','not-a-version',{tag_name:'v9.9.9',draft:true},{tag_name:5},null,'x');
  const newer=answerFrom(all,'v0.1.0-alpha.1');
  assert.deepEqual([newer.state,newer.latest,newer.url],['newer','0.1.0 alpha 2',`${RELEASES_PAGE}/tag/v0.1.0-alpha.2`]);
  assert.match(newer.message,/OpenStage Guide 0\.1\.0 alpha 2 is out\. This copy is 0\.1\.0 alpha 1\..*keeps your guides/);
  const same=answerFrom(all,'v0.1.0-alpha.2');
  assert.deepEqual([same.state,same.message],['current','This copy is the newest release, 0.1.0 alpha 2.']);
  assert.equal(answerFrom(all,'v0.2.0').state,'current','a copy newer than anything published is not told to update');
  const unknown=answerFrom(all,undefined);
  assert.equal(unknown.state,'unknown');assert.match(unknown.message,/newest release of OpenStage Guide is 0\.1\.0 alpha 2.*not made by the setup/);
  assert.equal(answerFrom(all,'garbage').state,'unknown','a release that is not a version is no release');
  for(const bad of ['not json','{}','"x"','null','']){const r=answerFrom(bad,'v0.1.0');assert.equal(r.state,'error',bad);assert.equal(r.url,RELEASES_PAGE);}
  const none=answerFrom('[]','v0.1.0');
  assert.deepEqual([none.state,none.url],['error',RELEASES_PAGE]);assert.match(none.message,/No release of OpenStage Guide was found/);
});

test('only a page of OpenStage Guide\'s releases is ever given, by the server and by the page',()=>{
  assert.equal(REPOSITORY,'jasonswrench-beep/OpenStageGuide');
  assert.equal(RELEASES_PATH,`/${REPOSITORY}/releases`,'the page names the same repository');
  assert.equal(releasePage('v0.1.0-alpha.1'),`${RELEASES_PAGE}/tag/v0.1.0-alpha.1`);
  assert.equal(releasePage('latest'),RELEASES_PAGE);
  assert.equal(releasePage('0.1.0'),RELEASES_PAGE);
  for(const ok of [RELEASES_PAGE,`${RELEASES_PAGE}/tag/v0.1.0`,`${RELEASES_PAGE}/tag/v0.1.0-rc.2`]){assert.equal(isReleasePage(ok),true,ok);assert.equal(releaseLink(ok),ok,`the page agrees: ${ok}`);}
  for(const bad of [`${RELEASES_PAGE}/tag/latest`,`${RELEASES_PAGE}/tag/v`,`${RELEASES_PAGE}/tag/v0.1.0?x=1`,`${RELEASES_PAGE}/tag/v0.1.0#a`,`${RELEASES_PAGE}/download/v0.1.0/setup.exe`,`${RELEASES_PAGE}/tag/v0.1.0/../../x`,
    'https://github.com/someone/OpenStageGuide/releases','http://github.com/jasonswrench-beep/OpenStageGuide/releases','https://github.com:8443/jasonswrench-beep/OpenStageGuide/releases',
    'https://user@github.com/jasonswrench-beep/OpenStageGuide/releases','https://github.com.example/jasonswrench-beep/OpenStageGuide/releases','javascript:alert(1)','','not a url',undefined,null,5]){
    assert.equal(isReleasePage(String(bad)),false,String(bad));
    assert.equal(releaseLink(bad),null,`the page refuses: ${bad}`);
  }
});

test('what the dialog says: the server\'s words, a link only when there is one worth offering, and nothing offered when checking is off',()=>{
  const newer=describeAnswer({state:'newer',message:'OpenStage Guide 0.1.0 alpha 2 is out.',latest:'0.1.0 alpha 2',url:`${RELEASES_PAGE}/tag/v0.1.0-alpha.2`});
  assert.deepEqual(newer,{words:'OpenStage Guide 0.1.0 alpha 2 is out.',link:`${RELEASES_PAGE}/tag/v0.1.0-alpha.2`,linkLabel:'Open the page of OpenStage Guide 0.1.0 alpha 2'});
  assert.equal(describeAnswer({state:'current',message:'This copy is the newest release, 0.1.0.',url:RELEASES_PAGE}).linkLabel,'Open the releases page');
  assert.equal(describeAnswer({state:'off',message:'Checking for updates is turned off on this computer.',url:RELEASES_PAGE}).link,null,'off offers nothing');
  assert.equal(describeAnswer({state:'error',message:'x',url:'https://evil.example/'}).link,null,'an address that is not a release page is not offered');
  assert.match(describeAnswer(null).words,/could not be read/);
  assert.match(describeAnswer({state:'error',message:''}).words,/could not be read/);
});

// ---- the request

function fakeRequest({status=200,chunks=[],hang=false}={}){
  const seen={};
  const request=(url,options,callback)=>{
    seen.url=url;seen.options=options;
    const req=new EventEmitter();
    req.destroyed=false;
    req.destroy=error=>{req.destroyed=true;if(error)setImmediate(()=>req.emit('error',error));};
    req.end=()=>{
      seen.ended=true;
      if(hang){setImmediate(()=>req.emit('timeout'));return;}
      const res=new EventEmitter();
      res.statusCode=status;res.resume=()=>{seen.resumed=true;};
      callback(res);
      setImmediate(()=>{for(const c of chunks){if(req.destroyed)return;res.emit('data',Buffer.from(c));}if(!req.destroyed)res.emit('end');});
    };
    return req;
  };
  return {request,seen};
}

test('the one request: GitHub\'s list of this repository\'s releases, a plain user agent, nothing about the person, a time limit',async()=>{
  const {request,seen}=fakeRequest({chunks:['[{"tag_name":','"v0.1.0"}]']});
  const text=await fetchReleases(undefined,request);
  assert.equal(text,'[{"tag_name":"v0.1.0"}]');
  assert.equal(seen.url,RELEASES_API);
  assert.equal(seen.url,`https://api.github.com/repos/${REPOSITORY}/releases?per_page=30`);
  assert.equal(seen.options.method,'GET');
  assert.equal(seen.options.timeout,TIMEOUT_MS);assert.equal(TIMEOUT_MS,15000);
  assert.equal(seen.options.headers['User-Agent'],USER_AGENT);
  assert.equal(USER_AGENT,'OpenStage-Guide-update-check');
  assert.deepEqual(Object.keys(seen.options.headers).sort(),['Accept','User-Agent','X-GitHub-Api-Version'],'no cookie, no key, nothing else');
  assert.ok(!('body' in seen.options)&&!('agent' in seen.options));
  assert.equal(seen.ended,true);
});

test('a redirect is not followed, an answer that is not 200 is an error, a missing repository (not public yet) is an empty list, and a larger answer than allowed is cut off',async()=>{
  await assert.rejects(fetchReleases(undefined,fakeRequest({status:301}).request),/redirect, which is not followed/);
  await assert.rejects(fetchReleases(undefined,fakeRequest({status:403}).request),/GitHub answered 403/);
  await assert.rejects(fetchReleases(undefined,fakeRequest({status:500}).request),/GitHub answered 500/);
  assert.equal(await fetchReleases(undefined,fakeRequest({status:404}).request),'[]','404: no release found, not a failure');
  const big=fakeRequest({chunks:['x'.repeat(MAX_ANSWER_BYTES/2),'x'.repeat(MAX_ANSWER_BYTES/2),'x']});
  await assert.rejects(fetchReleases(undefined,big.request),/larger than it should be/);
  await assert.rejects(fetchReleases(undefined,fakeRequest({hang:true}).request),/did not answer in time/);
});

test('the answer to the button never throws, says when checking is off without asking, and says what to do when GitHub cannot be reached',async()=>{
  let asked=0;
  const settings=over=>({current:'v0.1.0-alpha.1',enabled:true,fetchReleases:async()=>{asked++;return list('v0.1.0-alpha.2');},...over});
  const off=await checkForUpdates(settings({enabled:false}));
  assert.deepEqual(off,{state:'off',message:'Checking for updates is turned off on this computer.'});
  assert.equal(asked,0,'off asks nothing of the network');
  assert.equal((await checkForUpdates(settings())).state,'newer');
  assert.equal(asked,1);
  const down=await checkForUpdates(settings({fetchReleases:async()=>{throw new Error('ENOTFOUND api.github.com');}}));
  assert.equal(down.state,'error');assert.match(down.message,/could not reach GitHub/);assert.equal(down.url,RELEASES_PAGE);
  assert.doesNotMatch(down.message,/ENOTFOUND|api\.github/,'the failure\'s own words are not passed on');
  assert.equal((await checkForUpdates(settings({fetchReleases:async()=>'garbage'}))).state,'error');
});
