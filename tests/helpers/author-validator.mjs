// OpenStage Author's own validator (its tools/validate.py), run on a project folder. It lives in another repository: the tests that use it run only when OPENSTAGE_AUTHOR_VALIDATOR
// names that file, or when Author is in its usual place on the author's PC.
//
// Author's working copy there moves ahead of what Author has published while Author is being worked on (its schema version goes up before it is released), and its validator then says
// that a project "was made with an older version and needs updating" (rule needs-migration) and checks nothing else. That is not a fault in what Guide writes: Guide writes the
// version Author last published, and Author's editor offers to update such a project when it is opened. `staleOnly` says that is the only thing the validator found, so a test can
// say so and skip, and still fail for any other finding.
import fs from 'node:fs';
import {spawnSync} from 'node:child_process';
import path from 'node:path';

export const VALIDATOR=process.env.OPENSTAGE_AUTHOR_VALIDATOR||['C:/Dev/openstage-author/tools/validate.py'].find(f=>fs.existsSync(f))||'';
export const SKIP_WITHOUT=!VALIDATOR&&'set OPENSTAGE_AUTHOR_VALIDATOR to the OpenStage Author validate.py to run this';

/** Run the validator on a project folder. @returns {{status:number|null,stdout:string,stderr:string,staleOnly:boolean,newerVersion:string}} */
export function validateProject(projectDir){
  const run=spawnSync('python',['-B',VALIDATOR,'--json',projectDir],{encoding:'utf8',cwd:path.dirname(path.dirname(VALIDATOR))});
  let findings=[];
  try{findings=JSON.parse(run.stdout);}catch{/* not JSON (an older validator): the text is what there is */}
  const errors=Array.isArray(findings)?findings.filter(f=>String(f.severity||f.level||'').toLowerCase()==='error'):[];
  const staleOnly=run.status===1&&errors.length>0&&errors.every(f=>f.rule==='needs-migration');
  const newer=staleOnly?(/this version uses (\d+\.\d+\.\d+)/.exec(JSON.stringify(errors))||[])[1]||'a newer version':'';
  return {status:run.status,stdout:run.stdout,stderr:run.stderr,staleOnly,newerVersion:newer};
}
