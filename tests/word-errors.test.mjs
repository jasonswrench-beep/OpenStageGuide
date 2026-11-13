// Word's error codes are turned into plain instructions, so nobody is left with a bare "InvalidArgument".
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {explainWordError,isWordError} from '../src/word-errors.mjs';
import {createWordAdapter} from '../src/guide-word.mjs';

const wordError=(code,extra={})=>Object.assign(new Error(code),{code,...extra});

test('every code Word is known to give says what happened and what to do, with the code kept small at the end',()=>{
  const codes=['InvalidArgument','ItemNotFound','AccessDenied','InvalidSelection','GeneralException','InvalidOperation','InvalidObjectPath','NotImplemented','UnsupportedOperation','ApiNotFound','ActivityLimitReached','RequestPayloadSizeLimitExceeded','ResponsePayloadSizeLimitExceeded','RequestTimeout','Timeout'];
  for(const code of codes){
    const out=explainWordError(wordError(code));
    assert.notEqual(out.message,code,`${code} is no longer shown bare`);
    assert.ok(out.message.length>80,`${code} has a real explanation`);
    assert.match(out.message,/\b(Click|Choose|Wait|Update|Try|Save|Put)\b/,`${code} says what to do`);
    assert.ok(out.message.endsWith(`(Word's code: ${code}.)`),`${code} ends with its code`);
  }
});

test('InvalidArgument tells the author to click in the document, refresh and try again, and what to do if it repeats',()=>{
  const out=explainWordError(wordError('InvalidArgument',{debugInfo:{errorLocation:'InlinePicture.getRange'}}));
  assert.equal(out.message,"Word would not accept one of the changes the pane asked for. Click once in your document, choose Refresh in the pane, then try again. If it happens again, press Ctrl+Z to take back the last change and make the change by hand in Word. (Word's code: InvalidArgument, in InlinePicture.getRange.)");
});

test('a code that is not known still gets a plain message, and a missing code is said to be missing',()=>{
  const odd=explainWordError(wordError('SomethingNew'));
  assert.match(odd.message,/^Word reported a problem without saying what it was\. Click once in your document/);
  assert.match(odd.message,/\(Word's code: SomethingNew\.\)$/);
  const none=explainWordError(Object.assign(new Error('boom'),{name:'RichApiError'}));
  assert.match(none.message,/\(Word's code: none given\.\)$/);
});

test('the pane\'s own errors and Node\'s pass through untouched, and an explained error is not explained twice',()=>{
  const own=new Error('Type a title for the module first.'),node=Object.assign(new Error('x'),{code:'ENOENT'}),assertion=Object.assign(new Error('y'),{code:'ERR_ASSERTION'});
  for(const e of [own,node,assertion,null,undefined,'text'])assert.equal(explainWordError(e),e);
  assert.equal(isWordError(own),false);
  const once=explainWordError(wordError('ItemNotFound'));
  assert.equal(explainWordError(once),once);
  assert.equal(once.cause.code,'ItemNotFound','the original error is kept');
});

test('the Word adapter explains what Word refuses, whichever call it was',async()=>{
  const adapter=createWordAdapter({Word:{run:()=>Promise.reject(wordError('InvalidArgument'))},DOMParserClass:DOMParser});
  for(const call of [()=>adapter.paragraphs(),()=>adapter.title(),()=>adapter.cursorIndex(),()=>adapter.go(0),()=>adapter.sourcePackage()]){
    await assert.rejects(call(),e=>e.message!=='InvalidArgument'&&/Click once in your document, choose Refresh in the pane, then try again\./.test(e.message));
  }
  const plain=createWordAdapter({Word:{run:()=>Promise.reject(new Error('That block has changed since you chose it.'))},DOMParserClass:DOMParser});
  await assert.rejects(plain.paragraphs(),/^Error: That block has changed since you chose it\.$/);
});
