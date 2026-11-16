// The editor's decisions apart from the page (program/ui/editor-model.js): which rows are drawn again, and what each key does.
import test from 'node:test';
import assert from 'node:assert/strict';
import {reconcile,rowSignature,sameList,wordsKey,rowKey,wordsLabel,textKey} from '../program/ui/editor-model.js';

const sig=list=>list.map(x=>rowSignature({style:'S',xml:x}));

test('only the rows that changed are drawn again: the same at both ends, different in the middle',()=>{
  assert.deepEqual(reconcile(sig(['a','b','c']),sig(['a','b','c'])),{prefix:3,oldEnd:3,newEnd:3},'nothing changed');
  assert.deepEqual(reconcile(sig(['a','b','c']),sig(['a','B','c'])),{prefix:1,oldEnd:2,newEnd:2},'one row changed');
  assert.deepEqual(reconcile(sig(['a','b','c']),sig(['a','b1','b2','c'])),{prefix:1,oldEnd:2,newEnd:3},'one row became two (Enter)');
  assert.deepEqual(reconcile(sig(['a','b1','b2','c']),sig(['a','b','c'])),{prefix:1,oldEnd:3,newEnd:2},'two became one (Backspace)');
  assert.deepEqual(reconcile(sig(['a','b','c']),sig(['a','c'])),{prefix:1,oldEnd:2,newEnd:1},'one removed');
  assert.deepEqual(reconcile(sig(['a','c']),sig(['a','b','c'])),{prefix:1,oldEnd:1,newEnd:2},'one added');
  assert.deepEqual(reconcile([],sig(['a','b'])),{prefix:0,oldEnd:0,newEnd:2},'a guide that was empty');
  assert.deepEqual(reconcile(sig(['a','b']),[]),{prefix:0,oldEnd:2,newEnd:0},'a guide that became empty');
  assert.deepEqual(reconcile(sig(['x','x','x']),sig(['x','x'])),{prefix:2,oldEnd:3,newEnd:2},'rows that look the same are not confused: the overlap is counted once');
});

test('rebuilding from the replaced ranges always gives the new list, for any two lists',()=>{
  let seed=11;const rand=n=>{seed=(seed*1103515245+12345)&0x7fffffff;return seed%n;};
  for(let i=0;i<500;i++){
    const a=Array.from({length:rand(8)},()=>String(rand(3))),b=Array.from({length:rand(8)},()=>String(rand(3)));
    const {prefix,oldEnd,newEnd}=reconcile(a,b);
    const rebuilt=[...a.slice(0,prefix),...b.slice(prefix,newEnd),...a.slice(oldEnd)];
    assert.deepEqual(rebuilt,b,`${a} -> ${b}`);
    assert.ok(prefix<=oldEnd&&prefix<=newEnd);
  }
});

test('a row\'s signature changes with its words and with its style; lists are compared by content',()=>{
  assert.notEqual(rowSignature({style:'A',xml:'x'}),rowSignature({style:'B',xml:'x'}));
  assert.notEqual(rowSignature({style:'A',xml:'x'}),rowSignature({style:'A',xml:'y'}));
  assert.equal(rowSignature({style:'A',xml:'x'}),rowSignature({style:'A',xml:'x'}));
  assert.equal(rowSignature({}),rowSignature({style:'',xml:''}));
  assert.equal(sameList(['a','b'],['a','b']),true);
  assert.equal(sameList(['a','b'],['a']),false);
  assert.equal(sameList(['a','b'],['a','c']),false);
});

const k=(key,more={})=>({key,start:0,end:0,length:10,...more});

test('keys in the words: Enter, Escape, Backspace at the start, the arrows at the ends, Undo and Redo',()=>{
  assert.deepEqual(wordsKey(k('Enter',{start:4,end:4})),{do:'split'});
  assert.deepEqual(wordsKey(k('Escape')),{do:'stop'});
  assert.deepEqual(wordsKey(k('Backspace')),{do:'join'});
  assert.equal(wordsKey(k('Backspace',{start:3,end:3})),null,'Backspace in the words is an ordinary Backspace');
  assert.equal(wordsKey(k('Backspace',{start:0,end:4})),null,'with words selected it deletes them');
  assert.deepEqual(wordsKey(k('ArrowUp')),{do:'previous',caret:'end'});
  assert.deepEqual(wordsKey(k('ArrowLeft')),{do:'previous',caret:'end'});
  assert.equal(wordsKey(k('ArrowUp',{start:3,end:3})),null);
  assert.deepEqual(wordsKey(k('ArrowDown',{start:10,end:10})),{do:'next',caret:'start'});
  assert.deepEqual(wordsKey(k('ArrowRight',{start:10,end:10})),{do:'next',caret:'start'});
  assert.equal(wordsKey(k('ArrowDown',{start:9,end:9})),null);
  assert.equal(wordsKey(k('ArrowDown',{start:0,end:10})),null,'a selection that reaches the end is not "at the end"');
  assert.deepEqual(wordsKey(k('z',{ctrl:true})),{do:'undo'});
  assert.deepEqual(wordsKey(k('Z',{ctrl:true,shift:true})),{do:'redo'});
  assert.deepEqual(wordsKey(k('y',{ctrl:true})),{do:'redo'});
  assert.deepEqual(wordsKey(k('z',{meta:true})),{do:'undo'});
  assert.equal(wordsKey(k('z',{ctrl:true,alt:true})),null);
});

test('keys that are not the editor\'s are left to the text box: Shift+Enter, Ctrl+Enter, Tab, letters, and anything during IME composition',()=>{
  for(const more of [{shift:true},{ctrl:true},{alt:true},{meta:true}])assert.equal(wordsKey(k('Enter',more)),null,JSON.stringify(more));
  for(const key of ['Tab','a','Delete','Home','End','PageDown','F2','ArrowUp'].slice(0,6))assert.equal(wordsKey(k(key,{start:3,end:3})),null,key);
  assert.equal(wordsKey(k('Enter',{composing:true})),null,'Enter that picks an IME candidate is not a new paragraph');
  assert.equal(wordsKey(k('Escape',{composing:true})),null);
  assert.equal(wordsKey(k('z',{ctrl:true,composing:true})),null);
});

test('keys on a row: arrows, Home and End, Enter and F2 to change the words, Undo and Redo',()=>{
  assert.equal(rowKey({key:'ArrowDown'}),'down');
  assert.equal(rowKey({key:'ArrowUp'}),'up');
  assert.equal(rowKey({key:'Home'}),'first');
  assert.equal(rowKey({key:'End'}),'last');
  assert.equal(rowKey({key:'Enter'}),'edit');
  assert.equal(rowKey({key:'F2'}),'edit');
  assert.equal(rowKey({key:'z',ctrl:true}),'undo');
  assert.equal(rowKey({key:'Z',ctrl:true,shift:true}),'redo');
  assert.equal(rowKey({key:'y',ctrl:true}),'redo');
  assert.equal(rowKey({key:'ArrowDown',shift:true}),'extendDown','Shift+Down chooses the next row too');
  assert.equal(rowKey({key:'ArrowUp',shift:true}),'extendUp');
  assert.equal(rowKey({key:'ArrowDown',alt:true}),'moveDown');
  assert.equal(rowKey({key:'ArrowUp',alt:true}),'moveUp');
  for(const more of [{ctrl:true},{meta:true}])assert.equal(rowKey({key:'ArrowDown',...more}),null,JSON.stringify(more));
  assert.equal(rowKey({key:'Home',shift:true}),null);
  assert.equal(rowKey({key:'Delete'}),'delete');
  assert.equal(rowKey({key:'Escape'}),'clear');
  assert.equal(rowKey({key:'c',ctrl:true}),'copy');
  assert.equal(rowKey({key:'x',ctrl:true}),'cut');
  assert.equal(rowKey({key:'a',ctrl:true}),'selectAll');
  assert.equal(rowKey({key:'Enter',ctrl:true}),'add');
  assert.equal(rowKey({key:'Enter',ctrl:true,shift:true}),'block','Ctrl+Shift+Enter adds a block');
  assert.equal(rowKey({key:'Enter',meta:true,shift:true}),'block');
  assert.equal(rowKey({key:'Enter',shift:true}),null,'Shift+Enter alone does nothing on a row');
  assert.equal(rowKey({key:'Delete',shift:true}),null);
  assert.equal(rowKey({key:'c',ctrl:true,shift:true}),null,'Ctrl+Shift+C is not Copy');
  assert.equal(rowKey({key:'a'}),null);
  assert.equal(rowKey({key:'Tab'}),null,'Tab is left to the browser: it leaves the guide');
});

test('every text box has a name a screen reader can say, and an unknown style still has one',()=>{
  for(const style of ['OSG Guide Title','OSG Module','OSG Lesson','OSG Cue','OSG Cue Facilitator','OSG Cue Body','OSG Facilitator Notes','heading 1','Normal','Something else'])assert.match(wordsLabel(style),/\S/);
  assert.equal(wordsLabel('OSG Module'),'Module title');
  assert.equal(wordsLabel('Normal'),'Paragraph words');
  assert.equal(wordsLabel('Something else'),'Paragraph words');
});

import {cellKey,cellWordsKey} from '../program/ui/editor-model.js';

test('keys on a cell of a table: arrows between cells, Home and End along a row, Ctrl+Home and Ctrl+End, Enter to change the words, Escape to leave, Ctrl+Enter and Ctrl+Delete for rows',()=>{
  assert.deepEqual(['ArrowLeft','ArrowRight','ArrowUp','ArrowDown','Home','End','Enter','F2','Escape'].map(key=>cellKey({key})),['left','right','up','down','rowStart','rowEnd','edit','edit','leave']);
  assert.equal(cellKey({key:'Home',ctrl:true}),'top');
  assert.equal(cellKey({key:'End',ctrl:true}),'bottom');
  assert.equal(cellKey({key:'Enter',ctrl:true}),'addRow');
  assert.equal(cellKey({key:'Delete',ctrl:true}),'deleteRow');
  assert.equal(cellKey({key:'z',ctrl:true}),'undo');
  assert.equal(cellKey({key:'Z',ctrl:true,shift:true}),'redo');
  assert.equal(cellKey({key:'Delete'}),null,'a plain Delete does not take a row away');
  assert.equal(cellKey({key:'ArrowDown',shift:true}),null);
  assert.equal(cellKey({key:'a'}),null);
});

test('keys in the words of a cell: Enter goes down, Tab and Shift+Tab to the next and last cell, the arrows only at the ends of the words, Escape stops',()=>{
  const k=(key,more={})=>({key,start:0,end:0,length:6,...more});
  assert.deepEqual(cellWordsKey(k('Enter')),{do:'down'});
  assert.deepEqual(cellWordsKey(k('Tab')),{do:'next'});
  assert.deepEqual(cellWordsKey(k('Tab',{shift:true})),{do:'previous'});
  assert.deepEqual(cellWordsKey(k('Escape')),{do:'stop'});
  assert.deepEqual(cellWordsKey(k('ArrowLeft')),{do:'left'});
  assert.deepEqual(cellWordsKey(k('ArrowUp')),{do:'up'});
  assert.equal(cellWordsKey(k('ArrowLeft',{start:2,end:2})),null,'inside the words the arrow moves the caret');
  assert.deepEqual(cellWordsKey(k('ArrowRight',{start:6,end:6})),{do:'right'});
  assert.deepEqual(cellWordsKey(k('ArrowDown',{start:6,end:6})),{do:'downCell'});
  assert.equal(cellWordsKey(k('ArrowDown',{start:5,end:5})),null);
  assert.deepEqual(cellWordsKey(k('z',{ctrl:true})),{do:'undo'});
  assert.equal(cellWordsKey(k('Enter',{shift:true})),null,'Shift+Enter is left to the box');
  assert.equal(cellWordsKey(k('Enter',{composing:true})),null,'an input method\'s Enter is not "go down"');
  assert.equal(cellWordsKey(k('Backspace')),null,'Backspace in a cell is an ordinary Backspace: it never joins cells');
});

test('Word\'s alignment keys work on a row and in the words: Ctrl+L left, Ctrl+E centered, Ctrl+R right, Ctrl+J justified',()=>{
  assert.deepEqual(['l','e','r','j'].map(key=>rowKey({key,ctrl:true})),['alignLeft','alignCenter','alignRight','alignJustify']);
  assert.equal(rowKey({key:'E',ctrl:true,shift:true}),null,'with Shift it is not an alignment');
  assert.equal(rowKey({key:'e'}),null);
  const k=(key,more={})=>({key,start:2,end:2,length:6,...more});
  assert.deepEqual(wordsKey(k('e',{ctrl:true})),{do:'align',value:'center'});
  assert.deepEqual(wordsKey(k('L',{meta:true})),{do:'align',value:'left'});
  assert.deepEqual(wordsKey(k('r',{ctrl:true})),{do:'align',value:'right'});
  assert.deepEqual(wordsKey(k('j',{ctrl:true})),{do:'align',value:'justify'});
  assert.equal(wordsKey(k('j',{ctrl:true,alt:true})),null);
  assert.equal(wordsKey(k('e',{ctrl:true,composing:true})),null);
});

test('Word\'s keys for the look of words: Ctrl+B, Ctrl+I, Ctrl+U, Ctrl+= lowered, Ctrl+Shift+= raised, Ctrl+Space for none, in the words and on a row',()=>{
  const k=(key,extra={})=>({key,ctrl:true,...extra});
  assert.equal(textKey(k('b')),'bold');
  assert.equal(textKey(k('B',{shift:true})),null,'Ctrl+Shift+B is not bold');
  assert.equal(textKey(k('i')),'italic');
  assert.equal(textKey(k('u')),'underline');
  assert.equal(textKey(k('=')),'subscript');
  assert.equal(textKey(k('=',{shift:true})),'superscript');
  assert.equal(textKey(k('+',{shift:true})),'superscript');
  assert.equal(textKey(k(' ')),'clear');
  assert.equal(textKey({key:'b'}),null,'a letter alone is a letter');
  assert.equal(textKey(k('b',{alt:true})),null);
  assert.equal(textKey({key:'b',meta:true}),'bold','Command on a Mac');
  const words={start:2,end:5,length:10};
  assert.deepEqual(wordsKey({...k('b'),...words}),{do:'text',setting:'bold'});
  assert.deepEqual(wordsKey({...k(' '),...words}),{do:'text',setting:'clear'});
  assert.equal(wordsKey({...k('b'),...words,composing:true}),null,'not in the middle of a composition');
  assert.equal(rowKey(k('b')),'bold');
  assert.equal(rowKey(k('=')),'subscript');
  assert.equal(rowKey(k(' ')),'clearText');
  assert.equal(rowKey(k('i')),'italic');
});
