import test from 'node:test';
import assert from 'node:assert/strict';
import {
  ASSESSMENT_DEFINITION_VERSION, ASSESSMENT_RESPONSE_VERSION, ASSESSMENT_SCORE_VERSION, ASSESSMENT_LIMITS,
  validateAssessmentDefinition, publishAssessmentDefinition, createAssessmentResponse, scoreAssessment
} from '../src/assessment-scoring.mjs';

const clone = value => structuredClone(value);
function definition(min = 1, max = 5, count = 4, reversed = [2, 4]) {
  return {schemaVersion: 1, id: 'assessment-original', version: 1, title: 'Synthetic original assessment', evidenceStatus: 'not-established',
    items: Array.from({length: count}, (_, i) => ({id: `item-${i + 1}`, type: 'likert', prompt: `Authored item ${i + 1}`, reverse: reversed.includes(i + 1),
      options: Array.from({length: max - min + 1}, (_, oi) => ({id: `option-${oi + 1}`, label: `Authored anchor ${oi + 1}`, code: min + oi}))})),
    subscales: [{id: 'sum-authored', name: 'Authored sum', itemIds: Array.from({length: count}, (_, i) => `item-${i + 1}`), aggregation: 'sum'},
      {id: 'mean-authored', name: 'Authored mean', itemIds: Array.from({length: count}, (_, i) => `item-${i + 1}`), aggregation: 'mean'}]};
}
function answers(d, values) {
  return values.map((code, i) => code === null ? {itemId: d.items[i].id, state: 'missing'} : code === 'NA' ? {itemId: d.items[i].id, state: 'not-applicable'} :
    {itemId: d.items[i].id, state: 'answered', optionId: d.items[i].options.find(option => option.code === code)?.id || 'unknown-option'});
}
const score = (d, values) => scoreAssessment(d, createAssessmentResponse(d, answers(d, values)));
const assertFrozen = value => {
  if (value && typeof value === 'object') { assert.equal(Object.isFrozen(value), true); for (const v of Object.values(value)) assertFrozen(v); }
};

test('versioned local publishing is defensive immutable and preserves authored text, labels, presentation order and explicit direction', () => {
  assert.equal(ASSESSMENT_DEFINITION_VERSION, 1); assert.equal(ASSESSMENT_RESPONSE_VERSION, 1); assert.equal(ASSESSMENT_SCORE_VERSION, 1);
  assert.equal(Object.isFrozen(ASSESSMENT_LIMITS), true);
  const d = definition(); d.title = '  Exact authored title\r\n'; d.items[0].prompt = 'Texte \u00e9 / \u65e5\u672c\u8a9e / \u{1f600}';
  d.items[0].options.reverse(); const before = clone(d), descriptors = Object.getOwnPropertyDescriptors(d);
  const validated = validateAssessmentDefinition(d), published = publishAssessmentDefinition(d);
  assert.deepEqual(validated, published); assertFrozen(published); assert.deepEqual(d, before); assert.deepEqual(Object.getOwnPropertyDescriptors(d), descriptors);
  assert.equal(Object.isFrozen(d), false); assert.equal(Object.isFrozen(d.items[0]), false);
  assert.equal(published.title, before.title); assert.equal(published.items[0].prompt, 'Texte \u00e9 / \u65e5\u672c\u8a9e / \u{1f600}'); assert.deepEqual(published.items[0].options, before.items[0].options);
  d.items[0].options[0].label = 'Later edit'; assert.notEqual(d.items[0].options[0].label, published.items[0].options[0].label);
  assert.equal(published.subscales[0].missing.mode, 'complete-case');
});

test('required 1-5 reverse-keyed vector retains raw codes and gives only explicit subscale sum12 mean3', () => {
  const d = definition(), r = score(d, [1, 2, 5, 4]);
  assert.deepEqual(r.items.map(item => item.rawCode), [1, 2, 5, 4]);
  assert.deepEqual(r.items.map(item => item.keyedValue), [1, 4, 5, 2]);
  assert.deepEqual(r.subscales.map(x => x.value), [12, 3]); assert.equal(Object.hasOwn(r, 'overall'), false);
  for (const s of r.subscales) { assert.equal(s.status, 'complete'); assert.equal(s.answeredCount, 4); assert.equal(s.plannedCount, 4); assert.equal(s.partial, false); }
  assertFrozen(r); assert.equal(r.evidenceStatus, 'not-established'); assert.match(r.interpretation, /No norms/);
});

test('configurable zero, shifted positive and negative code ranges use L+U-x exactly once', () => {
  for (const [min, max, raw, reversed, keyed, total] of [[0, 4, [0, 1, 4], [2], [0, 3, 4], 7], [2, 6, [2, 3, 6], [1, 2, 3], [6, 5, 2], 13], [-3, 3, [-3, -1, 3], [2], [-3, 1, 3], 1]]) {
    const d = definition(min, max, raw.length, reversed), r = score(d, raw);
    assert.deepEqual(r.items.map(x => x.keyedValue), keyed); assert.deepEqual(r.response.responses.map(x => x.rawCode), raw); assert.equal(r.subscales[0].value, total);
    assert.deepEqual(scoreAssessment(d, r.response), r, 'repeat scoring must not reverse twice');
  }
});

test('semantic-differential endpoints/high-score meaning are explicit and no right-is-good inference replaces reverse keying', () => {
  const d = definition(1, 7, 2, [2]);
  for (const [i, endpoints] of [[0, ['Unclear', 'Clear']], [1, ['Clear', 'Unclear']]]) Object.assign(d.items[i], {type: 'semantic-differential', target: 'The synthetic explanation', leftEndpoint: endpoints[0], rightEndpoint: endpoints[1], highScoreMeaning: 'Greater perceived clarity'});
  const r = score(d, [6, 2]); assert.deepEqual(r.items.map(x => x.keyedValue), [6, 6]); assert.deepEqual(r.items.map(x => x.rawCode), [6, 2]);
  const direct = clone(d); direct.items[1].reverse = false; assert.equal(score(direct, [6, 2]).items[1].keyedValue, 2);
  for (const key of ['target', 'leftEndpoint', 'rightEndpoint', 'highScoreMeaning']) { const bad = clone(d); delete bad.items[0][key]; assert.throws(() => validateAssessmentDefinition(bad), /Unsupported assessment/); }
  const same = clone(d); same.items[0].rightEndpoint = ' unclear '; assert.throws(() => validateAssessmentDefinition(same), /distinct/);
});

test('complete-case defaults keep missing and not-applicable distinct, report n/N and return no aggregate', () => {
  const d = definition(), r = score(d, [1, 'NA', 5, 4]);
  assert.deepEqual(r.items.map(x => x.state), ['answered', 'not-applicable', 'answered', 'answered']);
  assert.deepEqual(r.items.map(x => x.keyedValue), [1, null, 5, 2]);
  for (const s of r.subscales) { assert.equal(s.value, null); assert.equal(s.status, 'insufficient-responses'); assert.equal(s.answeredCount, 3); assert.equal(s.plannedCount, 4); assert.equal(s.notApplicableCount, 1); assert.equal(s.missingCount, 0); assert.equal(s.partial, false); }
  const missing = score(d, [1, null, 5, 4]); assert.equal(missing.subscales[0].missingCount, 1); assert.equal(missing.subscales[0].notApplicableCount, 0);
});

test('explicit mean minimumAnswered3 returns8/3 partial3/4, threshold refusal and all-missing no-score', () => {
  const d = definition(); d.subscales[1].missing = {mode: 'minimum-answered', minimumAnswered: 3};
  const r = score(d, [1, 'NA', 5, 4]), mean = r.subscales[1];
  assert.equal(r.subscales[0].value, null); assert.equal(mean.value, 8 / 3); assert.equal(mean.status, 'partial'); assert.equal(mean.partial, true); assert.equal(mean.answeredCount, 3); assert.equal(mean.plannedCount, 4);
  assert.equal(score(d, [1, 'NA', null, 4]).subscales[1].value, null);
  const all = scoreAssessment(d, createAssessmentResponse(d, []));
  assert.ok(all.items.every(x => x.state === 'missing' && x.rawCode === null && x.keyedValue === null));
  assert.ok(all.subscales.every(x => x.status === 'no-responses' && x.value === null && x.partial === false));
  assert.equal(score(d, [1, 2, 5, 4]).subscales[1].status, 'complete');
});

test('response constructor fills omitted rows with missing, keeps explicit NA and maps stable IDs independently of display position', () => {
  const d = definition(); d.items[0].options.reverse();
  const entries = [{itemId: 'item-3', state: 'not-applicable'}, {itemId: 'item-1', state: 'answered', optionId: 'option-1'}], before = clone(entries);
  const r = createAssessmentResponse(d, entries); assert.deepEqual(entries, before); assert.equal(Object.isFrozen(entries[0]), false);
  assert.deepEqual(r.responses, [{itemId: 'item-1', state: 'answered', optionId: 'option-1', rawCode: 1}, {itemId: 'item-2', state: 'missing'}, {itemId: 'item-3', state: 'not-applicable'}, {itemId: 'item-4', state: 'missing'}]);
  assert.equal(scoreAssessment(d, r).items[0].rawCode, 1); assertFrozen(r);
});

test('per-item-only definitions and explicit overlapping named subscales have no inferred total or duplicate membership', () => {
  const d = definition(); d.subscales = [];
  const r = score(d, [1, 2, 5, 4]); assert.deepEqual(r.subscales, []); assert.equal(r.items.length, 4); assert.ok(r.items.every(x => x.subscaleIds.length === 0));
  const overlap = definition(); overlap.subscales.push({id: 'overlap', name: 'Explicit overlap', itemIds: ['item-1', 'item-2'], aggregation: 'mean'});
  assert.deepEqual(score(overlap, [1, 2, 5, 4]).items[0].subscaleIds, ['sum-authored', 'mean-authored', 'overlap']);
  for (const mutate of [d => d.subscales[0].itemIds.push('item-1'), d => d.subscales[0].itemIds.push('unknown'), d => d.subscales[0].itemIds = [], d => d.subscales.push(clone(d.subscales[0]))]) {
    const bad = definition(); mutate(bad); assert.throws(() => validateAssessmentDefinition(bad));
  }
});

test('mixed-range aggregation and partial sum are refused rather than silently rescaled or prorated', () => {
  const d = definition(); d.items[0].options.forEach(x => x.code++); assert.throws(() => validateAssessmentDefinition(d), /mixed numeric ranges/);
  d.subscales = []; assert.doesNotThrow(() => validateAssessmentDefinition(d));
  for (const policy of [{mode: 'minimum-answered', minimumAnswered: 3}, {mode: 'prorate'}, {mode: 'complete-case', minimumAnswered: 4}, null, undefined]) {
    const bad = definition(); bad.subscales[0].missing = policy; assert.throws(() => validateAssessmentDefinition(bad));
  }
  for (const k of [0, -0, -1, 5, 1.5, '3', null, undefined]) { const bad = definition(); bad.subscales[1].missing = {mode: 'minimum-answered', minimumAnswered: k}; assert.throws(() => validateAssessmentDefinition(bad)); }
});

test('response snapshots bind all definition fields and exact roster, not merely assessment ID/version', () => {
  const d = definition(), snapshot = createAssessmentResponse(d, answers(d, [1, 2, 5, 4]));
  for (const mutate of [d => d.id = 'other', d => d.version++, d => d.title += ' changed', d => d.items[0].prompt += ' changed', d => d.items[0].reverse = true,
    d => d.items[0].options.reverse(), d => d.items[0].options[0].label += ' changed', d => d.items[0].options[0].id = 'changed', d => d.subscales[0].name += ' changed', d => d.subscales[0].itemIds.reverse(), d => d.items.reverse()]) {
    const newer = clone(d); mutate(newer); assert.throws(() => scoreAssessment(newer, snapshot), /match|changed/);
  }
  for (const mutate of [r => r.definitionVersion++, r => r.assessmentId = 'other', r => r.schemaVersion++, r => r.responses.reverse(), r => r.responses.pop(), r => r.responses.push(clone(r.responses[0])), r => r.responses[0].rawCode = 999, r => r.responses[0].rawCode = '1', r => r.responses[0].optionId = 'unknown-option']) {
    const bad = clone(snapshot); mutate(bad); assert.throws(() => scoreAssessment(d, bad));
  }
});

test('duplicate, unknown and numeric-only response entries refuse with original raw data retained', () => {
  const d = definition();
  for (const entries of [[{itemId: 'item-1', state: 'answered', optionId: 'option-1'}, {itemId: 'item-1', state: 'missing'}], [{itemId: 'unknown', state: 'missing'}],
    [{itemId: 'item-1', state: 'answered', optionId: 'unknown'}], [{itemId: 'item-1', state: 'answered', code: 9}], [{itemId: 'item-1', state: 'missing', optionId: 'option-1'}],
    [{itemId: 'item-1', state: 'NA'}], [{itemId: 'item-1', state: 'answered', optionId: 'option-1', rawCode: 1}]]) {
    const before = clone(entries); assert.throws(() => createAssessmentResponse(d, entries)); assert.deepEqual(entries, before);
  }
});

test('strict schemas refuse null, unsupported versions/types/status, empty items, implicit fields and unsupported metadata', () => {
  for (const input of [null, undefined, [], 'text', 0, false, {}]) assert.throws(() => validateAssessmentDefinition(input), /Unsupported assessment/);
  for (const mutate of [d => d.schemaVersion = 2, d => d.version = 0, d => d.version = '1', d => d.version = -0, d => d.evidenceStatus = 'validated', d => d.items = [],
    d => d.items[0].type = 'free-text', d => delete d.items[0].reverse, d => d.items[0].reverse = 0, d => d.clinicalCutoff = 3, d => d.items[0].target = 'unused',
    d => d.items[0].id = d.items[1].id, d => d.subscales[0].aggregation = 'median']) { const bad = definition(); mutate(bad); assert.throws(() => validateAssessmentDefinition(bad)); }
  assert.throws(() => createAssessmentResponse(definition()), /argument count/); assert.throws(() => publishAssessmentDefinition(definition(), {}), /argument count/);
});

test('option/code boundaries refuse duplicate/nonconsecutive/unsafe/fractional/non-numeric/-0;2 and11 options valid', () => {
  for (const codes of [[1], Array.from({length: 12}, (_, i) => i), [1, 1], [1, 3], [0, -0], [1, '2'], [1, 2.5], [Infinity, 2], [NaN, 2], [Number.MAX_SAFE_INTEGER, 2], [-1000001, -1000000], [1000000, 1000001]]) {
    const d = definition(); d.items[0].options = codes.map((code, i) => ({id: `option-${i}`, label: `Label ${i}`, code})); assert.throws(() => validateAssessmentDefinition(d));
  }
  for (const count of [2, 11]) assert.doesNotThrow(() => validateAssessmentDefinition(definition(0, count - 1, 2, [])));
  const duplicate = definition(); duplicate.items[0].options[1].id = duplicate.items[0].options[0].id; assert.throws(() => validateAssessmentDefinition(duplicate));
});

test('getters at all nested record seams never execute and inherited/symbol/nonenumerable data refuses', () => {
  let getterCalls = 0;
  const getter = {get() { getterCalls++; throw new Error('getter executed'); }, enumerable: true};
  for (const build of [() => Object.defineProperty(definition(), 'title', getter), () => { const d = definition(); Object.defineProperty(d.items[0], 'type', getter); return d; },
    () => { const d = definition(); Object.defineProperty(d.items[0].options[0], 'code', getter); return d; }, () => { const d = definition(); d.subscales[0].missing = Object.defineProperty({}, 'mode', getter); return d; }]) assert.throws(() => validateAssessmentDefinition(build()), /own data/);
  const d = definition(), response = createAssessmentResponse(d, []), bad = clone(response); Object.defineProperty(bad, 'definition', getter); assert.throws(() => scoreAssessment(d, bad), /own data/);
  const entries = [{itemId: 'item-1', state: 'answered', optionId: 'option-1'}]; Object.defineProperty(entries[0], 'optionId', getter); assert.throws(() => createAssessmentResponse(d, entries), /own data/);
  assert.equal(getterCalls, 0);
  for (const build of [() => Object.assign(Object.create({extra: 1}), definition()), () => { const d = definition(); d[Symbol('extra')] = 'no'; return d; },
    () => Object.defineProperty(definition(), 'hidden', {value: 1}), () => Object.defineProperty(definition(), 'title', {value: 'Title', enumerable: false})]) assert.throws(() => validateAssessmentDefinition(build()));
});

test('dense-array validation refuses sparse/accessor/custom prototype/custom iterator arrays without invoking them', () => {
  let calls = 0;
  const malformed = [new Array(2), Object.assign([definition().items[0]], {extra: true}), Object.setPrototypeOf([definition().items[0]], null),
    Object.defineProperty([definition().items[0]], '0', {get() { calls++; return definition().items[0]; }, enumerable: true}),
    Object.assign([definition().items[0]], {[Symbol.iterator]: function* () { calls++; yield definition().items[0]; }})];
  for (const items of malformed) { const d = definition(); d.items = items; assert.throws(() => validateAssessmentDefinition(d)); }
  const d = definition(); const entries = Object.defineProperty([], '0', {get() { calls++; return {itemId: 'item-1', state: 'missing'}; }, enumerable: true}); assert.throws(() => createAssessmentResponse(d, entries));
  assert.equal(calls, 0);
});

test('bounded text/node/version limits and XML-safe Unicode refuse without mutating or freezing callers', () => {
  for (const text of ['', '   ', '\u0000', '\ud800', '\udfff', 'x'.repeat(16001)]) { const d = definition(); d.title = text; const before = clone(d); assert.throws(() => validateAssessmentDefinition(d)); assert.deepEqual(d, before); assert.equal(Object.isFrozen(d), false); }
  const many = definition(1, 2, 201, []); assert.throws(() => validateAssessmentDefinition(many), /length/);
  const manyScales = definition(); manyScales.subscales = Array.from({length: 51}, (_, i) => ({id: `scale-${i}`, name: 'Scale', itemIds: ['item-1'], aggregation: 'mean'})); assert.throws(() => validateAssessmentDefinition(manyScales), /length/);
  const budget = definition(1, 2, 20, []); budget.items.forEach(item => item.prompt = 'x'.repeat(16000)); assert.throws(() => validateAssessmentDefinition(budget), /budget/);
  const max = definition(-1000000, -999999, 200, []); max.version = 1000000; max.subscales = []; assert.doesNotThrow(() => validateAssessmentDefinition(max));
  const badVersion = definition(); badVersion.version = 1000001; assert.throws(() => validateAssessmentDefinition(badVersion));
});

test('canonical mappings are deterministic and never traverse opaque extra data or mutate input descriptors on refusal', () => {
  const d = definition(), entries = answers(d, [1, 2, 5, 4]), beforeD = clone(d), beforeE = clone(entries);
  const a = createAssessmentResponse(d, entries), b = createAssessmentResponse(d, entries); assert.deepEqual(a, b); assert.deepEqual(scoreAssessment(d, a), scoreAssessment(d, b));
  assert.deepEqual(d, beforeD); assert.deepEqual(entries, beforeE); assert.equal(Object.isFrozen(entries), false);
  let reads = 0; const opaque = Object.defineProperty({}, 'danger', {get() { reads++; throw new Error('opaque'); }}), bad = clone(d); bad.extra = opaque;
  assert.throws(() => validateAssessmentDefinition(bad)); assert.equal(reads, 0); assert.equal(bad.extra, opaque); assert.equal(Object.isFrozen(opaque), false);
  const inherited = clone(d); Object.setPrototypeOf(inherited.items[0], {get reverse() { reads++; return false; }}); assert.throws(() => validateAssessmentDefinition(inherited)); assert.equal(reads, 0);
});

test('zero and negative valid scores remain numeric scores, while all-N/A and all-missing never become zero', () => {
  const zero = definition(0, 4, 3, []), r = score(zero, [0, 0, 0]);
  assert.deepEqual(r.subscales.map(s => s.value), [0, 0]); assert.ok(r.subscales.every(s => s.status === 'complete'));
  const negative = definition(-3, 3, 3, []), n = score(negative, [-3, -2, -1]);
  assert.deepEqual(n.subscales.map(s => s.value), [-6, -2]);
  for (const values of [['NA', 'NA', 'NA'], [null, null, null]]) {
    const none = score(zero, values); assert.ok(none.subscales.every(s => s.value === null && s.status === 'no-responses'));
  }
  for (const entries of [null, undefined, 0, false, [{itemId: 'item-1', state: null}], [{itemId: 'item-1', state: 0}]]) assert.throws(() => createAssessmentResponse(zero, entries));
});

test('published and captured response copies survive later caller edits without input freezing or rescoring old keyed values', () => {
  const source = definition(), published = publishAssessmentDefinition(source), entries = answers(source, [1, 2, 5, 4]);
  const snapshot = createAssessmentResponse(source, entries), first = scoreAssessment(published, snapshot);
  source.title = 'Changed'; source.items[0].options[0].code = 99; source.items[1].reverse = false;
  entries[0].optionId = 'option-5'; entries.push({itemId: 'unknown', state: 'missing'});
  assert.deepEqual(scoreAssessment(published, snapshot), first); assert.deepEqual(first.items.map(s => s.rawCode), [1, 2, 5, 4]);
  assert.equal(Object.isFrozen(source.items[0].options[0]), false); assert.equal(Object.isFrozen(entries), false);
  assert.throws(() => { snapshot.responses[0].rawCode = 5; }, TypeError);
});

test('response snapshot refuses derived keyed/effective data and malformed own-data array fields with zero getter calls', () => {
  const d = definition(), snapshot = createAssessmentResponse(d, answers(d, [1, 2, 5, 4]));
  for (const key of ['keyedValue', 'effectiveScore', 'clinicalResult']) { const bad = clone(snapshot); bad.responses[0][key] = 9; assert.throws(() => scoreAssessment(d, bad)); }
  let getterCalls = 0; const bad = clone(snapshot); Object.defineProperty(bad.responses[0], 'rawCode', {get() { getterCalls++; return 1; }, enumerable: true});
  assert.throws(() => scoreAssessment(d, bad), /own data/); assert.equal(getterCalls, 0);
  const missing = createAssessmentResponse(d, []), mutated = clone(missing); mutated.responses[0].rawCode = 0; assert.throws(() => scoreAssessment(d, mutated));
  const negativeZero = definition(0, 4); const raw = clone(createAssessmentResponse(negativeZero, answers(negativeZero, [0, 0, 0, 0]))); raw.responses[0].rawCode = -0; assert.throws(() => scoreAssessment(negativeZero, raw));
});
