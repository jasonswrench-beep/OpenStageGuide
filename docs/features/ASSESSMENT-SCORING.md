# Pure assessment foundation v1 - API contract

This dependency-free module supplies validation, local immutable snapshots and deterministic raw/keyed scoring only. The separate assessment-panel controller consumes it for a session-only editor and single synthetic-response tester in the pane. The engine remains independent of the course schema, Word, exports, storage, response collection and a network. No new measure is validated by using it. Only original authored definitions with `evidenceStatus: 'not-established'` are accepted; established instruments and their specified manuals/permissions require a separate future contract.

## APIs

- `validateAssessmentDefinition(definition)` returns a defensive, recursively frozen canonical copy, or a controlled `Unsupported assessment:` Error. It does not mutate/freeze caller data.
- `publishAssessmentDefinition(definition)` performs the same local defensive validation/freeze for an immutable definition snapshot. This is **not a global version registry**, signature, or enforcement against two external callers choosing the same ID/version. An eventual authoring/persistence layer must manage published version identity. An old captured response cannot silently score against a changed same-ID/same-version definition because its full canonical definition is checked.
- `createAssessmentResponse(definition, entries)` captures one response as a defensive frozen full definition plus ordered raw response roster. Entry order is irrelevant, omitted items explicitly become missing, and every explicit entry is validated. Nothing is saved automatically.
- `scoreAssessment(definition, responseSnapshot)` revalidates both inputs, compares every canonical definition field and exact ordered response roster, then returns per-item raw/keyed state and explicitly requested subscale results. It never mutates an old keyed score or applies reverse-keying twice.

The three exported schema versions are 1. `ASSESSMENT_LIMITS` is immutable. APIs enforce their exact argument count; explicit null/undefined values are not treated as omitted records.

## Definition shape

```js
{
  schemaVersion: 1,
  id: 'assessment-unique-id',
  version: 1, // positive integer <= 1,000,000; authored content revision
  title: 'Exact authored title',
  evidenceStatus: 'not-established',
  items: [{
    id: 'stable-item-id',
    type: 'likert', // or 'semantic-differential'
    prompt: 'Exact authored prompt',
    reverse: false, // mandatory explicit boolean
    options: [
      {id: 'stable-option-a', label: 'Authored anchor A', code: 0},
      {id: 'stable-option-b', label: 'Authored anchor B', code: 1}
    ]
    // Semantic-differential items additionally require all four:
    // target, leftEndpoint, rightEndpoint, highScoreMeaning (nonblank text).
  }],
  subscales: [{
    id: 'stable-subscale-id', name: 'Authored group name',
    itemIds: ['stable-item-id'], aggregation: 'mean', // or 'sum'
    missing: {mode: 'complete-case'} // omission defaults to this exact policy
    // For mean only: {mode: 'minimum-answered', minimumAnswered: 1}
  }]
}
```

There must be 1-200 items, 2-11 options per item and 0-50 explicit subscales. Empty subscales means item scoring only, with no inferred aggregate/overall score. Item IDs are unique across items, subscale IDs across subscales, option IDs within their item; a response identifies both item and option. IDs use `[A-Za-z][A-Za-z0-9._:-]{0,79}`. Labels are separate from IDs, codes and array positions. Display order is preserved and need not be numeric order. Each item's distinct codes must form a consecutive integer set, bounded to +/-1,000,000; negative zero, nonfinite/fractional/unsafe integers and coercion are refused. These are implementation bounds, not measurement standards.

Subscales require nonempty explicit membership. Duplicate membership within a subscale refuses; explicitly authored overlap across different subscales is supported. Mixed minimum/maximum ranges within an aggregate refuse in v1. No rescaling, weighting, automatic overall score, norms, cutoff, reliability or validity inference is supplied.

Required text is preserved exactly, including supported whitespace, line breaks and Unicode; whitespace-only text and XML-unsafe controls/lone surrogates refuse. Each text field is <=16,000 UTF-16 units, total copied definition text <=250,000. Semantic endpoints must differ after trimming/case folding, but their substantive bipolarity/suitability is **not validated**. High-score meaning describes the author's intended keyed direction; the engine does not infer desirability from left/right placement.

## Raw response shapes and binding

Input entries are any authored subset, in any order:

```js
{itemId: 'stable-item-id', state: 'answered', optionId: 'stable-option-a'}
{itemId: 'other-item-id', state: 'missing'}
{itemId: 'third-item-id', state: 'not-applicable'}
```

Answered input accepts only an option ID, not a numeric-only answer or caller-derived keyed/effective score. Unknown IDs, duplicate responses and extra fields refuse. The constructor resolves and captures the selected option's `rawCode`; nonnumeric states carry no numeric/option fields. Returned response shape:

```js
{
  schemaVersion: 1, assessmentId: 'assessment-unique-id', definitionVersion: 1,
  definition: /* complete canonical immutable definition */,
  responses: [
    {itemId: 'stable-item-id', state: 'answered', optionId: 'stable-option-a', rawCode: 0}
    // missing/not-applicable rows have only itemId/state
  ]
}
```

Snapshot scoring requires the exact canonical definition (including labels/order/codes/direction/semantic fields/subscale rules), not merely its ID/version. Captured raw codes must match the captured selected options. Snapshots have no cryptographic authenticity guarantee; a fully reauthored self-consistent object is not evidence of who answered it. No person identifiers, timestamps or automatic persistence are added.

## Score behavior

Each item returns `itemId`, type, raw `state`, selected option ID/raw code (or null), explicit `reverse`, minimum/maximum code, `keyedValue` (or null), and explicitly authored subscale IDs. Reverse keying is `L + U - x` on the raw selected code exactly once. Zero and negative legitimate results remain numeric scores.

Each subscale returns the authored method/missing rule, member IDs, `answeredCount` n, `plannedCount` N, missing and N/A counts, status, partial flag and value. N/A remains in planned N. Complete-case requires n=N; sum is always complete-case. Optional mean requires an explicit `minimumAnswered` k in 1..N and averages only valid answered keyed values. A threshold-met incomplete mean is `partial` with n/N; insufficient or zero answered responses yield null value, never zero. Complete responses are `complete`. No rounding is applied; ordinary deterministic JavaScript arithmetic is used (safe bounded integer sums; mean may be fractional floating point).

## Validation and qualification limits

Only plain/null-prototype own-data records and dense standard arrays are supported. Extra, inherited, hidden, symbol, accessor, sparse/index-accessor, subclass or custom-iterator data refuse without invoking getters/custom iterators. As in ordinary JavaScript generally, validation does **not** promise to suppress arbitrary Proxy traps. Errors name controlled schema paths, not raw private response contents. Validation/freezing uses new records and arrays only.

Focused tests cover the supplied numerical vectors, all-zero/negative scores, all-missing/all-N/A, partial-mean thresholds, mixed ranges, explicit overlaps, snapshot freshness, caller preservation, strict schemas/getter-free refusals and budget boundaries. The separate session-only editor/tester has its own source/mock-DOM tests and does not supply native keyboard/screen-reader qualification, blank Word assessment/scoring-guide generation, published-version persistence or course migration. A pane close/reload or a different Word-document pane may lose session work; no persistence is implied.
