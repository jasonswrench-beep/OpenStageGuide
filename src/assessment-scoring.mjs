// Pure, local contracts for original authored assessments. No Word, storage,
// response collection, network, norms or clinical interpretation is supplied.
export const ASSESSMENT_DEFINITION_VERSION = 1;
export const ASSESSMENT_RESPONSE_VERSION = 1;
export const ASSESSMENT_SCORE_VERSION = 1;
export const ASSESSMENT_LIMITS = Object.freeze({
  maxItems: 200, maxSubscales: 50, minOptions: 2, maxOptions: 11,
  maxCodeMagnitude: 1000000, maxDefinitionVersion: 1000000,
  maxTextUnits: 250000, maxTextFieldUnits: 16000, maxIdUnits: 80
});

const fail = message => { throw new Error(`Unsupported assessment: ${message}`); };
const arity = (actual, expected) => { if (actual !== expected) fail('argument count is unsupported.'); };
const record = (value, required, optional, label) => {
  if (!value || typeof value !== 'object' || Array.isArray(value) ||
      ![Object.prototype, null].includes(Object.getPrototypeOf(value))) fail(`${label} must be a plain own-data record.`);
  const keys = Reflect.ownKeys(value), allowed = [...required, ...optional], result = Object.create(null);
  if (keys.some(key => typeof key !== 'string' || !allowed.includes(key)) ||
      required.some(key => !keys.includes(key))) fail(`${label} fields are unsupported or missing.`);
  for (const key of keys) {
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor?.enumerable || !Object.hasOwn(descriptor, 'value')) fail(`${label}.${key} must be enumerable own data.`);
    result[key] = descriptor.value;
  }
  return result;
};
const list = (value, label, max, min = 0) => {
  if (!Array.isArray(value) || Object.getPrototypeOf(value) !== Array.prototype) fail(`${label} must be a dense array.`);
  const length = Object.getOwnPropertyDescriptor(value, 'length')?.value;
  if (!Number.isSafeInteger(length) || length < min || length > max) fail(`${label} length is unsupported.`);
  const keys = Reflect.ownKeys(value);
  if (keys.length !== length + 1 || !keys.includes('length') ||
      keys.some(key => key !== 'length' && (typeof key !== 'string' || !/^(0|[1-9][0-9]*)$/u.test(key) || Number(key) >= length))) fail(`${label} must contain only dense indexed data.`);
  const result = [];
  for (let i = 0; i < length; i++) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(i));
    if (!descriptor?.enumerable || !Object.hasOwn(descriptor, 'value')) fail(`${label}[${i}] must be enumerable own data.`);
    result.push(descriptor.value);
  }
  return result;
};
const supportedXmlText = value => {
  for (const symbol of value) {
    const cp = symbol.codePointAt(0);
    if (!(cp === 9 || cp === 10 || cp === 13 || cp >= 0x20 && cp <= 0xd7ff || cp >= 0xe000 && cp <= 0xfffd || cp >= 0x10000 && cp <= 0x10ffff)) return false;
  }
  return true;
};
const text = (value, label, budget, max = ASSESSMENT_LIMITS.maxTextFieldUnits) => {
  if (typeof value !== 'string' || !value.trim() || value.length > max || !supportedXmlText(value)) fail(`${label} must be nonblank supported text.`);
  budget.units += value.length;
  if (budget.units > ASSESSMENT_LIMITS.maxTextUnits) fail('text budget exceeded.');
  return value;
};
const id = (value, label, budget) => {
  if (typeof value !== 'string' || !/^[A-Za-z][A-Za-z0-9._:-]{0,79}$/u.test(value)) fail(`${label} must be a supported stable ID.`);
  return text(value, label, budget, ASSESSMENT_LIMITS.maxIdUnits);
};
const integer = (value, label, min, max) => {
  if (!Number.isSafeInteger(value) || Object.is(value, -0) || value < min || value > max) fail(`${label} must be a supported integer.`);
  return value;
};
const freezeCopy = value => {
  if (value && typeof value === 'object') {
    for (const child of Object.values(value)) freezeCopy(child);
    Object.freeze(value);
  }
  return value;
};
const unique = (value, set, label) => {
  if (set.has(value)) fail(`${label} is duplicated.`);
  set.add(value);
  return value;
};

function normalizeDefinition(value) {
  const source = record(value, ['schemaVersion','id','version','title','evidenceStatus','items','subscales'], [], 'definition');
  if (source.schemaVersion !== ASSESSMENT_DEFINITION_VERSION) fail('definition schema version is unavailable.');
  if (source.evidenceStatus !== 'not-established') fail('only original authored measures with validation evidence not established are supported.');
  const budget = {units: 0};
  const assessmentId = id(source.id, 'assessment ID', budget);
  const version = integer(source.version, 'definition version', 1, ASSESSMENT_LIMITS.maxDefinitionVersion);
  const title = text(source.title, 'assessment title', budget);
  const itemIds = new Set();
  const items = list(source.items, 'items', ASSESSMENT_LIMITS.maxItems, 1).map((value, index) => {
    const label = `item[${index}]`;
    // Inspect type only as an own data field before choosing the exact shape.
    const peek = record(value, ['id','type','prompt','reverse','options'], ['target','leftEndpoint','rightEndpoint','highScoreMeaning'], label);
    if (!['likert','semantic-differential'].includes(peek.type)) fail(`${label} type is unavailable.`);
    const semantic = peek.type === 'semantic-differential';
    const source = record(value, semantic ? ['id','type','prompt','reverse','options','target','leftEndpoint','rightEndpoint','highScoreMeaning'] : ['id','type','prompt','reverse','options'], [], label);
    const itemId = unique(id(source.id, `${label} ID`, budget), itemIds, 'item ID');
    if (typeof source.reverse !== 'boolean') fail(`${label}.reverse must be explicit boolean data.`);
    const optionIds = new Set(), codes = new Set();
    const options = list(source.options, `${label} options`, ASSESSMENT_LIMITS.maxOptions, ASSESSMENT_LIMITS.minOptions).map((value, oi) => {
      const option = record(value, ['id','label','code'], [], `${label} option[${oi}]`);
      return {
        id: unique(id(option.id, `${label} option ID`, budget), optionIds, 'option ID'),
        label: text(option.label, `${label} option label`, budget),
        code: unique(integer(option.code, `${label} option code`, -ASSESSMENT_LIMITS.maxCodeMagnitude, ASSESSMENT_LIMITS.maxCodeMagnitude), codes, 'option code')
      };
    });
    const sorted = [...codes].sort((a, b) => a - b);
    if (sorted.some((code, i) => code !== sorted[0] + i)) fail(`${label} option codes must be consecutive integers.`);
    const result = {id: itemId, type: source.type, prompt: text(source.prompt, `${label} prompt`, budget), reverse: source.reverse, options};
    if (semantic) {
      result.target = text(source.target, `${label} target`, budget);
      result.leftEndpoint = text(source.leftEndpoint, `${label} left endpoint`, budget);
      result.rightEndpoint = text(source.rightEndpoint, `${label} right endpoint`, budget);
      result.highScoreMeaning = text(source.highScoreMeaning, `${label} high-score meaning`, budget);
      if (result.leftEndpoint.trim().toLowerCase() === result.rightEndpoint.trim().toLowerCase()) fail(`${label} endpoints must be distinct.`);
    }
    return result;
  });
  const byItem = new Map(items.map(item => [item.id, item]));
  const subscaleIds = new Set();
  const subscales = list(source.subscales, 'subscales', ASSESSMENT_LIMITS.maxSubscales).map((value, index) => {
    const label = `subscale[${index}]`, source = record(value, ['id','name','itemIds','aggregation'], ['missing'], label);
    const subscaleId = unique(id(source.id, `${label} ID`, budget), subscaleIds, 'subscale ID');
    if (!['sum','mean'].includes(source.aggregation)) fail(`${label} aggregation is unavailable.`);
    const members = new Set();
    const ids = list(source.itemIds, `${label} item IDs`, ASSESSMENT_LIMITS.maxItems, 1).map(value => {
      const itemId = unique(id(value, `${label} item ID`, budget), members, 'subscale membership');
      if (!byItem.has(itemId)) fail(`${label} references an unknown item.`);
      return itemId;
    });
    const ranges = ids.map(itemId => {
      const codes = byItem.get(itemId).options.map(option => option.code);
      return `${Math.min(...codes)}:${Math.max(...codes)}`;
    });
    if (new Set(ranges).size !== 1) fail(`${label} cannot aggregate mixed numeric ranges in version1.`);
    let missing = {mode: 'complete-case'};
    if (Object.hasOwn(source, 'missing')) {
      const policy = record(source.missing, ['mode'], ['minimumAnswered'], `${label} missing policy`);
      if (policy.mode === 'complete-case') {
        record(source.missing, ['mode'], [], `${label} complete-case policy`);
      } else if (policy.mode === 'minimum-answered') {
        const exact = record(source.missing, ['mode','minimumAnswered'], [], `${label} minimum-answered policy`);
        if (source.aggregation !== 'mean') fail(`${label} sum requires complete-case responses.`);
        missing = {mode: 'minimum-answered', minimumAnswered: integer(exact.minimumAnswered, `${label} minimum answered`, 1, ids.length)};
      } else fail(`${label} missing policy is unavailable.`);
    }
    return {id: subscaleId, name: text(source.name, `${label} name`, budget), itemIds: ids, aggregation: source.aggregation, missing};
  });
  return {schemaVersion: ASSESSMENT_DEFINITION_VERSION, id: assessmentId, version, title, evidenceStatus: 'not-established', items, subscales};
}

/** Validate and return a defensive immutable canonical definition. */
export function validateAssessmentDefinition(value) {
  arity(arguments.length, 1);
  return freezeCopy(normalizeDefinition(value));
}

/** Publish a local immutable snapshot, not a global version registry or validity claim. */
export function publishAssessmentDefinition(value) {
  arity(arguments.length, 1);
  return freezeCopy(normalizeDefinition(value));
}

function normalizeResponses(value, definition, snapshot) {
  const byItem = new Map(definition.items.map(item => [item.id, item])), seen = new Set(), budget = {units: 0};
  const rows = list(value, 'responses', ASSESSMENT_LIMITS.maxItems);
  const responses = new Map();
  for (let i = 0; i < rows.length; i++) {
    const label = `response[${i}]`, peek = record(rows[i], ['itemId','state'], ['optionId','rawCode'], label);
    const itemId = unique(id(peek.itemId, `${label} item ID`, budget), seen, 'response item ID');
    const item = byItem.get(itemId);
    if (!item) fail(`${label} refers to an unknown item.`);
    if (peek.state === 'answered') {
      const source = record(rows[i], snapshot ? ['itemId','state','optionId','rawCode'] : ['itemId','state','optionId'], [], label);
      const optionId = id(source.optionId, `${label} option ID`, budget);
      const option = item.options.find(option => option.id === optionId);
      if (!option) fail(`${label} has an unknown option.`);
      if (snapshot && (Object.is(source.rawCode, -0) || source.rawCode !== option.code)) fail(`${label} raw code does not match its definition option.`);
      responses.set(itemId, {itemId, state: 'answered', optionId, rawCode: option.code});
    } else if (peek.state === 'missing' || peek.state === 'not-applicable') {
      record(rows[i], ['itemId','state'], [], label);
      responses.set(itemId, {itemId, state: peek.state});
    } else fail(`${label} state is unavailable.`);
  }
  if (snapshot && rows.length !== definition.items.length) fail('response snapshot must contain the exact item roster.');
  const result = definition.items.map(item => responses.get(item.id) || {itemId: item.id, state: 'missing'});
  if (snapshot && result.some((row, i) => row.itemId !== record(rows[i], ['itemId','state'], ['optionId','rawCode'], `response[${i}]`).itemId)) fail('response snapshot order changed.');
  return result;
}

/** Construct a single local response; omitted items become explicit missing rows. */
export function createAssessmentResponse(value, entries) {
  arity(arguments.length, 2);
  const definition = normalizeDefinition(value);
  const responses = normalizeResponses(entries, definition, false);
  return freezeCopy({schemaVersion: ASSESSMENT_RESPONSE_VERSION, assessmentId: definition.id, definitionVersion: definition.version, definition, responses});
}

function normalizeSnapshot(value, expectedDefinition) {
  const source = record(value, ['schemaVersion','assessmentId','definitionVersion','definition','responses'], [], 'response snapshot');
  if (source.schemaVersion !== ASSESSMENT_RESPONSE_VERSION || source.assessmentId !== expectedDefinition.id || source.definitionVersion !== expectedDefinition.version) fail('response snapshot assessment/version does not match.');
  const definition = normalizeDefinition(source.definition);
  if (JSON.stringify(definition) !== JSON.stringify(expectedDefinition)) fail('response snapshot definition changed, even if its ID/version is unchanged.');
  const responses = normalizeResponses(source.responses, definition, true);
  return {schemaVersion: ASSESSMENT_RESPONSE_VERSION, assessmentId: definition.id, definitionVersion: definition.version, definition, responses};
}

/** Deterministic raw/keyed item values and explicitly requested subscale scores. */
export function scoreAssessment(value, responseSnapshot) {
  arity(arguments.length, 2);
  const definition = normalizeDefinition(value), response = normalizeSnapshot(responseSnapshot, definition);
  const items = definition.items.map((item, i) => {
    const raw = response.responses[i], codes = item.options.map(option => option.code), minimumCode = Math.min(...codes), maximumCode = Math.max(...codes);
    const answered = raw.state === 'answered';
    return {itemId: item.id, type: item.type, state: raw.state, optionId: answered ? raw.optionId : null,
      rawCode: answered ? raw.rawCode : null, keyedValue: answered ? (item.reverse ? minimumCode + maximumCode - raw.rawCode : raw.rawCode) : null,
      reverse: item.reverse, minimumCode, maximumCode, subscaleIds: definition.subscales.filter(subscale => subscale.itemIds.includes(item.id)).map(subscale => subscale.id)};
  });
  const byItem = new Map(items.map(item => [item.itemId, item]));
  const subscales = definition.subscales.map(subscale => {
    const members = subscale.itemIds.map(itemId => byItem.get(itemId)), answered = members.filter(item => item.state === 'answered');
    const n = answered.length, N = members.length;
    const eligible = n > 0 && (subscale.missing.mode === 'complete-case' ? n === N : n >= subscale.missing.minimumAnswered);
    const total = answered.reduce((sum, item) => sum + item.keyedValue, 0);
    return {id: subscale.id, name: subscale.name, itemIds: [...subscale.itemIds], aggregation: subscale.aggregation,
      missing: {...subscale.missing}, answeredCount: n, plannedCount: N,
      missingCount: members.filter(item => item.state === 'missing').length, notApplicableCount: members.filter(item => item.state === 'not-applicable').length,
      status: eligible ? (n === N ? 'complete' : 'partial') : (n === 0 ? 'no-responses' : 'insufficient-responses'),
      partial: eligible && n < N, value: eligible ? (subscale.aggregation === 'sum' ? total : total / n) : null};
  });
  return freezeCopy({schemaVersion: ASSESSMENT_SCORE_VERSION, assessmentId: definition.id, definitionVersion: definition.version,
    evidenceStatus: 'not-established', interpretation: 'Raw scores only; validation evidence is not established. No norms, cutoffs or clinical interpretation are supplied.', response, items, subscales});
}
