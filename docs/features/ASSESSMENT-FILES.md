# PRIVATE assessment authoring files

This optional local-file feature saves authoring work only after an explicit Save action. Load prepares a candidate and requires a Keep/Replace choice before replacing current assessment work. It supplies no automatic storage, network transmission, Word-document insertion or course migration.

Both file kinds are **PRIVATE author material**. Even a definition-only file contains numeric codes, reverse keys, subscale memberships and scoring rules. It is not the blank respondent-facing DOCX. Test responses, learner identities, calculated scores, captured test state, review/consent acknowledgements, export preferences/instructions and download readiness are excluded. Import never restores authority to export or score: review and capture the current definition again. The session's captured ID/version ledger remains session-only, not a durable registry or evidence of authorship/measurement validity.

## Version 1 envelope

```json
{"format":"openstage-private-assessment","version":1,"kind":"definition","payload":{}}
```

`kind` is exactly `definition` or `draft`. The example payload is illustrative, not a valid empty definition. Unknown fields, versions, kinds or migrations are refused. Files are plaintext JSON, not encrypted or digitally signed; validation does not establish ownership, permissions, psychometric validity, norms or clinical interpretation.

- A **definition** payload uses the existing assessment definition schema1: stable IDs, authored version, title, original-authored `not-established` evidence status, ordered Likert-type/semantic-differential items, options and explicit scoring rules. The current scoring validator is always rerun. Saving a definition from the editor refuses unfinished active item/rule work instead of silently omitting it.
- A **draft** payload retains the raw title/version strings, saved items/rules, their saved editor buffers, active item/rule buffers and selected rows. Inactive semantic text and inactive missing-rule text are retained. Raw unfinished numeric/blank text remains raw, and cannot silently become a valid definition. Saved buffers must compile to their exact saved items/rules. All saved and active membership rosters contain exactly the known items, with explicit check states and authored ordering.

No browser validity flags, serialized dirty/baseline/serial claims, DOM objects, response state or registry are accepted. Clean/unfinished state is derived under current policy from exact saved buffers or a pristine new-item/rule default. An imported ID/version with different definition content does not authorize replacing a previously captured identity; the author must use a new authored version before validated save/review/capture.

## Bounds and text

- File input/output: at most **4 MiB UTF-8**. The parser checks size before copying/decoding and refuses malformed UTF-8, UTF-8 BOM, duplicate JSON object keys (including escaped aliases), trailing data, nesting deeper than16 levels and more than100,000 JSON value nodes.
- Draft text: at most **1,000,000 UTF-16 code units** across retained text, up to16,000 per ordinary field and80 per supported stable ID. Version/minimum/code raw strings allow12 characters, anchor-count strings2 and minimum-answered strings3. Compiled definitions retain their existing stricter250,000-unit budget.
- Existing model limits remain200 items,50 subscales and2–11 options per item. Numeric codes, reversal and missing rules use the existing scoring contract.
- Raw drafts preserve supported Unicode scalar strings exactly, including escaped control characters and line breaks. Unpaired UTF-16 surrogates are refused rather than being replaced during UTF-8 encoding. Compiled definitions still require XML-safe text. The pane must also prove that its actual input controls can represent imported values exactly before Replace; unsupported browser value sanitization refuses the candidate while preserving current work and the original file.

Canonical output is complete-string UTF-8 JSON with two-space indentation and a final LF. No title/path/filename is authority for the payload kind or version. Strict schema validation is not a cryptographic authenticity check, and imported originals are never rewritten automatically.

## Pure module API

`src/assessment-file.mjs` contains no storage, file picker, download, fetch or Word dependency. All APIs enforce exact arity and return defensive immutable data except serializers, which return fresh bytes.

- `validateAssessmentDraft(draft)` validates the version1 editable payload.
- `assessmentDraftEditorState(draft)` derives `{itemClean, ruleClean}` without trusting saved claims.
- `assessmentDraftIdentity(draft)` returns a validated saved-core identity or `null`; it deliberately ignores unfinished active editors and is for collision checks only, never export authorization.
- `assessmentDefinitionFromDraft(draft)` additionally refuses unfinished active work and validates the complete definition.
- `assessmentDraftFromDefinition(definition)` constructs current editable buffers while preserving item/option/subscale/member order. It restores no capture, response or consent state.
- `createAssessmentFile(kind, payload)` validates the exact PRIVATE envelope.
- `serializeAssessmentFile(file)` returns bounded canonical `Uint8Array` bytes.
- `parseAssessmentFile(bytes)` accepts `Uint8Array` bytes (including a Buffer), makes an independent snapshot, parses strict UTF-8/JSON and validates the envelope. It does not accept strings, arbitrary views or ArrayBuffers as an alternative implicit decoding route.

Keyboard/screen-reader behavior, native Word interaction and print/PDF qualification require their own supported tests. Portable workbook-design/logo bundles, automatic origin/document persistence, response saving, batch collection, cloud sync, delivery-pack encoding changes and in-place template transfer are separate scopes and are not implemented by this file format.
