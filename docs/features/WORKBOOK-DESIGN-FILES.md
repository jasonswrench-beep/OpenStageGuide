# PRIVATE workbook-design authoring files

Explicit local Save/Load recovers authored workbook design drafts and an optional verified PNG. It is not automatic browser/document storage, a saved course/master, an assessment file, a generated workbook or an in-place template transfer. A bundle may contain unpublished cover/research/rights text and an organization logo: keep it PRIVATE. It is plaintext JSON, not encryption, a signature or proof of ownership.

## Version 1 data

The exact envelope is `{format:"openstage-workbook-design",version:1,payload:...}`. `payload` has exactly:

- `familyId` and `baseLanguage`: a current listed family/language ID. No recipe version, font installation, old effective tokens or prior accessibility assertion is restored. New output uses current policy and requires a fresh review.
- `brand`: the sixteen authored raw brand override strings, including empty/default choices and incomplete numeric/color text. Browser validity flags, presets/catalog names and resolved values are absent.
- `frontMatter`: `includeCover`, `includeCopyright`, the six raw cover fields and nine raw copyright fields. No rights holder, license or permissions are inferred.
- `contents`: the authored `include` choice only. Heading lists, bookmarks, page numbers and course-derived content are regenerated for the current source; they are not stored here.
- `research`: `include` and raw rationale/evidence/application/limitations/references. These remain authored, unverified claims/citations.
- `logo`: raw `include`, `role`, `alt` and `widthIn`, plus `image:null` or `{mediaType:"image/png",normalizerVersion:1,sha256,base64}`. The image has no stored path, filename, asset handle/ID or trusted dimensions.

No course/project/plan IDs or source content, assessment/scoring material, tester responses, generated package, review/consent, generation stamp or ready-download state is accepted. Current full/selected scope remains a separate source choice. These are raw drafts: incomplete colors/numbers/year/notice text, inactive fields, decorative role or empty descriptions may be retained without becoming export-eligible. The current resolver and output validators still decide whether a newly reviewed workbook can be generated.

## PNG and source integrity

The Save factory accepts only an owned verified current `workbook-logo` handle or explicit `null`. It privately reads that handle's normalized bytes, including when logo inclusion is off. The pane must refuse Save while logo selection is loading, pending or invalid; it must not substitute a hidden last-valid asset. Empty-image drafts remain possible and cannot silently become a valid included logo.

Load validates bounded canonical base64 and the SHA-256 of the **stored bytes**, then runs the current `prepareWorkbookLogo` parser, decompressor, normalization and readback again. A new owned session handle is created. Stored hashes and `normalizerVersion` are integrity/profile metadata, not validation authority, authorship or appearance certification. JPEG, SVG, data URLs, paths, animation and unsupported PNG profiles remain unsupported.

Current normalization can produce different bytes, for example when safe metadata is removed or compression differs. The prepared result reports old/current hashes and `bytesChanged`; the pane previews the current normalized image before explicit replacement. It does not silently rewrite the imported file, demand equality of old/new compressor bytes or certify visual appearance/alt quality. Input/object/byte snapshots and a final asynchronous source check protect the prepared candidate.

## Bounds and parsing

- At most **6 MiB UTF-8** per file; malformed UTF-8, BOM, duplicate decoded JSON keys, trailing data, unknown fields/versions, more than16 nesting levels or100,000 JSON value nodes refuse.
- Raw text has a200,000 UTF-16-unit aggregate backstop plus the current controls' tighter field caps. Organization is120; colors7; numeric brand strings12; logo width32 and alt1000; cover/copyright and research use their current individual maxima, including16,000 for long text. The fixed current schema can reach a lower total than the aggregate backstop.
- Raw Unicode scalar strings, including escaped controls/line breaks in unfinished drafts, remain exact. Unpaired surrogates refuse instead of being replaced during UTF-8 encoding. The pane must separately prove that its controls can represent imported values exactly before Replace; unsupported browser sanitization refuses without changing current work.
- A PNG is at most2 MiB encoded,2048 pixels per side and4,194,304 pixels under the current normalizer. Canonical ASCII base64 length/decoded-byte bounds and padding bits are checked before allocation. Safe encoded text is not proof that its decoded bytes are a valid image.

All files and result data are defensive immutable copies; serialized bytes are fresh. The original file is not rewritten or migrated automatically. Keep, cancellation, malformed/oversized input and stale reads must preserve current drafts and source work. Imported drafts/assets require explicit replacement and a new output review; consent is never restored.

## Module API

`src/workbook-design-file.mjs` has no picker, storage, download, course read, network or Word dependency. Each function enforces exact arity:

- `validateWorkbookDesignDraft(draft)` validates the exact raw design payload **without** a file image field.
- `createWorkbookDesignFile(draft, logoAssetOrNull)` captures a verified current handle into the exact version1 file, or stores an absent image. A copied/forged handle refuses.
- `serializeWorkbookDesignFile(file)` returns canonical UTF-8 `Uint8Array` JSON with two-space indentation and final LF. This is format encoding, not adoption or export approval of an arbitrary image.
- `parseWorkbookDesignFile(bytes)` snapshots bounded `Uint8Array` input (including Buffer), checks strict UTF-8/JSON/schema and returns an immutable file. It does not adopt an image.
- `prepareWorkbookDesignFile(file)` asynchronously checks stored image integrity, applies the current PNG normalizer and returns `{file,draft,logoAsset,logoInfo,normalization}`. `normalization` is `null` without an image or `{storedSha256,currentSha256,bytesChanged}` with one. Adoption requires this prepared result and the pane's final source/visibility/replacement guard.

Manual design files do not add automatic origin/document persistence, a reusable image catalog, durable assessment/response storage, arbitrary image formats, delivery-pack changes or in-place transfer. Native file-picker/download, keyboard/screen-reader, Word layout/Accessibility Checker and print/PDF qualification remain separate.
