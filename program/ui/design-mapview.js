/**
 * The design desk's Course map: Guide's learning alignment, extended. The assessments are planned here (what checks which objective, at what level, with how many item slots
 * for OpenAssess), and two maps are drawn: the plan's, from the notes of the desk, and the guide's, from what is in the guide now. Either goes to a syllabus or an
 * accreditation report as a table in the guide or a CSV file. A gap is said in words, and what to do about it is the author's.
 */
import { el } from "./menubar.js";
import { field, choice, tickGroup, button, row, note, table, give } from "./design-ui.js";
import { planMap, guideMap, mapSummaryText, mapTable, mapCsv } from "../../src/design-map.mjs";
import { itemSlots, itemSlotsCsv } from "../../src/design-storyboard.mjs";
import { ASSESSMENT_KINDS, nextId, removeAssessment } from "../../src/design-model.mjs";
import { BLOOM_LEVELS } from "../../src/design-bloom.mjs";
import { formTable, formTitleBlock, printLine } from "../../src/guide-forms-ooxml.mjs";
import { newFormId } from "../../src/guide-forms.mjs";
import { safeFileName } from "./word-files.js";

/**
 * Put a map (or any table of words) in the guide as a landscape form for the facilitator: a title, a print line and the table, after the block the cursor is in. At most 40 rows
 * go in the guide's table; the file has all of them.
 * @param {any} S @param {string} title @param {{ head: string[], body: string[][] }} t
 */
export async function putTableInGuide(S, title, t) {
  const paras = await S.guide.paragraphs();
  const id = newFormId(paras);
  const rows = t.body.slice(0, 40);
  const cut = t.body.length - rows.length;
  await S.guide.insert((ctx) => formTitleBlock({ id, label: "Template", number: "", title, facilitator: true }) + printLine({ copies: { n: 1, per: "room" }, orientation: "landscape" }) +
    formTable({ columns: t.head.map((name, i) => ({ name, weight: i === 2 ? 4 : 2 })), rows, ctx, caption: title }), { where: "after", after: Math.max(await S.guide.cursorIndex(), 0), spare: false, select: "last" });
  S.say(`${title} is in the guide, after the paragraph you were on${cut ? `. It holds the first 40 of ${t.body.length} rows; save the CSV for all of them` : ""}. Ctrl+Z in the guide takes it back.`);
}

/** @param {any} S @returns {HTMLElement} */
export function mapPanel(S) {
  const { doc } = S;
  const d = S.design;
  const panel = el(doc, "div", { class: "dd-panel" });
  panel.append(el(doc, "h3", { text: "Course map" }), note(doc, "Each objective with the activities that practice it and the assessments that check it. Plan the assessments here, then read the map: it says what is missing, and it becomes a table for a syllabus or an accreditation report."));

  // ---- the assessments
  const as = el(doc, "section", { "aria-labelledby": "dd-asm-h" }, el(doc, "h4", { id: "dd-asm-h", text: "Assessments in the plan" }));
  if (!d.assessments.length) as.append(note(doc, "None yet."));
  d.assessments.forEach((a, i) => {
    const goals = d.objectives.map((o, j) => ({ id: o.id, label: `Objective ${j + 1}: ${o.text.slice(0, 80) || "(no words yet)"}` }));
    as.append(el(doc, "div", { class: "dd-item" },
      field(doc, { label: `Assessment ${i + 1}: title`, value: a.title, max: 200, onInput: (v) => { a.title = v; S.changed(); } }),
      row(doc,
        choice(doc, { label: "Kind", value: a.kind, options: ASSESSMENT_KINDS.map((k) => ({ value: k.id, label: k.label })), onChange: (v) => { a.kind = v; S.changed(); } }),
        choice(doc, { label: "Checks at the level of", value: a.level, small: true, options: [{ value: "", label: "Not said" }, ...BLOOM_LEVELS.map((l) => ({ value: l.id, label: l.label }))], onChange: (v) => { a.level = v; S.changed(); S.render(); } }),
        field(doc, { label: "Item slots for each objective", value: a.slots ? String(a.slots) : "", type: "number", small: true, max: 50, hint: "For OpenAssess.", onInput: (v) => { const n = Number(v); a.slots = Number.isInteger(n) && n >= 0 && n <= 50 ? n : 0; S.changed(); } }),
      ),
      tickGroup(doc, { legend: `Assessment ${i + 1} checks these objectives`, items: goals, chosen: a.objectives, empty: "Write objectives on the Needs and objectives screen first.", onChange: (ids) => { a.objectives = ids; S.changed(); S.render(); } }),
      field(doc, { label: "Note", value: a.note, rows: 2, max: 4000, onInput: (v) => { a.note = v; S.changed(); } }),
      row(doc, button(doc, "Remove", () => { S.design = removeAssessment(S.design, a.id); S.changed(); S.render(); }, { small: true, danger: true, label: `Remove assessment ${i + 1}` })),
    ));
  });
  as.append(row(doc, button(doc, "Add an assessment", () => { d.assessments.push({ id: nextId("s", d.assessments), title: "", kind: "quiz", level: "", slots: 0, objectives: [], note: "" }); S.changed(); S.render(); S.focus(`Assessment ${d.assessments.length}`); })));
  panel.append(as);

  // ---- the two maps
  const drawMap = (map, caption) => {
    const box = el(doc, "section", { class: "dd-map", "aria-label": caption });
    box.append(el(doc, "h4", { text: caption }), el(doc, "p", { class: "dd-summary", text: mapSummaryText(map) }));
    if (map.rows.length) box.append(table(doc, { caption: `${caption}: each objective with its activities and assessments`, ...mapTable(map) }));
    const gaps = map.rows.flatMap((r, i) => r.issues.map((x) => `Objective ${i + 1}: ${x.message}`));
    const loose = [...map.loose.activities, ...map.loose.assessments].map((x) => `${x.title || "Untitled"}: ${x.message}`);
    if (gaps.length || loose.length) box.append(el(doc, "ul", { class: "dd-issues", "aria-label": `What is missing from the ${map.kind === "plan" ? "plan" : "guide"}'s map` }, ...[...gaps, ...loose].map((g) => el(doc, "li", { text: g }))));
    const name = S.fileName();
    box.append(row(doc,
      button(doc, "Save as CSV", () => { give(doc, mapCsv(map), `${safeFileName(name)} - course map (${map.kind === "plan" ? "plan" : "guide"}).csv`, "text/csv"); S.say("The course map was saved as a CSV file in your downloads. A spreadsheet opens it with its accents right."); }, { label: `Save the ${map.kind}'s course map as a CSV file`, small: true }),
      button(doc, "Put it in the guide", async () => { try { await putTableInGuide(S, "Course map", mapTable(map)); } catch (e) { S.say(e.message, true); } }, { label: `Put the ${map.kind}'s course map in the guide`, small: true }),
    ));
    return box;
  };
  panel.append(drawMap(planMap(d), "The plan's map"));
  const guideSection = el(doc, "section", {});
  panel.append(guideSection);
  S.guide.paragraphs().then((paras) => guideSection.replaceChildren(drawMap(guideMap(paras), "The guide's map"))).catch(() => {});

  // ---- the item slots
  const slots = el(doc, "section", { "aria-labelledby": "dd-slot-h" }, el(doc, "h4", { id: "dd-slot-h", text: "Item slots for OpenAssess" }));
  const n = itemSlots(d).length;
  slots.append(note(doc, n ? `${n} empty ${n === 1 ? "slot" : "slots"}, each tagged with its objective, level, module and lesson. In OpenAssess, choose Import, pick this file, and save the slots into a bank; each is an essay question with a placeholder stem, tagged "slot", for you to write.` : "No slots yet: give an assessment above a number of item slots, and the objectives it checks."));
  slots.append(row(doc, button(doc, "Save item slots (CSV)", () => { try { give(doc, itemSlotsCsv(d), `${safeFileName(S.fileName())} - item slots.csv`, "text/csv"); S.say("The item slots were saved as a CSV file in your downloads, for OpenAssess's Import."); } catch (e) { S.say(e.message, true); } }, { disabled: !n })));
  panel.append(slots);
  return panel;
}
