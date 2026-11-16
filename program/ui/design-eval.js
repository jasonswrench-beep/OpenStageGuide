/**
 * The design desk's Evaluation plan, by Kirkpatrick's four levels: what is measured, of whom and when, for each level. The plan goes to PrismStatistics as a file with a variable
 * for every measure and empty rows to type or paste the answers into, and to the guide as a table. What is missing from the plan is said in words and is the author's to settle.
 */
import { el } from "./menubar.js";
import { field, choice, tickGroup, button, row, note } from "./design-ui.js";
import { KIRKPATRICK, nextId } from "../../src/design-model.mjs";
import { kindsForLevel, evaluationIssues, evaluationProject, evaluationTable } from "../../src/design-evaluation.mjs";
import { pstatBytes, MAX_ROWS } from "../../src/guide-pstat.mjs";
import { safeFileName, zipLibrary } from "./word-files.js";
import { putTableInGuide } from "./design-mapview.js";
import { give } from "./design-ui.js";

/** @param {any} S @returns {HTMLElement} */
export function evalPanel(S) {
  const { doc } = S;
  const d = S.design;
  const panel = el(doc, "div", { class: "dd-panel" });
  panel.append(el(doc, "h3", { text: "Evaluation plan" }), note(doc, "How you will know the course worked, by Kirkpatrick's four levels: how they found it, what they learned, whether they do it at work, and whether the work changed. Add a measure for each level you will evaluate."));

  const measures = d.evaluation.measures;
  for (const k of KIRKPATRICK) {
    const box = el(doc, "section", { class: "dd-levelbox", "aria-label": k.label });
    box.append(el(doc, "h4", { text: k.label }), el(doc, "p", { class: "hint", text: `${k.asks} Usually: ${k.kinds.join(", ")}.` }));
    const mine = measures.filter((m) => m.level === k.id);
    if (!mine.length) box.append(note(doc, "Nothing planned for this level yet."));
    for (const m of mine) {
      const n = mine.indexOf(m) + 1;
      const kinds = kindsForLevel(k.id);
      if (!kinds.some((x) => x.id === m.kind)) m.kind = kinds[0].id;
      const goals = d.objectives.map((o, j) => ({ id: o.id, label: `Objective ${j + 1}: ${o.text.slice(0, 70) || "(no words yet)"}` }));
      const stm = field(doc, { label: "Statements, one to a line", value: m.statements.join("\n"), rows: 4, max: 4000, hint: "Each is rated on the scale below.", onInput: (v) => { m.statements = v.split(/\r?\n/).map((x) => x.trim()).filter(Boolean); S.changed(); } });
      const scale = field(doc, { label: "The points of the scale", value: m.scale, max: 200, hint: "Two to seven, separated by commas: 1, 2, 3, 4, 5 or Rarely, Sometimes, Often.", onInput: (v) => { m.scale = v; S.changed(); } });
      const showRating = () => { stm.hidden = scale.hidden = m.kind !== "rating"; };
      box.append(el(doc, "div", { class: "dd-item" },
        field(doc, { label: `${k.label}, measure ${n}: name`, value: m.name, max: 200, onInput: (v) => { m.name = v; S.changed(); } }),
        choice(doc, { label: "What is measured", value: m.kind, options: kinds.map((x) => ({ value: x.id, label: x.label })), onChange: (v) => { m.kind = v; showRating(); S.changed(); } }),
        stm, scale,
        row(doc,
          field(doc, { label: "When", value: m.when, max: 200, hint: "For example: at the end of the second day.", onInput: (v) => { m.when = v; S.changed(); } }),
          field(doc, { label: "Of whom", value: m.who, max: 200, hint: "For example: every learner, or their manager.", onInput: (v) => { m.who = v; S.changed(); } }),
        ),
        tickGroup(doc, { legend: `${k.label}, measure ${n}: for these objectives`, items: goals, chosen: m.objectives, empty: "No objectives yet.", onChange: (ids) => { m.objectives = ids; S.changed(); } }),
        field(doc, { label: "Note", value: m.note, rows: 2, max: 4000, onInput: (v) => { m.note = v; S.changed(); } }),
        row(doc, button(doc, "Remove", () => { measures.splice(measures.indexOf(m), 1); S.changed(); S.render(); }, { small: true, danger: true, label: `Remove ${k.label} measure ${n}` })),
      ));
      showRating();
    }
    box.append(row(doc, button(doc, "Add a measure", () => {
      measures.push({ id: nextId("e", measures), level: k.id, name: "", kind: kindsForLevel(k.id)[0].id, when: "", who: "", scale: k.id === 1 || k.id === 3 ? "1, 2, 3, 4, 5" : "", statements: [], objectives: [], note: "" });
      S.changed(); S.render(); S.focus(`${k.label}, measure ${measures.filter((m) => m.level === k.id).length}`);
    }, { small: true, label: `Add a measure to ${k.label}` })));
    panel.append(box);
  }
  panel.append(field(doc, { label: "Notes on the plan as a whole", value: d.evaluation.notes, rows: 3, max: 4000, onInput: (v) => { d.evaluation.notes = v; S.changed(); } }));

  // ---- what is missing, and what it becomes
  const issues = evaluationIssues(d);
  const out = el(doc, "section", { "aria-labelledby": "dd-eval-out" }, el(doc, "h4", { id: "dd-eval-out", text: "What the plan is missing" }));
  out.append(issues.length ? el(doc, "ul", { class: "dd-issues" }, ...issues.map((x) => el(doc, "li", { text: x }))) : note(doc, "Nothing: every level has a measure, with its time."));
  const people = el(doc, "div", {});
  const rows = field(doc, { label: `People to enter answers for (0 to ${MAX_ROWS})`, value: String(S.evalRows ?? 30), type: "number", small: true, hint: "One empty row for each; rows can be added in PrismStatistics later.", onInput: (v) => { S.evalRows = v; } });
  people.append(rows);
  out.append(
    el(doc, "p", { class: "hint", text: "The file for PrismStatistics has a variable for each measure (the level in its label), empty rows for the answers, and the plan in the notes it carries. Type or paste the answers in, then analyse." }),
    people,
    row(doc,
      button(doc, "Save for PrismStatistics", async () => {
        try {
          const n = String(S.evalRows ?? "30").trim() === "" ? 0 : Number(String(S.evalRows ?? "30").trim());
          const bytes = await pstatBytes(evaluationProject(S.design, { guideTitle: S.title(), rows: n, appVersion: S.appVersion() }), zipLibrary());
          const file = `${safeFileName(S.fileName())} - evaluation plan.pstat`;
          give(doc, bytes, file, "application/octet-stream");
          S.say(`${file} was saved in your downloads. Open it in PrismStatistics; its notes carry the plan.`);
        } catch (e) { S.say(e.message, true); }
      }, { primary: true, disabled: !measures.length }),
      button(doc, "Put the plan in the guide", async () => { try { await putTableInGuide(S, "Evaluation plan", evaluationTable(S.design)); } catch (e) { S.say(e.message, true); } }, { disabled: !measures.length }),
    ),
  );
  panel.append(out);
  return panel;
}
