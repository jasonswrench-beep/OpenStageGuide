/**
 * The design desk's first screen: the needs analysis (what is wrong, for whom, and whether training is the answer) and the objectives, written with the verbs of Bloom's levels.
 * Everything here is advice: a weak verb or a long objective is pointed out and never refused, and the author's own words are what is kept.
 */
import { el } from "./menubar.js";
import { field, choice, tickGroup, button, row, heading, note } from "./design-ui.js";
import { BLOOM_LEVELS, checkObjective, composeObjective, levelById } from "../../src/design-bloom.mjs";
import { NEED_CAUSES, nextId, removeObjective } from "../../src/design-model.mjs";

const AUTO = "";

/** What the ticked causes say about whether training is the answer. @param {string[]} causes */
export function trainingVerdict(causes) {
  if (!causes.length) return "";
  const help = NEED_CAUSES.filter((c) => causes.includes(c.id) && c.training).map((c) => c.label.toLowerCase());
  const other = NEED_CAUSES.filter((c) => causes.includes(c.id) && !c.training);
  if (help.length && !other.length) return "Training can help with this.";
  if (help.length) return `Training can help with part of it. For the rest (${other.map((c) => c.label.toLowerCase()).join("; ")}) something else is needed too.`;
  return `Training alone will not fix this: ${other.map((c) => c.label.toLowerCase()).join("; ")}. ${other[0]?.note ?? ""}`.trim();
}

/** @param {any} S the desk's state @returns {HTMLElement} */
export function needsPanel(S) {
  const { doc } = S;
  const d = S.design;
  const panel = el(doc, "div", { class: "dd-panel" });

  // ---- the needs analysis
  const needs = el(doc, "section", { "aria-labelledby": "dd-needs-h" }, el(doc, "h3", { id: "dd-needs-h", text: "Needs analysis" }));
  needs.append(note(doc, "Before any objective: what is wrong, who it is wrong for, and whether training is the answer. Write a line or two for each; nothing here is required."));
  const text = (key, label, rows = 2, hint) => field(doc, { label, value: d.needs[key], rows, hint, max: 4000, onInput: (v) => { d.needs[key] = v; S.changed(); } });
  needs.append(text("problem", "What is the problem, or the chance?", 3, "For example: new supervisors avoid hard conversations, and small problems become formal complaints."));
  needs.append(text("audience", "Who are the learners?", 2));
  needs.append(text("current", "What do they do now?", 2));
  needs.append(text("desired", "What should they do?", 2));
  const verdict = el(doc, "p", { class: "dd-verdict", role: "status", text: trainingVerdict(d.needs.causes) });
  needs.append(
    tickGroup(doc, {
      legend: "What is in the way?",
      items: NEED_CAUSES.map((c) => ({ id: c.id, label: c.label, note: c.note })),
      chosen: d.needs.causes,
      onChange: (ids) => {
        d.needs.causes = ids;
        verdict.textContent = trainingVerdict(ids);
        S.changed();
      },
    }),
    verdict,
    text("constraints", "What limits the course? (time, money, place, technology)", 2),
    text("success", "How will you know it worked?", 2),
    text("delivery", "How will it be delivered?", 2),
  );
  panel.append(needs);

  // ---- the objectives
  const section = el(doc, "section", { "aria-labelledby": "dd-obj-h" }, el(doc, "h3", { id: "dd-obj-h", text: "Objectives" }));
  section.append(note(doc, "What learners will be able to do, in words that can be seen. The level comes from the verb unless you choose one. A weak verb, two verbs in one objective or a very long one is pointed out, and is yours to keep."));
  const list = el(doc, "div", { class: "dd-list" });
  const spread = el(doc, "p", { class: "hint dd-spread", "aria-live": "polite" });
  const drawSpread = () => {
    const counts = BLOOM_LEVELS.map((l) => [l.label, d.objectives.filter((o) => (o.level || checkObjective(o.text).level) === l.id).length]);
    const none = d.objectives.filter((o) => !(o.level || checkObjective(o.text).level)).length;
    spread.textContent = d.objectives.length
      ? `${d.objectives.length} ${d.objectives.length === 1 ? "objective" : "objectives"}: ${counts.filter(([, n]) => n).map(([l, n]) => `${l} ${n}`).join(", ") || "none with a level yet"}${none ? `; ${none} with no level` : ""}.`
      : "";
  };
  const drawOne = (o, i) => {
    const issues = el(doc, "ul", { class: "dd-issues", "aria-live": "polite" });
    const levelLine = el(doc, "p", { class: "dd-level" });
    const redraw = () => {
      const r = checkObjective(o.text);
      const level = o.level || r.level || "";
      levelLine.textContent = level ? `Level: ${levelById(level).label}${o.level ? " (chosen)" : " (from the verb)"}` : o.text.trim() ? "Level: not known yet" : "";
      issues.replaceChildren(...(o.text.trim() ? r.issues.filter((x) => x.rule !== "OBJ-EMPTY") : []).map((x) => el(doc, "li", { text: x.message })));
      drawSpread();
    };
    const box = field(doc, { label: `Objective ${i + 1}: what learners will be able to do`, value: o.text, rows: 2, max: 4000, onInput: (v) => { o.text = v; redraw(); S.changed(); } });
    const level = choice(doc, {
      label: "Level",
      small: true,
      value: o.level || AUTO,
      options: [{ value: AUTO, label: "From the verb" }, ...BLOOM_LEVELS.map((l) => ({ value: l.id, label: l.label }))],
      onChange: (v) => { o.level = v; redraw(); S.changed(); },
    });
    const gone = button(doc, "Remove", () => { S.design = removeObjective(S.design, o.id); S.changed(); S.render(); S.say(`Objective ${i + 1} is removed from the notes, and from every activity, assessment and lesson that named it.`); }, { small: true, danger: true, label: `Remove objective ${i + 1}` });
    redraw();
    return el(doc, "div", { class: "dd-item" }, box, row(doc, level, gone), levelLine, issues);
  };
  const draw = () => { list.replaceChildren(...d.objectives.map(drawOne)); drawSpread(); };
  draw();
  section.append(list, spread);
  const add = (text = "", level = "") => {
    d.objectives.push({ id: nextId("o", d.objectives), text, level });
    S.changed();
    S.render();
    S.focus(`Objective ${d.objectives.length}`);
  };
  section.append(row(doc, button(doc, "Add an objective", () => add(), { primary: true })));

  // ---- one written from verbs
  const helper = el(doc, "details", { class: "dd-helper" }, el(doc, "summary", { text: "Write one from the verbs of a level" }));
  const parts = { level: "apply", verb: "", what: "", condition: "", standard: "" };
  const preview = el(doc, "p", { class: "dd-preview", "aria-live": "polite" });
  const verbs = el(doc, "div", {});
  const asks = el(doc, "p", { class: "hint" });
  const sync = () => {
    const l = levelById(parts.level);
    asks.textContent = l.asks;
    if (!l.verbs.includes(parts.verb)) parts.verb = l.verbs[0];
    verbs.replaceChildren(choice(doc, { label: "Verb", value: parts.verb, options: l.verbs.map((v) => ({ value: v, label: v })), onChange: (v) => { parts.verb = v; show(); } }));
    show();
  };
  const show = () => { try { preview.textContent = composeObjective({ condition: parts.condition, verb: parts.verb, what: parts.what, standard: parts.standard }); } catch { preview.textContent = ""; } };
  helper.append(
    choice(doc, { label: "Level", value: parts.level, options: BLOOM_LEVELS.map((l) => ({ value: l.id, label: l.label })), onChange: (v) => { parts.level = v; sync(); } }),
    asks,
    verbs,
    field(doc, { label: "What", hint: "For example: the three steps of a coaching conversation.", max: 300, onInput: (v) => { parts.what = v; show(); } }),
    field(doc, { label: "Under what conditions (optional)", hint: "For example: after the role play.", max: 200, onInput: (v) => { parts.condition = v; show(); } }),
    field(doc, { label: "How well (optional)", hint: "For example: without notes, in under five minutes.", max: 200, onInput: (v) => { parts.standard = v; show(); } }),
    preview,
    row(doc, button(doc, "Add this objective", () => { if (!parts.what.trim()) { S.say("Say what learners will do the verb to, for example the three steps.", true); return; } add(composeObjective(parts), parts.level); }, { primary: true })),
  );
  sync();
  section.append(helper);
  panel.append(section);
  return panel;
}
