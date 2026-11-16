/**
 * The design desk's Storyboard: the course as modules and lessons, each lesson with its minutes, its objectives, its activities and its assessments. From it, when the author
 * asks: a new guide with the course plan written in it, a skeleton for OpenStage Author or OpenStage Scroll, and the item slots for OpenAssess. Each is a guide or a file the
 * author keeps; nothing is sent anywhere.
 */
import { el } from "./menubar.js";
import { field, tickGroup, button, row, note, give } from "./design-ui.js";
import { nextId, storyboardMinutes } from "../../src/design-model.mjs";
import { authorScreens, scrollCourse, scrollBytes, scrollFileName, itemSlotsCsv, itemSlots } from "../../src/design-storyboard.mjs";
import { authorProject, authorZipName } from "../../src/guide-author.mjs";
import { zipWorkbook } from "../../src/guide-workbook.mjs";
import { safeFileName, zipLibrary } from "./word-files.js";

const lessonsIn = (d) => d.storyboard.reduce((a, m) => a + m.lessons.length, 0);
const swap = (list, i, by) => {
  const j = i + by;
  if (j < 0 || j >= list.length) return false;
  [list[i], list[j]] = [list[j], list[i]];
  return true;
};

/** @param {any} S @returns {HTMLElement} */
export function boardPanel(S) {
  const { doc } = S;
  const d = S.design;
  const panel = el(doc, "div", { class: "dd-panel" });
  panel.append(el(doc, "h3", { text: "Storyboard" }), note(doc, "The course as modules and lessons. Say which objectives each lesson teaches, which activities are in it and which assessments check it. Then make the guide, or a skeleton for another tool, from it."));

  // The summary line is drawn again as the minutes are typed, so the total is always what the lessons add up to.
  const summary = el(doc, "p", { class: "dd-summary", role: "status" });
  const drawSummary = () => {
    const mins = storyboardMinutes(d);
    const lessons = lessonsIn(d);
    summary.textContent = d.storyboard.length ? `${d.storyboard.length} ${d.storyboard.length === 1 ? "module" : "modules"}, ${lessons} ${lessons === 1 ? "lesson" : "lessons"}, ${mins.total} minutes in all (${Math.floor(mins.total / 60)} h ${mins.total % 60} min).` : "No modules yet.";
  };
  drawSummary();
  panel.append(summary);

  d.storyboard.forEach((m, mi) => {
    const box = el(doc, "section", { class: "dd-module", "aria-label": `Module ${mi + 1}` });
    box.append(
      field(doc, { label: `Module ${mi + 1}: title`, value: m.title, max: 200, onInput: (v) => { m.title = v; S.changed(); } }),
      row(doc,
        button(doc, "Move up", () => { if (swap(d.storyboard, mi, -1)) { S.changed(); S.render(); S.focus(`Module ${mi}`); } }, { small: true, label: `Move module ${mi + 1} up`, disabled: mi === 0 }),
        button(doc, "Move down", () => { if (swap(d.storyboard, mi, 1)) { S.changed(); S.render(); S.focus(`Module ${mi + 2}`); } }, { small: true, label: `Move module ${mi + 1} down`, disabled: mi === d.storyboard.length - 1 }),
        button(doc, "Remove module", () => { d.storyboard.splice(mi, 1); S.changed(); S.render(); S.say(`Module ${mi + 1} and its lessons are removed from the storyboard. Its objectives, activities and assessments stay in the notes.`); }, { small: true, danger: true, label: `Remove module ${mi + 1} and its lessons` }),
      ),
    );
    m.lessons.forEach((l, li) => {
      const name = `${mi + 1}.${li + 1}`;
      const goals = d.objectives.map((o, j) => ({ id: o.id, label: `Objective ${j + 1}: ${o.text.slice(0, 70) || "(no words yet)"}` }));
      const acts = d.activities.map((a, j) => ({ id: a.id, label: `Activity ${j + 1}: ${a.title || "(untitled)"}${a.minutes ? ` (${a.minutes} min)` : ""}` }));
      const checks = d.assessments.map((a, j) => ({ id: a.id, label: `Assessment ${j + 1}: ${a.title || "(untitled)"}` }));
      box.append(el(doc, "div", { class: "dd-item dd-lesson" },
        row(doc,
          field(doc, { label: `Lesson ${name}: title`, value: l.title, max: 200, onInput: (v) => { l.title = v; S.changed(); } }),
          field(doc, { label: "Minutes", value: l.minutes ? String(l.minutes) : "", type: "number", small: true, onInput: (v) => { const n = Number(v); l.minutes = Number.isInteger(n) && n >= 0 && n <= 999 ? n : 0; drawSummary(); S.changed(); } }),
        ),
        tickGroup(doc, { legend: `Lesson ${name} teaches`, items: goals, chosen: l.objectives, empty: "No objectives yet.", onChange: (ids) => { l.objectives = ids; S.changed(); } }),
        tickGroup(doc, { legend: `Lesson ${name} has these activities`, items: acts, chosen: l.activities, empty: "No activities yet.", onChange: (ids) => { l.activities = ids; S.changed(); } }),
        tickGroup(doc, { legend: `Lesson ${name} is checked by`, items: checks, chosen: l.assessments, empty: "No assessments yet.", onChange: (ids) => { l.assessments = ids; S.changed(); } }),
        field(doc, { label: "Note", value: l.note, rows: 2, max: 4000, onInput: (v) => { l.note = v; S.changed(); } }),
        row(doc,
          button(doc, "Move up", () => { if (swap(m.lessons, li, -1)) { S.changed(); S.render(); S.focus(`Lesson ${mi + 1}.${li}`); } }, { small: true, label: `Move lesson ${name} up`, disabled: li === 0 }),
          button(doc, "Move down", () => { if (swap(m.lessons, li, 1)) { S.changed(); S.render(); S.focus(`Lesson ${mi + 1}.${li + 2}`); } }, { small: true, label: `Move lesson ${name} down`, disabled: li === m.lessons.length - 1 }),
          button(doc, "Remove lesson", () => { m.lessons.splice(li, 1); S.changed(); S.render(); }, { small: true, danger: true, label: `Remove lesson ${name}` }),
        ),
      ));
    });
    box.append(row(doc, button(doc, "Add a lesson", () => {
      const ids = d.storyboard.flatMap((x) => [x, ...x.lessons]);
      m.lessons.push({ id: nextId("l", ids), title: "", minutes: 0, objectives: [], activities: [], assessments: [], note: "" });
      S.changed(); S.render(); S.focus(`Lesson ${mi + 1}.${m.lessons.length}`);
    })));
    panel.append(box);
  });
  panel.append(row(doc, button(doc, "Add a module", () => {
    const ids = d.storyboard.flatMap((x) => [x, ...x.lessons]);
    d.storyboard.push({ id: nextId("m", ids), title: "", lessons: [] });
    S.changed(); S.render(); S.focus(`Module ${d.storyboard.length}`);
  }, { primary: !d.storyboard.length })));

  // ---- what the storyboard becomes
  const outputs = el(doc, "section", { "aria-labelledby": "dd-out-h", class: "dd-outputs" }, el(doc, "h4", { id: "dd-out-h", text: "Make something from the storyboard" }));
  const ready = d.storyboard.some((m) => m.lessons.length);
  const name = S.fileName();
  const attempt = (fn) => async () => { try { await fn(); } catch (e) { S.say(e.message, true); } };
  outputs.append(
    el(doc, "div", { class: "dd-output" }, el(doc, "p", {}, el(doc, "strong", { text: "A new guide" })), note(doc, "Modules and lessons with their times, an Objective block for each objective, the activities and assessments, and a Transition at the end of each module. It opens as a new guide beside the others; the guide you have open does not change."),
      row(doc, button(doc, "Make a new guide from the storyboard", attempt(() => S.newGuideFromPlan()), { primary: true, disabled: !ready }))),
    el(doc, "div", { class: "dd-output" }, el(doc, "p", {}, el(doc, "strong", { text: "A skeleton for OpenStage Author" })), note(doc, "A project with a scene for each lesson and a slide for each lesson, activity and assessment. Author has no field for an objective or its level, so they are in its notes file. Unzip it, then open the folder that ends in .osproj in Author."),
      row(doc, button(doc, "Save the Author skeleton", attempt(async () => {
        const files = authorProject(authorScreens(d), { title: S.title() });
        const bytes = await zipWorkbook(files, zipLibrary(), "uint8array");
        give(doc, bytes, authorZipName(S.title()), "application/zip");
        S.say(`The Author skeleton was saved as "${authorZipName(S.title())}" in your downloads.`);
      }), { disabled: !ready }))),
    el(doc, "div", { class: "dd-output" }, el(doc, "p", {}, el(doc, "strong", { text: "A skeleton for OpenStage Scroll" })), note(doc, "A .openscroll file with a section for each module and a lesson for each lesson, with its objectives and activities as lists. In Scroll, import the file."),
      row(doc, button(doc, "Save the Scroll skeleton", attempt(async () => {
        const bytes = await scrollBytes(scrollCourse(d, { title: S.title() }), zipLibrary());
        give(doc, bytes, scrollFileName(S.title()), "application/x-openscroll-project");
        S.say(`The Scroll skeleton was saved as "${scrollFileName(S.title())}" in your downloads.`);
      }), { disabled: !ready }))),
    el(doc, "div", { class: "dd-output" }, el(doc, "p", {}, el(doc, "strong", { text: "Item slots for OpenAssess" })), note(doc, itemSlots(d).length ? `${itemSlots(d).length} empty slots, each tagged with its objective and level. In OpenAssess, choose Import and pick the file.` : "Give an assessment a number of item slots on the Course map first."),
      row(doc, button(doc, "Save item slots (CSV)", attempt(async () => { give(doc, itemSlotsCsv(d), `${safeFileName(name)} - item slots.csv`, "text/csv"); S.say("The item slots were saved as a CSV file in your downloads."); }), { disabled: !itemSlots(d).length }))),
  );
  panel.append(outputs);
  return panel;
}
