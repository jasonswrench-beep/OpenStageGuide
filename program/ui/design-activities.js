/**
 * The design desk's Models screen: the learning design models (each event with activities that fit it), the activities the author has planned, and for each model the work it
 * comes from, cited in APA 7 by the program and, when Prism Reference Manager is connected, by Prism Reference Manager. The author chooses; nothing is added unasked.
 */
import { el } from "./menubar.js";
import { field, tickGroup, button, row, note } from "./design-ui.js";
import { MODELS, modelById, sourceFor, referenceText, inTextCitation } from "../../src/design-models.mjs";
import { nextId, removeActivity } from "../../src/design-model.mjs";
import { planReference } from "../../src/guide-reference.mjs";
import { post } from "./api.js";

/** @param {any} S @returns {HTMLElement} */
export function activitiesPanel(S) {
  const { doc } = S;
  const d = S.design;
  const panel = el(doc, "div", { class: "dd-panel" });
  panel.append(el(doc, "h3", { text: "Models and activities" }), note(doc, "Choose a model to see its events and activities that fit each. Add the ones you want; they join the plan, where each is linked to the objectives it serves."));

  const chosen = S.modelId && modelById(S.modelId) ? S.modelId : MODELS[0].id;
  S.modelId = chosen;
  const layout = el(doc, "div", { class: "dd-split" });
  const nav = el(doc, "nav", { "aria-label": "Models", class: "dd-nav" }, ...MODELS.map((m) => {
    const b = button(doc, m.name, () => { S.modelId = m.id; S.render(); S.focus(m.name); }, { label: `${m.name}, by ${m.by}` });
    b.setAttribute("aria-pressed", String(m.id === chosen));
    return el(doc, "div", {}, b);
  }));
  layout.append(nav);

  const m = modelById(chosen);
  const detail = el(doc, "div", { class: "dd-detail" });
  detail.append(el(doc, "h4", { text: m.name }), el(doc, "p", { class: "hint", text: `By ${m.by}. ${m.kind}.` }), el(doc, "p", { text: m.summary }), el(doc, "p", { class: "hint", text: `Use it when: ${m.useWhen}` }));
  const addActivity = (title, event) => {
    d.activities.push({ id: nextId("a", d.activities), title, model: m.id, event, minutes: 0, objectives: [], note: "" });
    if (!d.models.includes(m.id)) d.models.push(m.id);
    S.changed();
    S.render();
    S.say(`"${title}" is in the plan, under Activities below. Link it to the objectives it serves.`);
  };
  const events = el(doc, "ol", { class: "dd-events" });
  for (const e of m.events) {
    const ideas = el(doc, "ul", { class: "dd-ideas" }, ...e.ideas.map((idea) => el(doc, "li", {}, el(doc, "span", { text: idea }), button(doc, "Add", () => addActivity(idea, e.name), { small: true, label: `Add the activity: ${idea}, for ${e.name}` }))));
    events.append(el(doc, "li", {}, el(doc, "strong", { text: e.name }), ideas));
  }
  detail.append(events);

  // ---- where it comes from
  const src = sourceFor(m.id);
  const cite = el(doc, "section", { class: "dd-cite", "aria-label": `The source of ${m.name}` }, el(doc, "h4", { text: "Where it comes from" }));
  const refText = referenceText(src);
  cite.append(el(doc, "p", { class: "dd-ref", text: refText }), el(doc, "p", { class: "hint", text: `In a sentence: ${inTextCitation(src)}. This is the program's own APA 7 text${src.doi ? "" : "; this work is a book, so check its edition and year against your copy"}.` }));
  const where = el(doc, "div", { class: "dd-prm" });
  const result = el(doc, "div", { class: "dd-prm-result", "aria-live": "polite" });
  const putInGuide = async (text) => {
    const paras = await S.guide.paragraphs();
    const plan = planReference(paras, text);
    await S.guide.insert(() => plan.body, { where: "after", after: plan.after, select: "last" });
    S.say(`The reference is in the guide's list of references${plan.created ? ", which was started at the end of the guide" : ""}. Ctrl+Z in the guide takes it back.`);
  };
  cite.append(row(doc, button(doc, "Put this reference in the guide", async () => { try { await putInGuide(refText); } catch (e) { S.say(e.message, true); } }, { label: `Put the reference of ${m.name} in the guide's list of references` })));
  cite.append(prmBlock(S, where, result, m.source, putInGuide), where, result);
  detail.append(cite);
  layout.append(detail);
  panel.append(layout);

  // ---- the activities in the plan
  const acts = el(doc, "section", { "aria-labelledby": "dd-acts-h" }, el(doc, "h3", { id: "dd-acts-h", text: "Activities in the plan" }));
  if (!d.activities.length) acts.append(note(doc, "None yet. Add them from a model above, or write your own."));
  d.activities.forEach((a, i) => {
    const goals = d.objectives.map((o, j) => ({ id: o.id, label: `Objective ${j + 1}: ${o.text.slice(0, 80) || "(no words yet)"}` }));
    const model = a.model ? modelById(a.model) : null;
    acts.append(el(doc, "div", { class: "dd-item" },
      field(doc, { label: `Activity ${i + 1}: title`, value: a.title, max: 200, onInput: (v) => { a.title = v; S.changed(); } }),
      row(doc,
        field(doc, { label: "Minutes", value: a.minutes ? String(a.minutes) : "", type: "number", small: true, onInput: (v) => { const n = Number(v); a.minutes = Number.isInteger(n) && n >= 0 && n <= 999 ? n : 0; S.changed(); } }),
        button(doc, "Remove", () => { S.design = removeActivity(S.design, a.id); S.changed(); S.render(); }, { small: true, danger: true, label: `Remove activity ${i + 1}` }),
      ),
      model ? el(doc, "p", { class: "hint", text: `From ${model.name}${a.event ? `: ${a.event}` : ""}.` }) : el(doc, "p", { class: "hint", text: "Your own." }),
      tickGroup(doc, { legend: `Activity ${i + 1} serves these objectives`, items: goals, chosen: a.objectives, empty: "Write objectives on the Needs and objectives screen first.", onChange: (ids) => { a.objectives = ids; S.changed(); } }),
      field(doc, { label: "Note", value: a.note, rows: 2, max: 4000, onInput: (v) => { a.note = v; S.changed(); } }),
    ));
  });
  acts.append(row(doc, button(doc, "Add an activity of your own", () => { d.activities.push({ id: nextId("a", d.activities), title: "", model: "", event: "", minutes: 0, objectives: [], note: "" }); S.changed(); S.render(); S.focus(`Activity ${d.activities.length}`); })));
  panel.append(acts);
  return panel;
}

/**
 * The part of the screen that asks Prism Reference Manager: whether it is connected, to connect, and to cite. The token never comes here: the program's server holds it.
 * @param {any} S @param {HTMLElement} host @param {HTMLElement} result @param {string} source @param {(text: string) => Promise<void>} putInGuide
 */
function prmBlock(S, host, result, source, putInGuide) {
  const { doc } = S;
  host.replaceChildren();
  const status = el(doc, "p", { class: "dd-prm-status", role: "status", text: "Prism Reference Manager: not checked." });
  const portBox = field(doc, { label: "Its port (Settings, Local connections)", value: S.prmPort ?? "23119", small: true, max: 5, onInput: (v) => { S.prmPort = v; } });
  const actions = el(doc, "div", { class: "dd-row" });
  const ask = async (body) => {
    try { return await post("prm", body); } catch (e) { return { ok: false, message: e.message }; }
  };
  const draw = (r) => {
    const state = r?.state ?? "not-connected";
    status.textContent = `Prism Reference Manager: ${r?.ok === false ? r.message : state === "connected" ? `connected${r.reachable === false ? ", but not answering now. Start it to cite" : ""}.` : state === "pending" ? r.message : state === "denied" ? r.message : r?.message ? r.message : r?.reachable ? "it is running, and this program is not allowed yet. Choose Connect, then Allow in its window." : "not connected. Turn on Allow local connections in its Settings, then choose Connect."}`;
    actions.replaceChildren(
      ...(state === "connected"
        ? [button(doc, "Cite this with Prism Reference Manager", async () => {
            status.textContent = "Asking Prism Reference Manager...";
            const c = await ask({ action: "cite", source });
            if (!c.ok) { draw(c.code === "unauthorized" ? { state: "not-connected", message: c.message } : { ok: false, state: "connected", message: c.message }); return; }
            result.replaceChildren(
              el(doc, "p", { class: "dd-ref", text: c.reference }),
              el(doc, "p", { class: "hint", text: `In a sentence: ${c.inText}. Formatted by Prism Reference Manager in ${c.style}${c.added ? "; the work was added to its library (see Machine-created)" : ""}.` }),
              row(doc, button(doc, "Put this reference in the guide", async () => { try { await putInGuide(c.reference); } catch (e) { S.say(e.message, true); } })),
            );
            status.textContent = "Prism Reference Manager: connected.";
          }, { primary: true }), button(doc, "Disconnect", async () => { draw(await ask({ action: "forget" })); result.replaceChildren(); })]
        : [portBox, button(doc, "Connect", async () => { draw(await ask({ action: "connect", port: S.prmPort ?? "23119" })); }, { primary: state !== "pending" }), button(doc, "Check again", async () => { draw(await ask({ action: "status" })); })]),
    );
  };
  host.append(el(doc, "h4", { text: "Cite it with Prism Reference Manager (optional)" }), note(doc, "Prism Reference Manager can format the citation in your style and keep the work in your library. It must be running, with Allow local connections on; you choose Allow in its window the first time. Without it, everything on this screen still works."), status, actions);
  void ask({ action: "status" }).then(draw);
  return el(doc, "span", { hidden: "" });
}
