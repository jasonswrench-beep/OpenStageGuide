// @ts-check
/**
 * The Styles panel: every style of the guide with what it comes to, and the means to change it or to make one of your own. A change is made to
 * the style, so every paragraph in it follows (and so does the Word file made from the guide). Only what is changed in a style is stored; the
 * rest goes on following the style it is based on. The panel talks to the guide's document (listStyles, styleEdits, setStyleEdits) and changes
 * nothing else. Text and attributes only, never markup.
 */
import { contrastNeeded, contrastRatio, diffSpec, idForName, mergeSpec, specCss } from "../../src/guide-styles-edit.mjs";
import { askName, openDialog } from "./dialogs.js";
import { el } from "./menubar.js";
import { drawCss } from "./preview.js";

/** Fonts offered to start from. Any name can be typed: a Word file only names its font, and the computer it is opened on has to have it. */
const FONTS = ["Aptos", "Aptos Display", "Calibri", "Cambria", "Arial", "Verdana", "Tahoma", "Segoe UI", "Georgia", "Times New Roman", "Trebuchet MS", "Atkinson Hyperlegible", "Lexend", "Source Sans 3", "Source Serif 4", "OpenDyslexic"];

/** @param {any} v */
const show = (v) => (v === undefined || v === null ? "" : String(v));
/** @param {string} v */
const number = (v) => (v.trim() === "" ? undefined : Number(v.trim()));

/**
 * @param {Document} doc
 * @param {any} guide the guide's document (listStyles, styleEdits, styleUse, setStyleEdits)
 * @param {(message: string, bad?: boolean) => void} say
 * @param {string} [startId] the style to show first
 */
export function openStylesPanel(doc, guide, say, startId) {
  let session = 0;
  /** @type {any[]} */
  let list = [];
  /** @type {any} */
  let edits = { overrides: {}, custom: [] };
  /** @type {any} */
  let current = null;
  /** The spec the chosen style came to when it was shown: what the fields are compared with. @type {any} */
  let loaded = {};

  const field = (/** @type {string} */ id) => /** @type {HTMLInputElement} */ (doc.getElementById(`sp-${id}`));

  function reload(/** @type {string | undefined} */ keep) {
    list = guide.listStyles().filter((/** @type {any} */ s) => s.type === "paragraph");
    edits = guide.styleEdits();
    const select = /** @type {HTMLSelectElement | null} */ (doc.getElementById("sp-list"));
    if (!select) return;
    const mine = list.filter((s) => s.custom);
    const own = list.filter((s) => !s.custom);
    select.replaceChildren();
    /** @param {string} label @param {any[]} items */
    const group = (label, items) => {
      if (!items.length) return;
      const g = el(doc, "optgroup", { label });
      for (const s of items) {
        const changed = !!edits.overrides[s.id];
        g.append(el(doc, "option", { value: s.id, text: `${s.name}${changed ? " (changed)" : ""}` }));
      }
      select.append(g);
    };
    group("Your styles", mine);
    group("Styles of the guide", own);
    const want = keep && list.some((s) => s.id === keep) ? keep : current?.id && list.some((s) => s.id === current.id) ? current.id : (list.find((s) => s.name === "Normal") ?? list[0])?.id;
    if (want) {
      select.value = want;
      choose(want);
    }
  }

  function choose(/** @type {string} */ id) {
    current = list.find((s) => s.id === id) ?? null;
    if (!current) return;
    const e = current.effective;
    loaded = { font: { ...(e.font || {}) }, paragraph: { ...(e.paragraph || {}) }, ...(e.shading ? { shading: e.shading } : {}) };
    // Not set anywhere is not bold, not italic, and so on: the fields show a style as it reads.
    for (const k of ["bold", "italic", "underline", "strike"]) if (loaded.font[k] === undefined) loaded.font[k] = false;
    if (loaded.paragraph.keepWithNext === undefined) loaded.paragraph.keepWithNext = false;
    /** @type {HTMLElement} */ (doc.getElementById("sp-name")).textContent = current.name;
    /** @type {HTMLElement} */ (doc.getElementById("sp-based")).textContent = current.custom
      ? `Your own style, based on ${list.find((s) => s.id === current.basedOn)?.name ?? current.basedOn}.`
      : current.basedOn
        ? `Based on ${list.find((s) => s.id === current.basedOn)?.name ?? current.basedOn}.`
        : "The guide's ordinary text. Every other style starts from it.";
    field("font").value = show(loaded.font.name);
    field("size").value = show(loaded.font.size);
    field("color").value = show(loaded.font.color);
    field("colorpick").value = loaded.font.color ?? "#000000";
    for (const k of ["bold", "italic", "underline", "strike"]) field(k).checked = loaded.font[k] === true;
    field("alignment").value = loaded.paragraph.alignment ?? "left";
    field("before").value = show(loaded.paragraph.spaceBefore);
    field("after").value = show(loaded.paragraph.spaceAfter);
    ensureLine(loaded.paragraph.lineSpacing);
    field("line").value = show(loaded.paragraph.lineSpacing);
    field("left").value = show(loaded.paragraph.leftIndent);
    field("first").value = show(loaded.paragraph.firstLineIndent);
    field("keep").checked = loaded.paragraph.keepWithNext === true;
    field("shading").value = show(loaded.shading);
    field("shadingpick").value = loaded.shading ?? "#FFFFFF";
    /** @type {HTMLButtonElement} */ (doc.getElementById("sp-reset")).disabled = current.custom || !edits.overrides[current.id];
    /** @type {HTMLButtonElement} */ (doc.getElementById("sp-delete")).disabled = !current.custom;
    problem("");
    refreshSample();
  }

  /** @param {number | undefined} v */
  function ensureLine(v) {
    const select = /** @type {HTMLSelectElement} */ (field("line"));
    if (v !== undefined && ![...select.options].some((o) => o.value === String(v))) select.append(el(doc, "option", { value: String(v), text: String(v) }));
  }

  function problem(/** @type {string} */ message) {
    const p = /** @type {HTMLElement} */ (doc.getElementById("sp-problem"));
    p.textContent = message;
    p.hidden = !message;
  }

  /** The settings the fields hold now, as a spec. */
  function fieldsSpec() {
    /** @type {any} */
    const spec = { font: {}, paragraph: {} };
    const f = spec.font;
    const p = spec.paragraph;
    const text = (/** @type {string} */ id) => field(id).value.trim();
    if (text("font")) f.name = text("font");
    if (number(text("size")) !== undefined) f.size = number(text("size"));
    if (text("color")) f.color = text("color").toUpperCase();
    for (const k of ["bold", "italic", "underline", "strike"]) f[k] = field(k).checked;
    p.alignment = field("alignment").value;
    if (number(text("before")) !== undefined) p.spaceBefore = number(text("before"));
    if (number(text("after")) !== undefined) p.spaceAfter = number(text("after"));
    if (number(text("line")) !== undefined) p.lineSpacing = number(text("line"));
    if (number(text("left")) !== undefined) p.leftIndent = number(text("left"));
    if (number(text("first")) !== undefined) p.firstLineIndent = number(text("first"));
    p.keepWithNext = field("keep").checked;
    if (text("shading")) spec.shading = text("shading").toUpperCase();
    return spec;
  }

  function refreshSample() {
    const sample = /** @type {HTMLElement} */ (doc.getElementById("sp-sample"));
    sample.style.cssText = "";
    const spec = fieldsSpec();
    // A value that is not a number or a colour yet (half typed) is left out of the sample.
    const ok = JSON.parse(JSON.stringify(spec, (k, v) => (typeof v === "number" && !Number.isFinite(v) ? undefined : v)));
    if (ok.font?.color && !/^#[0-9A-F]{6}$/.test(ok.font.color)) delete ok.font.color;
    if (ok.shading && !/^#[0-9A-F]{6}$/.test(ok.shading)) delete ok.shading;
    const css = specCss(ok);
    // The sample shows the type, not the margins and indents (the panel is not the page).
    for (const k of ["marginTop", "marginBottom", "marginLeft", "textIndent"]) delete css[k];
    drawCss(sample, css);
    const fg = ok.font?.color ?? "#1B1B1B";
    const bg = ok.shading ?? "#FFFFFF";
    const ratio = contrastRatio(fg, bg);
    const need = contrastNeeded(ok.font?.size, ok.font?.bold);
    const note = /** @type {HTMLElement} */ (doc.getElementById("sp-contrast"));
    if (ratio === null) note.textContent = "";
    else if (ratio < need) note.textContent = `Contrast ${ratio} to 1: too low for text this size. People with low vision, and anyone reading a printout, need at least ${need} to 1.`;
    else note.textContent = `Contrast ${ratio} to 1: good.`;
    note.classList.toggle("bad", ratio !== null && ratio < need);
  }

  /** @param {Record<string, any>} next */
  async function save(next) {
    try {
      await guide.setStyleEdits(next, { group: `panel:${session}` });
      return true;
    } catch (e) {
      problem(/** @type {Error} */ (e).message);
      return false;
    }
  }

  async function apply() {
    if (!current) return;
    const change = diffSpec(loaded, fieldsSpec());
    if (!Object.keys(change).length) {
      say("Nothing was changed in this style.");
      return;
    }
    const next = JSON.parse(JSON.stringify(edits));
    if (current.custom) {
      const mine = next.custom.find((/** @type {any} */ c) => c.id === current.id);
      mine.spec = mergeSpec(mine.spec, change);
    } else next.overrides[current.id] = mergeSpec(next.overrides[current.id] ?? {}, change);
    if (await save(next)) {
      say(`${current.name} changed. Every paragraph in it follows.`);
      reload(current.id);
    }
  }

  async function reset() {
    if (!current || current.custom) return;
    const next = JSON.parse(JSON.stringify(edits));
    delete next.overrides[current.id];
    if (await save(next)) {
      say(`${current.name} is as the program makes it again.`);
      reload(current.id);
    }
  }

  async function remove() {
    if (!current || !current.custom) return;
    const next = JSON.parse(JSON.stringify(edits));
    next.custom = next.custom.filter((/** @type {any} */ c) => c.id !== current.id);
    const name = current.name;
    if (await save(next)) {
      say(`${name} was deleted.`);
      reload(undefined);
    }
  }

  async function create() {
    if (!current) return;
    const name = await askName(doc, { title: "New style", label: "Name of the style", confirm: "Create", hint: `It starts as a copy of ${current.name}, and follows it when that changes.` });
    if (!name) return;
    const clean = name.replace(/\s+/g, " ").trim();
    if (list.some((s) => s.name.toLowerCase() === clean.toLowerCase())) {
      problem(`There is already a style called ${clean}.`);
      return;
    }
    const id = idForName(clean, [...list.map((s) => s.id), ...edits.custom.map((/** @type {any} */ c) => c.id)]);
    const next = JSON.parse(JSON.stringify(edits));
    next.custom.push({ id, name: clean, type: "paragraph", basedOn: current.id, spec: {} });
    if (await save(next)) {
      say(`${clean} is a style of your own now. Change it here, then give it to paragraphs with Kind of paragraph.`);
      reload(id);
    }
  }

  /** @param {string} id @param {string} label @param {string} [type] @param {Record<string, string>} [extra] */
  const input = (id, label, type = "text", extra = {}) =>
    el(doc, "div", { class: "field" }, el(doc, "label", { class: "field-label", for: `sp-${id}`, text: label }), el(doc, "input", { id: `sp-${id}`, type, autocomplete: "off", ...extra }));
  /** @param {string} id @param {string} label */
  const check = (id, label) => el(doc, "label", { class: "check-label", for: `sp-${id}` }, el(doc, "input", { type: "checkbox", id: `sp-${id}` }), el(doc, "span", { text: ` ${label}` }));
  /** @param {string} id @param {string} label @param {[string, string][]} options */
  const select = (id, label, options) =>
    el(doc, "div", { class: "field" }, el(doc, "label", { class: "field-label", for: `sp-${id}`, text: label }), el(doc, "select", { id: `sp-${id}` }, ...options.map(([v, t]) => el(doc, "option", { value: v, text: t }))));
  /** @param {string} text @param {() => unknown} run @param {string} id */
  const button = (text, run, id) => {
    const b = el(doc, "button", { type: "button", id, text });
    b.addEventListener("click", () => void run());
    return b;
  };

  return openDialog(doc, {
    title: "Styles",
    intro: "Change how a kind of paragraph looks, and every paragraph in it follows, in the guide and in the Word file made from it. Only what you change is stored.",
    confirm: "Close",
    cancel: "",
    finish: () => true,
    first: () => doc.getElementById("sp-list"),
    body: () => {
      const datalist = el(doc, "datalist", { id: "sp-fonts" }, ...FONTS.map((f) => el(doc, "option", { value: f })));
      const left = el(
        doc,
        "div",
        { class: "sp-left" },
        el(doc, "label", { class: "field-label", for: "sp-list", text: "Style" }),
        el(doc, "select", { id: "sp-list", size: "14", "aria-label": "Styles of the guide" }),
        el(doc, "div", { class: "sp-buttons" }, button("New style...", create, "sp-new"), button("Delete style", remove, "sp-delete")),
      );
      const fontField = input("font", "Font", "text", { list: "sp-fonts", maxlength: "64" });
      const colorRow = el(doc, "div", { class: "field" }, el(doc, "label", { class: "field-label", for: "sp-color", text: "Color (six digits after a #)" }), el(doc, "div", { class: "sp-pair" }, el(doc, "input", { id: "sp-color", type: "text", autocomplete: "off", maxlength: "7", placeholder: "#1F3A5F" }), el(doc, "input", { id: "sp-colorpick", type: "color", "aria-label": "Pick a color" })));
      const shadeRow = el(doc, "div", { class: "field" }, el(doc, "label", { class: "field-label", for: "sp-shading", text: "Shading behind the paragraph" }), el(doc, "div", { class: "sp-pair" }, el(doc, "input", { id: "sp-shading", type: "text", autocomplete: "off", maxlength: "7", placeholder: "#FFF4CE" }), el(doc, "input", { id: "sp-shadingpick", type: "color", "aria-label": "Pick a shading color" })));
      const right = el(
        doc,
        "div",
        { class: "sp-right" },
        el(doc, "h3", { id: "sp-name" }),
        el(doc, "p", { id: "sp-based", class: "hint" }),
        el(doc, "div", { class: "sp-grid" }, fontField, input("size", "Size (points)", "text", { inputmode: "decimal", maxlength: "5" }), colorRow),
        el(doc, "div", { class: "sp-checks" }, check("bold", "Bold"), check("italic", "Italic"), check("underline", "Underline"), check("strike", "Strikethrough")),
        el(doc, "div", { class: "sp-grid" }, select("alignment", "Alignment", [["left", "Left"], ["center", "Centered"], ["right", "Right"], ["justify", "Justified"]]), input("before", "Space before (points)", "text", { inputmode: "decimal", maxlength: "5" }), input("after", "Space after (points)", "text", { inputmode: "decimal", maxlength: "5" })),
        el(doc, "div", { class: "sp-grid" }, select("line", "Line spacing", [["", "As the style"], ["1", "Single"], ["1.15", "1.15"], ["1.5", "1.5"], ["2", "Double"]]), input("left", "Indent from the left (points)", "text", { inputmode: "decimal", maxlength: "5" }), input("first", "First line (points)", "text", { inputmode: "decimal", maxlength: "5" })),
        el(doc, "div", { class: "sp-grid" }, shadeRow, el(doc, "div", { class: "sp-checks" }, check("keep", "Keep with the next paragraph"))),
        el(doc, "p", { id: "sp-sample", class: "sp-sample", text: "The quick brown fox jumps over the lazy dog. Handouts, cues and notes all start from a style like this." }),
        el(doc, "p", { id: "sp-contrast", class: "sp-contrast", role: "status" }),
        el(doc, "p", { id: "sp-problem", class: "problem", role: "alert", hidden: "" }),
        el(doc, "div", { class: "sp-buttons" }, button("Apply to this style", apply, "sp-apply"), button("Put back as the program makes it", reset, "sp-reset")),
      );
      const panel = el(doc, "div", { class: "styles-panel" }, datalist, left, right);
      // Everything wired once the panel is in the dialog (the dialog adds the body before it shows it).
      queueMicrotask(() => {
        const list = /** @type {HTMLSelectElement} */ (doc.getElementById("sp-list"));
        list.addEventListener("change", () => choose(list.value));
        panel.addEventListener("input", (e) => {
          const t = /** @type {HTMLInputElement} */ (e.target);
          if (t.id === "sp-colorpick") field("color").value = t.value.toUpperCase();
          if (t.id === "sp-shadingpick") field("shading").value = t.value.toUpperCase();
          refreshSample();
        });
        reload(startId);
      });
      return [panel];
    },
  });
}
