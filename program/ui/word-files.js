// @ts-check
/**
 * The guide and Word files: the guide made into a Word document, its styles made into a Word template, and the styles of a Word file taken
 * into the guide. The files are made and read in the page (src/guide-docx.mjs and src/guide-styles-import.mjs do the work; the zip library is
 * the add-in's own copy, loaded by the page). A file the author chooses is read, never opened or run, and is not changed. Text and attributes
 * only, never markup.
 */
import { DOCX_LIMITS, docxFromPackage, readDocxStyles } from "../../src/guide-docx.mjs";
import { FONT_URL, STYLES as FONT_STYLES, packFont } from "../../src/guide-fonts.mjs";
import { listStyles, niceName, specCss } from "../../src/guide-styles-edit.mjs";
import { AS_NEW, candidates, importSummary, planImport } from "../../src/guide-styles-import.mjs";
import { askFields, openDialog } from "./dialogs.js";
import { el } from "./menubar.js";
import { drawCss } from "./preview.js";

const parseXml = (/** @type {string} */ xml) => new DOMParser().parseFromString(xml, "application/xml");
const io = { parse: parseXml, serialize: (/** @type {Element} */ node) => new XMLSerializer().serializeToString(node) };
const zipLibrary = () => {
  const lib = /** @type {any} */ (window).JSZip;
  if (!lib) throw new Error("The part of the program that makes Word files did not load. Close the program and open it again.");
  return lib;
};

const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const DOTX = "application/vnd.openxmlformats-officedocument.wordprocessingml.template";
let groupCounter = 0;

/** A file name Windows accepts, from the name of a guide. @param {string} name */
export function safeFileName(name) {
  const clean = String(name ?? "").replace(/[\\/:*?"<>|\u0000-\u001f]/g, "-").replace(/\s+/g, " ").trim().replace(/[. ]+$/, "");
  return (clean || "guide").slice(0, 100);
}

/** Give the author a file: the browser's own download. @param {Document} doc @param {Uint8Array} bytes @param {string} fileName @param {string} type */
export function downloadFile(doc, bytes, fileName, type) {
  const url = URL.createObjectURL(new Blob([/** @type {BlobPart} */ (bytes)], { type }));
  const link = el(doc, "a", { href: url, download: fileName, hidden: "" });
  doc.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 30000);
}

/** The files of the program's own fonts, from the program's server. @param {string[]} families */
async function loadPackFonts(families) {
  /** @type {Record<string, Record<string, Uint8Array>>} */
  const out = {};
  for (const family of families) {
    const font = packFont(family);
    if (!font) continue;
    out[family] = {};
    for (const style of FONT_STYLES) {
      const file = font.files[style];
      if (!file) continue;
      const res = await fetch(`${FONT_URL}${file}`);
      if (!res.ok) throw new Error(`The font ${family} could not be read, so it was not put in the file.`);
      out[family][style] = new Uint8Array(await res.arrayBuffer());
    }
  }
  return out;
}

/**
 * Make the guide a Word document (or, with `template`, a Word template that holds only its styles) and give it to the author.
 * @param {Document} doc @param {any} guide the guide's document @param {{ name: string, template: boolean, say: (message: string, bad?: boolean) => void }} opts
 */
export async function saveWordFile(doc, guide, { name, template, say }) {
  try {
    const state = guide.exportState();
    const pkg = await guide.sourcePackage();
    const title = String(state.title || name);
    /** @type {{ family: string }[]} */
    let carried = [];
    const options = { ...(template ? { template: true } : { entries: state.paragraphs }), title, fonts: loadPackFonts, onEmbedded: (/** @type {{ family: string }[]} */ list) => (carried = list) };
    const bytes = await docxFromPackage(pkg, zipLibrary(), io, options);
    const file = `${safeFileName(name)}${template ? " styles.dotx" : ".docx"}`;
    downloadFile(doc, bytes, file, template ? DOTX : DOCX);
    const names = carried.map((f) => f.family);
    const fonts = names.length ? ` ${names.length === 1 ? `The font ${names[0]} is` : `The fonts ${names.slice(0, -1).join(", ")} and ${names[names.length - 1]} are`} inside the file, so it looks right on a computer that does not have ${names.length === 1 ? "it" : "them"}.` : "";
    say(
      template
        ? `${file} has the guide's styles. In Word, start a document from it, or copy it to your templates folder. The styles are real Word styles, so they appear in Word's Styles gallery.${fonts}`
        : `${file} was made from the guide. Open it in Word, press Ctrl+A and then F9 to fill in the page numbers of the contents.${fonts}`,
    );
  } catch (e) {
    say(/** @type {Error} */ (e).message, true);
  }
}

/** A line saying how a style looks, for a person who cannot see the sample. @param {any} spec */
function describe(spec) {
  const f = spec.font ?? {};
  const p = spec.paragraph ?? {};
  const parts = [f.name ?? "the usual font", f.size ? `${f.size} pt` : "", f.bold ? "bold" : "", f.italic ? "italic" : "", f.color ? `color ${f.color}` : "", p.alignment && p.alignment !== "left" ? `${p.alignment} aligned` : "", spec.shading ? `shaded ${spec.shading}` : ""];
  return parts.filter(Boolean).join(", ");
}

/**
 * Choose a Word file and take its styles into the guide: the styles of the same name change the guide's, and the rest can be added as the
 * author's own. One Undo takes it all back.
 * @param {Document} doc @param {any} guide @param {(message: string, bad?: boolean) => void} say
 */
export async function importStylesFromWord(doc, guide, say) {
  const read = await askFields(doc, {
    title: "Styles from a Word file",
    intro: "Choose a Word document or template. Its styles can change the guide's styles of the same name, or be added as styles of your own. The file is only read: it is not opened in Word, run or changed.",
    confirm: "Read the styles",
    fields: [{ id: "file", label: "Word file (.docx or .dotx)", type: "file", accept: ".docx,.dotx" }],
    finish: async (values) => {
      const file = /** @type {File | null} */ (values.file);
      if (!file) throw new Error("Choose a Word file first.");
      if (file.size > DOCX_LIMITS.fileBytes) throw new Error(`Choose a Word file under ${DOCX_LIMITS.fileBytes / 1048576} MB.`);
      return readDocxStyles(new Uint8Array(await file.arrayBuffer()), zipLibrary(), io);
    },
  });
  if (!read) return;
  const guideStyles = guide.listStyles().filter((/** @type {any} */ s) => s.type === "paragraph");
  const fileStyles = listStyles(read.stylesXml, io, new Set(), { themeFonts: read.themeFonts, paragraphDefaults: true });
  const offered = candidates(fileStyles, guideStyles);
  if (!offered.length) {
    say("That file has no paragraph styles to bring in.", true);
    return;
  }
  const plan = await chooseStyles(doc, guide, offered, fileStyles, guideStyles);
  if (!plan) return;
  const adjusted = plan.notes.length ? ` ${plan.notes.length === 1 ? "One setting" : `${plan.notes.length} settings`} had to be adjusted to fit the guide: ${plan.notes[0]}` : "";
  say(`${importSummary(plan)}${adjusted}`);
}

/**
 * The styles of the file, each with whether to bring it in and which style of the guide it becomes. Resolves to what was done, or null.
 * @param {Document} doc @param {any} guide @param {ReturnType<typeof candidates>} offered @param {any[]} fileStyles @param {any[]} guideStyles
 */
function chooseStyles(doc, guide, offered, fileStyles, guideStyles) {
  const matched = offered.filter((c) => c.to).length;
  const targets = [...guideStyles].sort((a, b) => niceName(a.name).localeCompare(niceName(b.name)));
  const rows = offered.map((c, i) => {
    const check = el(doc, "input", { type: "checkbox", class: "si-pick", id: `si-pick-${i}`, "data-file": c.fileId, ...(c.to ? { checked: "" } : {}) });
    const sample = el(doc, "span", { class: "si-sample", "aria-hidden": "true", text: "The quick brown fox" });
    /** @type {Record<string, string>} */
    const css = { ...specCss(c.spec) };
    for (const k of ["marginTop", "marginBottom", "marginLeft", "textIndent"]) delete css[k];
    // Whatever size the style has, the sample stays a size that fits the row.
    if (c.spec.font?.size) css.fontSize = `${Math.min(c.spec.font.size, 20) * 1.3333}px`;
    drawCss(sample, css);
    const to = el(doc, "select", { class: "si-to", id: `si-to-${i}`, "aria-label": `${niceName(c.name)} becomes` }, el(doc, "option", { value: AS_NEW, text: "A new style of your own" }), ...targets.map((t) => el(doc, "option", { value: t.id, text: `The guide's ${niceName(t.name)}` })));
    to.value = c.to ?? AS_NEW;
    return el(
      doc,
      "tr",
      {},
      el(doc, "th", { scope: "row" }, el(doc, "label", { for: `si-pick-${i}` }, check, el(doc, "span", { text: ` ${niceName(c.name)}` }))),
      el(doc, "td", {}, sample, el(doc, "span", { class: "hint si-desc", text: describe(c.spec) })),
      el(doc, "td", {}, to),
    );
  });
  const intro = matched
    ? `${matched === 1 ? "One style" : `${matched} styles`} in this file ${matched === 1 ? "has" : "have"} the name of a style of the guide, and ${matched === 1 ? "is" : "are"} chosen to change it. For any other, choose whether to add it as a style of your own.`
    : "No style in this file has the name of a style of the guide. Choose the ones to add as styles of your own, or choose which style of the guide each one changes.";
  return openDialog(doc, {
    title: "Bring in styles",
    intro,
    confirm: "Bring in",
    first: (d) => d.querySelector("input.si-pick"),
    body: (d) => {
      const all = el(doc, "button", { type: "button", text: "Choose all" });
      const none = el(doc, "button", { type: "button", text: "Choose none" });
      const set = (/** @type {boolean} */ on) => d.querySelectorAll("input.si-pick").forEach((c) => (/** @type {HTMLInputElement} */ (c).checked = on));
      all.addEventListener("click", () => set(true));
      none.addEventListener("click", () => set(false));
      const table = el(
        doc,
        "table",
        { class: "si-table" },
        el(doc, "thead", {}, el(doc, "tr", {}, el(doc, "th", { scope: "col", text: "Style in the file" }), el(doc, "th", { scope: "col", text: "How it looks" }), el(doc, "th", { scope: "col", text: "Becomes" }))),
        el(doc, "tbody", {}, ...rows),
      );
      return [el(doc, "div", { class: "sp-buttons" }, all, none), el(doc, "div", { class: "si-wrap", role: "region", tabindex: "0", "aria-label": "Styles in the file" }, table)];
    },
    finish: async (d) => {
      const picks = [...d.querySelectorAll("input.si-pick:checked")].map((c) => {
        const box = /** @type {HTMLInputElement} */ (c);
        const to = /** @type {HTMLSelectElement} */ (d.querySelector(`#${box.id.replace("pick", "to")}`));
        return { fileId: String(box.dataset.file), to: to.value };
      });
      if (!picks.length) throw new Error("Choose at least one style to bring in.");
      const plan = planImport({ file: fileStyles, guide: guideStyles, edits: guide.styleEdits(), picks });
      if (JSON.stringify(plan.edits) !== JSON.stringify(guide.styleEdits())) await guide.setStyleEdits(plan.edits, { group: `import:${++groupCounter}` });
      return plan;
    },
  });
}
