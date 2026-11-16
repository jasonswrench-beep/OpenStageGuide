// @ts-check
/**
 * The Insert menu: forms from templates, and the elements a form is made of (lines to write on, a box, a row of fields, a check list, a table,
 * a rating grid, cards, a source line). Each command asks for what the element needs in a dialog and gives back a function that builds it
 * with the same builders the rest of the program uses (src/guide-forms-ooxml.mjs). The element is checked once, as the dialog's main button
 * is chosen, so a problem is explained in the dialog and the dialog stays open.
 */
import { makeContext, para, pictureRun, run } from "../../src/guide-ooxml.mjs";
import { PICTURE_HELP, PICTURE_WIDTHS, pictureSize, prepareImage, toBase64 } from "../../src/image-prepare.mjs";
import { STYLE_IDS } from "../../src/guide-styles.mjs";
import { writingLines, writingBox, fieldRow, checkList, formTable, ratingGrid, cardSet, sourceLine } from "../../src/guide-forms-ooxml.mjs";
import { FORM_TEMPLATES, buildForm } from "../../src/guide-forms-templates.mjs";
import { LABELS, newFormId } from "../../src/guide-forms.mjs";
import { blockChoices, buildChoice, choiceById } from "../../src/guide-block-choice.mjs";
import { askFields } from "./dialogs.js";

/** @typedef {{ build: (ctx: any) => string, what: string, edit?: boolean }} Made */

/** The kind of block chosen last, so adding several of the same kind in a row is one Enter each. */
let lastBlockKind = "say";

/** @param {string} text */
const lines = (text) =>
  text
    .split(/\r?\n/)
    .map((s) => s.trim())
    .filter(Boolean);
/** @param {string} text */
const list = (text) =>
  text
    .split(/[,\n]/)
    .map((s) => s.trim())
    .filter(Boolean);
/** @param {string | boolean} v */
const num = (v) => Number(String(v).trim());

/**
 * Make sure what was asked for can be built, once, so the dialog can explain a problem and stay open. Gives the builder back.
 * @param {(ctx: any) => string} build @param {string} what @returns {Made}
 */
function checked(build, what) {
  build(makeContext());
  return { build, what };
}

/** @type {Record<string, (doc: Document) => Promise<Made | null>>} */
export const INSERTS = {
  block: (doc) =>
    askFields(doc, {
      title: "Add a block here",
      intro: "The block goes in after the block you are on, with its icon and its label. Its words are typed in the guide itself.",
      confirm: "Add block",
      fields: [
        {
          id: "kind",
          label: "Kind of block",
          type: "select",
          value: choiceById(lastBlockKind) ? lastBlockKind : "say",
          options: blockChoices().map((c) => ({ value: c.id, label: `${c.groupLabel}: ${c.label}` })),
          hint: choiceById(lastBlockKind)?.help ?? "",
        },
        { id: "minutes", label: "Minutes (optional)", type: "number", hint: "Written after the words as (N min); a lesson's time line adds them up." },
      ],
      // The hint under the choice says what the chosen kind is for.
      wire: (inputs) => {
        const kind = /** @type {HTMLSelectElement} */ (inputs.get("kind"));
        const say = () => {
          const hint = kind.parentElement?.querySelector(".hint");
          if (hint) hint.textContent = choiceById(kind.value)?.help ?? "";
        };
        kind.addEventListener("change", say);
      },
      finish: (v) => {
        const id = String(v.kind);
        const build = buildChoice(id, { minutes: String(v.minutes).trim() });
        build(makeContext());
        lastBlockKind = id;
        return { build, what: `${choiceById(id)?.label ?? "Block"} block`, edit: true };
      },
    }),
  lines: (doc) =>
    askFields(doc, {
      title: "Lines to write on",
      intro: "Ruled lines for a written answer. A label above them is optional.",
      confirm: "Insert",
      fields: [
        { id: "label", label: "Label above the lines (optional)", type: "text" },
        { id: "count", label: "Number of lines (1 to 12)", type: "number", value: "4" },
      ],
      finish: (v) => checked((_ctx) => writingLines({ label: String(v.label), count: num(v.count) }), "Lines"),
    }),
  box: (doc) =>
    askFields(doc, {
      title: "Writing box",
      intro: "An empty box for a drawing or a longer answer.",
      confirm: "Insert",
      fields: [
        { id: "label", label: "Label above the box (optional)", type: "text" },
        { id: "height", label: "Height in lines (1 to 12)", type: "number", value: "5" },
      ],
      finish: (v) => checked((_ctx) => writingBox({ label: String(v.label), lines: num(v.height) }), "Box"),
    }),
  fields: (doc) =>
    askFields(doc, {
      title: "Row of fields",
      intro: "Labels with a line after each to write on, side by side (up to four).",
      confirm: "Insert",
      fields: [
        { id: "labels", label: "Labels, one to a line", type: "lines", value: "Name\nDate", rows: 4, hint: "For example: Name, Date, Signature, Organization." },
        {
          id: "width",
          label: "Width of each",
          type: "select",
          value: "medium",
          options: [
            { value: "short", label: "Short" },
            { value: "medium", label: "Medium" },
            { value: "full", label: "Full" },
          ],
        },
      ],
      finish: (v) => checked((_ctx) => fieldRow(lines(String(v.labels)).map((label) => ({ label, width: String(v.width) }))), "Row of fields"),
    }),
  checklist: (doc) =>
    askFields(doc, {
      title: "Check list",
      intro: "Items with a check box each, and up to three more columns to fill in.",
      confirm: "Insert",
      fields: [
        { id: "items", label: "Items, one to a line", type: "lines", rows: 6 },
        { id: "columns", label: "Extra columns (optional)", type: "text", value: "Who, When", hint: "Separate the names with commas. Up to three." },
      ],
      finish: (v) => checked((ctx) => checkList({ items: lines(String(v.items)), columns: list(String(v.columns)), ctx }), "Check list"),
    }),
  table: (doc) =>
    askFields(doc, {
      title: "Table to fill in",
      intro: "A header row naming each column, and empty rows to write in.",
      confirm: "Insert",
      fields: [
        { id: "columns", label: "Column names", type: "text", value: "Topic, Notes", hint: "Separate the names with commas. Up to eight." },
        { id: "rows", label: "Empty rows (0 to 40)", type: "number", value: "5" },
        { id: "caption", label: "What the table is (for a screen reader)", type: "text", value: "Table" },
      ],
      finish: (v) =>
        checked(
          (ctx) => formTable({ columns: list(String(v.columns)).map((name) => ({ name, weight: 1 })), blankRows: num(v.rows), ctx, caption: String(v.caption) || "Table" }),
          "Table",
        ),
    }),
  rating: (doc) =>
    askFields(doc, {
      title: "Rating grid",
      intro: "Statements to rate on a scale, with a check box under each point.",
      confirm: "Insert",
      fields: [
        { id: "statements", label: "Statements, one to a line", type: "lines", rows: 6 },
        { id: "scale", label: "Points on the scale", type: "text", value: "1, 2, 3, 4, 5", hint: "From two to seven, separated by commas." },
        { id: "evidence", label: "Add a column to write the evidence in", type: "check", value: false },
      ],
      finish: (v) => checked((ctx) => ratingGrid({ statements: lines(String(v.statements)), labels: list(String(v.scale)), evidence: v.evidence === true, ctx }), "Rating grid"),
    }),
  cards: (doc) =>
    askFields(doc, {
      title: "Cards to cut apart",
      intro: "Each card is a paragraph of words. Separate the cards with a blank line.",
      confirm: "Insert",
      fields: [
        { id: "cards", label: "Cards", type: "lines", rows: 8, value: "First card\nWords for the card.\n\nSecond card\nWords for the card." },
        {
          id: "perPage",
          label: "Cards to a page",
          type: "select",
          value: "4",
          options: [
            { value: "2", label: "2" },
            { value: "4", label: "4" },
            { value: "6", label: "6" },
            { value: "8", label: "8" },
          ],
        },
      ],
      finish: (v) =>
        checked(
          (_ctx) =>
            cardSet({
              cards: String(v.cards)
                .split(/\r?\n[ \t]*\r?\n/)
                .map((c) => c.trim())
                .filter(Boolean),
              perPage: num(v.perPage),
            }),
          "Cards",
        ),
    }),
  picture: (doc) =>
    askFields(doc, {
      title: "Picture",
      intro: `${PICTURE_HELP} A picture needs a description for people who cannot see it, or a tick to say it is only decoration.`,
      confirm: "Insert",
      fields: [
        { id: "file", label: "Picture file", type: "file", accept: "image/png,image/jpeg,image/gif" },
        { id: "alt", label: "Describe the picture in a few words", type: "text", hint: "For example: A facilitator at a flip chart, pointing to three columns." },
        { id: "decorative", label: "This picture is only decoration", type: "check", value: false },
        { id: "width", label: "Width", type: "select", value: "large", options: PICTURE_WIDTHS.map((w) => ({ value: w.id, label: w.label })) },
        { id: "caption", label: "Caption under the picture (optional)", type: "text" },
      ],
      finish: async (v) => {
        if (!(v.file instanceof File)) throw new Error(`Choose a picture file first. ${PICTURE_HELP}`);
        const info = await prepareImage(v.file);
        const picture = { base64: toBase64(info.bytes), type: info.type, mime: info.mime, size: pictureSize(info, String(v.width)), alt: String(v.alt), decorative: v.decorative === true, name: info.name };
        const caption = String(v.caption).trim();
        return checked((ctx) => para(STYLE_IDS.body, pictureRun(picture, ctx)) + (caption ? para(null, run(caption)) : ""), "Picture");
      },
    }),
  source: (doc) =>
    askFields(doc, {
      title: "Source line",
      intro: "A short credit or permission line, under a form.",
      confirm: "Insert",
      fields: [{ id: "text", label: "The line", type: "text", value: "" }],
      finish: (v) => checked((_ctx) => sourceLine(String(v.text)), "Source line"),
    }),
};

/**
 * A whole form from a template. It needs the guide's paragraphs to give the form an id nothing else has.
 * @param {Document} doc @param {{ paragraphs: () => Promise<any[]> }} guide @returns {Promise<Made | null>}
 */
export async function newForm(doc, guide) {
  const paras = await guide.paragraphs();
  const made = await askFields(doc, {
    title: "New form",
    intro: "A start to write over. The words in it say what to type where. Update guide numbers the form.",
    confirm: "Insert",
    // The planning forms are for the facilitator whatever is ticked (src/guide-forms-templates.mjs): the box says so, ticked and dimmed, while one of them is chosen,
    // and goes back to what the author had chosen when another kind is.
    wire: (inputs) => {
      const select = /** @type {HTMLSelectElement} */ (inputs.get("template"));
      const box = /** @type {HTMLInputElement} */ (inputs.get("facilitator"));
      const words = box.closest("label")?.querySelector("span");
      const plain = words?.textContent ?? "";
      let own = box.checked;
      box.addEventListener("change", () => {
        if (!box.disabled) own = box.checked;
      });
      const follow = () => {
        const always = FORM_TEMPLATES.find((t) => t.id === select.value)?.facilitator === true;
        box.disabled = always;
        box.checked = always || own;
        if (words) words.textContent = always ? `${plain} Always, for this kind of form.` : plain;
      };
      select.addEventListener("change", follow);
      follow();
    },
    fields: [
      {
        id: "template",
        label: "What kind of form",
        type: "select",
        value: FORM_TEMPLATES[0].id,
        options: FORM_TEMPLATES.map((t) => ({ value: t.id, label: `${t.label}: ${t.note}` })),
      },
      { id: "title", label: "Title of the form", type: "text" },
      { id: "kind", label: "Called a", type: "select", value: "", options: [{ value: "", label: "Whatever suits the form" }, ...LABELS.map((l) => ({ value: l, label: l }))] },
      { id: "facilitator", label: "For the facilitator only (left out of the participant workbook)", type: "check", value: false },
    ],
    finish: (v) => {
      const id = newFormId(paras);
      const spec = { template: String(v.template), id, label: String(v.kind) || undefined, title: String(v.title), facilitator: v.facilitator === true };
      return checked((ctx) => buildForm(spec, ctx), "Form");
    },
  });
  return made;
}
