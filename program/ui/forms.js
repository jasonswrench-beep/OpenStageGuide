// @ts-check
/**
 * The Insert menu: forms from templates, and the elements a form is made of (lines to write on, a box, a row of fields, a check list, a table,
 * a rating grid, cards, a source line). Each command asks for what the element needs in a dialog and gives back a function that builds it
 * with the same builders the rest of the program uses (src/guide-forms-ooxml.mjs). The element is checked once, as the dialog's main button
 * is chosen, so a problem is explained in the dialog and the dialog stays open.
 */
import { makeContext } from "../../src/guide-ooxml.mjs";
import { writingLines, writingBox, fieldRow, checkList, formTable, ratingGrid, cardSet, sourceLine } from "../../src/guide-forms-ooxml.mjs";
import { FORM_TEMPLATES, buildForm } from "../../src/guide-forms-templates.mjs";
import { LABELS, newFormId } from "../../src/guide-forms.mjs";
import { askFields } from "./dialogs.js";

/** @typedef {{ build: (ctx: any) => string, what: string }} Made */

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
