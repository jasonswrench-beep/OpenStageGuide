// @ts-check
/**
 * The dialogs of the Format menu. Each asks for what it needs and gives the result back; the editor applies it. Text and attributes only.
 * The dialog checks what was typed with the same code that will apply it (`validate`), so a problem is explained in the dialog, which stays
 * open, and nothing is changed until the settings are good.
 */
import { askFields } from "./dialogs.js";

/**
 * @typedef {{ alignment: "left" | "center" | "right" | "justify" | null, spaceBefore: number | null, spaceAfter: number | null, lineSpacing: number | null, leftIndent: number | null, firstLineIndent: number | null }} ParagraphFormat
 */

/** @param {number | null} v */
const show = (v) => (v === null ? "" : String(v));
/** @param {string | boolean | File | null} v */
const num = (v) => (String(v).trim() === "" ? null : Number(String(v).trim()));

/**
 * How a paragraph, or each of the paragraphs chosen, looks on the page. Every box left empty (or "As the style") lets the style decide.
 * @param {Document} doc @param {ParagraphFormat} current @param {(fmt: ParagraphFormat) => void} validate throws a message for settings that cannot be used
 * @returns {Promise<ParagraphFormat | null>}
 */
export function askParagraphFormat(doc, current, validate) {
  const lines = [
    { value: "", label: "As the style" },
    { value: "1", label: "Single" },
    { value: "1.15", label: "1.15" },
    { value: "1.5", label: "1.5" },
    { value: "2", label: "Double" },
  ];
  const now = show(current.lineSpacing);
  if (now && !lines.some((l) => l.value === now)) lines.push({ value: now, label: now });
  return askFields(doc, {
    title: "Paragraph",
    intro: "How the paragraph looks on the page. Leave a box empty, or choose \"As the style\", to let its style decide. These settings go on every paragraph chosen.",
    confirm: "Apply",
    fields: [
      {
        id: "alignment",
        label: "Alignment",
        type: "select",
        value: current.alignment ?? "",
        options: [
          { value: "", label: "As the style" },
          { value: "left", label: "Left" },
          { value: "center", label: "Centered" },
          { value: "right", label: "Right" },
          { value: "justify", label: "Justified (the gaps between words vary, which is harder to read)" },
        ],
      },
      { id: "before", label: "Space before (points)", type: "number", value: show(current.spaceBefore) },
      { id: "after", label: "Space after (points)", type: "number", value: show(current.spaceAfter) },
      { id: "line", label: "Line spacing", type: "select", value: now, options: lines },
      { id: "left", label: "Indent from the left (points)", type: "number", value: show(current.leftIndent) },
      { id: "first", label: "First line (points)", type: "number", value: show(current.firstLineIndent), hint: "A negative number hangs the first line out to the left." },
    ],
    finish: (v) => {
      /** @type {ParagraphFormat} */
      const fmt = {
        alignment: /** @type {ParagraphFormat["alignment"]} */ (String(v.alignment) || null),
        spaceBefore: num(v.before),
        spaceAfter: num(v.after),
        lineSpacing: num(v.line),
        leftIndent: num(v.left),
        firstLineIndent: num(v.first),
      };
      validate(fmt);
      return fmt;
    },
  });
}
