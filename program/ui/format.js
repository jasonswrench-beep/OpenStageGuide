// @ts-check
/**
 * The dialogs of the Format menu. Each asks for what it needs and gives the result back; the editor applies it. Text and attributes only.
 * The dialog checks what was typed with the same code that will apply it (`validate`), so a problem is explained in the dialog, which stays
 * open, and nothing is changed until the settings are good.
 */
import { fontChoices } from "../../src/guide-fonts.mjs";
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

/**
 * How the chosen words look. A box left as it is changes nothing; a look that was switched is given back as on or off, and what "off" means (the
 * style may have it on) is for the editor to say. Resolves to the changes (possibly none), or null when the dialog was closed.
 * @param {Document} doc
 * @param {{ font?: string, size?: number, color?: string, bold?: boolean, italic?: boolean, underline?: boolean, strike?: boolean, script?: string }} current what the words are like now
 * @param {(change: Record<string, unknown>) => void} validate throws a message for settings that cannot be used
 * @returns {Promise<Record<string, unknown> | null>}
 */
export function askTextFormat(doc, current, validate) {
  const was = (/** @type {unknown} */ v) => (v === undefined || v === null ? "" : String(v));
  return askFields(doc, {
    title: "Text",
    intro: "How the chosen words look (with none chosen, every word of the paragraph). A box left as it is changes nothing. Underlined words look like links on a screen, so bold or italic is clearer for emphasis.",
    confirm: "Apply",
    fields: [
      { id: "font", label: "Font", type: "text", value: was(current.font), list: fontChoices(), hint: "Any font name. The fonts that come with the program are carried inside the Word file; any other is only named, so the computer that opens the file needs it." },
      { id: "size", label: "Size (points, 6 to 96)", type: "text", value: was(current.size) },
      { id: "color", label: "Color (six digits after a #)", type: "text", value: was(current.color), hint: "Dark words on a light page read best. Review guide checks the contrast." },
      { id: "bold", label: "Bold", type: "check", value: current.bold === true },
      { id: "italic", label: "Italic", type: "check", value: current.italic === true },
      { id: "underline", label: "Underline", type: "check", value: current.underline === true },
      { id: "strike", label: "Strikethrough", type: "check", value: current.strike === true },
      {
        id: "script",
        label: "Raised or lowered",
        type: "select",
        value: was(current.script),
        options: [
          { value: "", label: "As it is" },
          { value: "superscript", label: "Raised (superscript)" },
          { value: "subscript", label: "Lowered (subscript)" },
          { value: "baseline", label: "Neither" },
        ],
      },
    ],
    finish: (v) => {
      /** @type {Record<string, unknown>} */
      const change = {};
      const font = String(v.font).trim();
      if (font && font !== was(current.font)) change.font = font;
      const size = String(v.size).trim();
      if (size && size !== was(current.size)) change.size = Number(size);
      const color = String(v.color).trim().toUpperCase();
      if (color && color !== was(current.color).toUpperCase()) change.color = color;
      for (const k of ["bold", "italic", "underline", "strike"]) if (v[k] !== (/** @type {Record<string, unknown>} */ (current)[k] === true)) change[k] = v[k] === true;
      const script = String(v.script);
      if (script && script !== was(current.script)) change.script = script;
      validate(change);
      return change;
    },
  });
}

