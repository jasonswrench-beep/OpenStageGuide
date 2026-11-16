// @ts-check
/**
 * File > Save for PrismStatistics: the guide's rating grids and self-scored instruments as a .pstat file (src/guide-pstat.mjs), so the answers, once they are
 * typed in or pasted, are ready to analyse. The author chooses which of the instruments to send and how many people there are to enter answers for; the
 * file goes to the downloads like the Word file does, and PrismStatistics opens it. Nothing is sent anywhere by the program.
 */
import { readInstruments, pstatProject, pstatBytes, pstatFileName, MAX_ROWS } from "../../src/guide-pstat.mjs";
import { readSettings } from "../../src/guide-forms.mjs";
import { askFields } from "./dialogs.js";
import { downloadFile, zipLibrary } from "./word-files.js";

const io = {
  parse: (/** @type {string} */ xml) => new DOMParser().parseFromString(xml, "application/xml"),
  serialize: (/** @type {Node} */ node) => new XMLSerializer().serializeToString(node),
};

/**
 * @param {Document} doc @param {any} guide the guide's document
 * @param {{ name: string, version?: string, say: (message: string, bad?: boolean) => void }} options
 */
export async function saveForStatistics(doc, guide, { name, version = "", say }) {
  try {
    const state = guide.exportState();
    const found = readInstruments(state.paragraphs, io);
    if (!found.length) {
      say("This guide has no rating grid or self-scored instrument yet. Insert > Form has a rating form and a self-scored instrument, and Insert > Rating grid makes a grid by itself.", true);
      return;
    }
    // How many people there will be to enter answers for: what the Forms tool was told, when it was told.
    const people = Number(readSettings(await guide.formSettings?.()).group?.people) || 0;
    const chosen = await askFields(doc, {
      title: "Save for PrismStatistics",
      intro: "Each statement becomes a variable, with its words as its label and its scale's points as value labels; the answers are empty. When they are typed in or pasted, they are ready to analyse. How each scale is added up, and what each range of totals means, is written in the notes the file carries.",
      confirm: "Save file",
      fields: [
        ...found.map((f, i) => ({
          id: `i${i}`,
          label: `${f.title}${f.formTitle && f.formTitle !== f.title ? ` (in ${f.formTitle})` : ""}: ${f.statements.length} ${f.statements.length === 1 ? "statement" : "statements"}${f.kind === "instrument" ? ", self-scored" : ""}`,
          type: /** @type {const} */ ("check"),
          value: true,
        })),
        { id: "rows", label: `People to enter answers for (0 to ${MAX_ROWS})`, type: /** @type {const} */ ("number"), value: String(Math.min(people, MAX_ROWS) || 30), hint: "One empty row for each. Rows can be added in PrismStatistics later." },
      ],
      finish: (v) => {
        const picked = found.filter((_, i) => v[`i${i}`] === true);
        if (!picked.length) throw new Error("Choose at least one to save.");
        const rows = String(v.rows).trim() === "" ? 0 : Number(String(v.rows).trim());
        return pstatProject(picked, { guideTitle: String(state.title || name), rows, appVersion: `OpenStage Guide${version ? ` ${version}` : ""}` });
      },
    });
    if (!chosen) return;
    const bytes = await pstatBytes(chosen, zipLibrary());
    const file = pstatFileName(name);
    downloadFile(doc, bytes, file, "application/octet-stream");
    const statements = chosen.dataset.variables.length - 1;
    say(`${file} was saved: ${statements} ${statements === 1 ? "statement" : "statements"} as variables, ${chosen.source.report.rows} empty ${chosen.source.report.rows === 1 ? "row" : "rows"}. Open it in PrismStatistics; its notes say how each scale is added up.`);
  } catch (e) {
    say(/** @type {Error} */ (e).message, true);
  }
}
