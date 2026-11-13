// @ts-check
// What the program asks its own server (scripts/project-server.mjs): which guides there are, make one, open one, close the open one. The
// guide itself is read and stored on /project (see src/guide-project-web.mjs). Nothing here knows a file's folder or full name.

/** @param {Response} r */
async function read(r) {
  const text = await r.text();
  /** @type {any} */
  let data = null;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    data = null;
  }
  if (!r.ok) throw new Error((data && data.error) || text || `The program answered ${r.status}`);
  return data;
}

/** @param {string} path */
export async function get(path) {
  return read(await fetch(`/api/${path}`, { cache: "no-store" }));
}

/** @param {string} path @param {Record<string, unknown>} [body] */
export async function post(path, body = {}) {
  return read(await fetch(`/api/${path}`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body), cache: "no-store" }));
}
