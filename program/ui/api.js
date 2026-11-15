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

/**
 * A file the server makes (a PDF), as bytes, with what it said about it in its headers.
 * @param {string} path @param {Record<string, unknown>} body
 * @returns {Promise<{ bytes: Uint8Array, header: (name: string) => string | null }>}
 */
export async function postForFile(path, body) {
  const r = await fetch(`/api/${path}`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body), cache: "no-store" });
  if (!r.ok) {
    let message = `The program answered ${r.status}`;
    try {
      message = (await r.json()).error || message;
    } catch {
      /* the answer was not words: the status stays */
    }
    throw new Error(message);
  }
  return { bytes: new Uint8Array(await r.arrayBuffer()), header: (name) => r.headers.get(name) };
}

/** @param {string} path @param {Record<string, unknown>} [body] */
export async function post(path, body = {}) {
  return read(await fetch(`/api/${path}`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body), cache: "no-store" }));
}
