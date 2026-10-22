// The Teacher Portal (ibaclinton.com/teachers/): a hidden, password-protected page for IBA staff.
//
// Its content is not on the public site. The office edits it in the private repo TEACHER_REPO (content/teachers.json
// and files/); that repo's workflow runs teachers.py and commits built/portal.json: { html, forms, password } with
// the password as a hash. This service reads that file and the uploaded files from the private repo with
// TEACHER_REPO_TOKEN (read-only), and hands them to a browser only with a signed sign-in token:
//
//   POST /teachers/login   {password, remember}  -> {token}: valid 30 days with "remember", else 12 hours
//   GET  /teachers/portal  Authorization: Bearer <token> -> {html, files}; files is a 6-hour file pass
//   GET  /teachers/file/<file pass>/<name>  one uploaded file (a plain link, so the pass is in the address)
//   POST /teachers/form    Authorization: Bearer <token>: a Teacher Portal form, saved and emailed like /form
//
// Tokens are signed with SETUP_KEY and the password hash, so changing the password in the editor signs everyone out.

export const FILE_TYPES = {
  pdf: "application/pdf", doc: "application/msword",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xls: "application/vnd.ms-excel", xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ppt: "application/vnd.ms-powerpoint", pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png", gif: "image/gif", webp: "image/webp",
}; // teachers.py FILE_TYPES
export const LOGIN_DAYS = 30;
const SHORT_HOURS = 12, FILE_HOURS = 6;
const MAX_FAILS = 20; // wrong passwords per connection per hour (a whole school can share one connection)

// ------------------------------------------------------------------ the private repo
let cache = { at: 0, data: null };

export function resetCache() { cache = { at: 0, data: null }; }

export function gitHub(env, path, accept = "application/vnd.github.raw+json") {
  const url = `https://api.github.com/repos/${env.TEACHER_REPO}/contents/${path.split("/").map(encodeURIComponent).join("/")}?ref=main`;
  return fetch(url, { headers: { Authorization: `Bearer ${env.TEACHER_REPO_TOKEN}`, Accept: accept,
    "User-Agent": "ibaclinton-teachers", "X-GitHub-Api-Version": "2022-11-28" } });
}

// built/portal.json, kept for 30 seconds so an office save shows within a minute or so of its build.
export async function loadPortal(env, now = Date.now()) {
  if (!env.TEACHER_REPO || !env.TEACHER_REPO_TOKEN) return null;
  if (cache.data && now - cache.at < 30000) return cache.data;
  const r = await gitHub(env, "built/portal.json");
  if (!r.ok) throw new Error(`teacher portal ${r.status}`);
  cache = { at: now, data: await r.json() };
  return cache.data;
}

// ------------------------------------------------------------------ passwords and tokens
async function sha256(text) {
  const d = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return Array.from(new Uint8Array(d), (b) => b.toString(16).padStart(2, "0")).join("");
}

export function passwordHash(password) { // teachers.py password_hash
  const p = String(password == null ? "" : password).trim();
  return p ? sha256("ibaclinton-teachers|" + p) : Promise.resolve("");
}

async function hmac(env, portal, text) {
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(`${env.SETUP_KEY}|${portal.password}`),
    { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const sig = new Uint8Array(await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(text)));
  return btoa(String.fromCharCode(...sig)).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

export async function makeToken(env, portal, kind, hours, now = Date.now()) {
  const exp = Math.floor(now / 1000) + hours * 3600;
  return `${exp}.${await hmac(env, portal, `${kind}|${exp}`)}`;
}

export async function checkToken(env, portal, kind, token, now = Date.now()) {
  const m = /^(\d{9,11})\.([\w-]{43})$/.exec(String(token || ""));
  if (!m || !portal || !portal.password || !env.SETUP_KEY || +m[1] < now / 1000) return false;
  const want = await hmac(env, portal, `${kind}|${m[1]}`);
  let diff = 0; // same time whatever the mismatch
  for (let i = 0; i < want.length; i++) diff |= want.charCodeAt(i) ^ m[2].charCodeAt(i);
  return diff === 0;
}

function bearer(request) {
  return (/^Bearer\s+(\S+)$/i.exec(request.headers.get("Authorization") || "") || [])[1] || "";
}

// The portal for a signed-in request, or a Response to send back instead.
export async function signedIn(request, env, cors, json) {
  const portal = await loadPortal(env);
  if (!portal) return json({ ok: false, error: "The Teacher Portal isn't set up yet" }, 503, cors);
  if (!(await checkToken(env, portal, "tp", bearer(request)))) {
    return json({ ok: false, error: "Please sign in again", signin: true }, 401, cors);
  }
  return portal;
}

// ------------------------------------------------------------------ requests
export async function login(request, env, cors, json) {
  if (!cors["Access-Control-Allow-Origin"]) return json({ ok: false, error: "Not allowed" }, 403, cors);
  const text = await request.text();
  if (text.length > 2000) return json({ ok: false, error: "Too large" }, 413, cors);
  let body;
  try { body = JSON.parse(text); } catch { return json({ ok: false, error: "Bad request" }, 400, cors); }
  const ip = request.headers.get("CF-Connecting-IP") || "unknown";
  const fails = parseInt((await env.LUNCH_KV.get(`rltp:${ip}`)) || "0", 10);
  if (fails >= MAX_FAILS) return json({ ok: false, error: "Too many tries; please wait an hour or ask the office" }, 429, cors);
  const portal = await loadPortal(env);
  if (!portal || !portal.password) return json({ ok: false, error: "The Teacher Portal isn't open yet" }, 503, cors);
  if ((await passwordHash(String(body.password || "").slice(0, 200))) !== portal.password) {
    await env.LUNCH_KV.put(`rltp:${ip}`, String(fails + 1), { expirationTtl: 3600 });
    return json({ ok: false, error: "That password isn't right" }, 401, cors);
  }
  const hours = body.remember ? LOGIN_DAYS * 24 : SHORT_HOURS;
  return json({ ok: true, token: await makeToken(env, portal, "tp", hours), hours }, 200, cors);
}

export async function portalPage(request, env, cors, json) {
  if (!cors["Access-Control-Allow-Origin"]) return json({ ok: false, error: "Not allowed" }, 403, cors);
  const portal = await signedIn(request, env, cors, json);
  if (portal instanceof Response) return portal;
  return json({ ok: true, html: portal.html, files: await makeToken(env, portal, "tpf", FILE_HOURS) },
    200, { ...cors, "Cache-Control": "no-store" });
}

// /teachers/file/<pass>/<name>: streamed from the private repo's files/ folder.
export async function file(url, env) {
  const deny = (status, msg) => new Response(msg, { status, headers: { "Content-Type": "text/plain; charset=utf-8" } });
  const m = /^\/teachers\/file\/([^/]+)\/([^/]+)$/.exec(url.pathname);
  if (!m) return deny(404, "Not found");
  let name;
  try { name = decodeURIComponent(m[2]); } catch { return deny(404, "Not found"); }
  const type = FILE_TYPES[(/\.([a-z0-9]+)$/i.exec(name) || [])[1]?.toLowerCase()];
  if (!type || /[\\/\u0000-\u001f]|^\./.test(name)) return deny(404, "Not found");
  const portal = await loadPortal(env);
  if (!portal || !(await checkToken(env, portal, "tpf", m[1]))) {
    return deny(403, "This link has expired. Go back to the Teacher Portal (ibaclinton.com/teachers/) and open the file again.");
  }
  const r = await gitHub(env, `files/${name}`);
  if (r.status === 404) return deny(404, "That file isn't on the Teacher Portal any more.");
  if (!r.ok) return deny(502, "The file couldn't be loaded; please try again.");
  return new Response(r.body, { headers: {
    "Content-Type": type, "Cache-Control": "private, no-store", "X-Robots-Tag": "noindex",
    "Content-Disposition": `${type.startsWith("image/") || type === "application/pdf" ? "inline" : "attachment"}; ` +
      `filename*=UTF-8''${encodeURIComponent(name)}`,
  } });
}
