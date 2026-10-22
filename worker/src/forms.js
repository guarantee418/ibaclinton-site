// Forms the office adds to The Kings Corner sections in the online editor (forms.py; site/forms.js sends them),
// on the same service and workbook as hot lunch, conferences, tours, transcripts, Request Info and feedback.
//
// The questions come from the published site/forms/forms.json (FORMS_URL), never from the browser. Each form has
// its own tab in the office workbook, named after the form and made by its first response. A question added later
// adds a column at the end of that tab's table; when the questions no longer line up with the table's headings
// (a question renamed, removed or moved, or a heading changed in Excel), the next response starts a new tab
// ("Name 2", table "Form_Name_2"), so an answer never lands under the wrong heading. The tab in use is kept in KV.
// The person gets a confirmation with their answers; FORM_NOTIFY (iba@ and ibs@) gets the details, Reply-To them.

export const FIXED = ["Received", "Response #", "Status", "Notes", "Name", "Email"]; // forms.py FIXED
const FIXED_WIDTHS = [130, 90, 100, 200, 160, 210];
const MAX_QUESTIONS = 15;
export const WAIT = { ms: 4000 }; // how long to wait for a tab another response is making (tests set 0)

export function columns(form) { return [...FIXED, ...form.questions.map((q) => q.label)]; }

export function tableFor(form, version) { return version > 1 ? `${form.table}_${version}` : form.table; }
export function sheetFor(form, version) { return version > 1 ? `${form.sheet} ${version}` : form.sheet; }

// What to do with a table whose headings are `head`, for a form whose columns are `want`: write as is, add the
// new questions' columns at the end, or start a new tab.
export function plan(head, want) {
  const h = (head || []).map((x) => String(x == null ? "" : x).trim());
  if (h.length > want.length || h.some((x, i) => x.toLowerCase() !== want[i].toLowerCase())) return { start: true };
  return { add: want.slice(h.length) };
}

export function validateResponse(body, forms, now = Date.now()) {
  const str = (v, max) => String(v == null ? "" : v).replace(/[\u0000-\u0008\u000b-\u001f]/g, " ").trim().slice(0, max);
  const form = (forms && Array.isArray(forms.forms) ? forms.forms : []).find((f) => f && f.id === str(body.form, 80));
  if (!form || !Array.isArray(form.questions) || !form.questions.length) {
    return { error: "This form isn't on the website any more; please reload the page" };
  }
  if (form.closed) return { error: "This form isn't taking responses right now" };
  if (str(body.version, 20) !== form.version) return { error: "This form changed; please reload the page and fill it out again" };
  const name = str(body.name, 120).replace(/\s+/g, " "), email = str(body.email, 120).toLowerCase();
  if (!name) return { error: "Please enter your name" };
  if (!/^[^@\s<>"]+@[^@\s<>"]+\.[^@\s<>"]+$/.test(email)) return { error: "Please enter a valid email address" };
  const given = body.answers && typeof body.answers === "object" ? body.answers : {};
  const answers = [];
  for (const q of form.questions.slice(0, MAX_QUESTIONS)) {
    const raw = given[q.key];
    let v = "";
    if (q.type === "multi") {
      const picked = (Array.isArray(raw) ? raw : []).map((x) => str(x, 80).toLowerCase());
      v = q.choices.filter((c) => picked.includes(c.toLowerCase())).join("; ");
      if (picked.length && !v) return { error: `Please choose again for "${q.label}"` };
    } else if (q.type === "check") {
      v = raw === true || /^(yes|true|on)$/i.test(str(raw, 5)) ? "Yes" : "";
    } else {
      v = str(raw, q.type === "long" ? 4000 : 200);
      if (v && q.type === "choice") {
        v = q.choices.find((c) => c.toLowerCase() === v.toLowerCase()) || "";
        if (!v) return { error: `Please choose an answer for "${q.label}"` };
      }
      if (v && q.type === "email" && !/^[^@\s<>"]+@[^@\s<>"]+\.[^@\s<>"]+$/.test(v)) return { error: `Please check "${q.label}"` };
      if (v && q.type === "phone" && (v.replace(/\D/g, "").length < 7 || /[^\d\s()+.\-x]/i.test(v))) return { error: `Please check "${q.label}"` };
      if (v && q.type === "number" && !/^-?\d+(\.\d+)?$/.test(v.replace(/,/g, ""))) return { error: `Please enter a number for "${q.label}"` };
      if (v && q.type === "date") {
        const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
        const d = m && new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
        if (!d || d.getUTCDate() !== +m[3] || +m[1] < 1900 || +m[1] > 2100) return { error: `Please check the date for "${q.label}"` };
        v = `${+m[2]}/${+m[3]}/${m[1]}`;
      }
    }
    if (q.required && !v) {
      return { error: q.type === "check" ? `Please tick "${q.label}"` : `Please answer "${q.label}"` };
    }
    answers.push(v);
  }
  return { form, entry: { name, email, answers } };
}

export function responseRow(entry, id, received) {
  return [received, id, "New", "", text(entry.name), entry.email, ...entry.answers.map((a) => (/^\d+\/\d+\/\d{4}$/.test(a) ? a : text(a)))];
}

// Adds the row to the form's table, making the tab, adding columns or starting a new tab as needed.
// Returns { ok, table, sheet } or { ok: false, status }.
export async function saveResponse(call, kv, form, row) {
  const key = `formv:${form.id}`;
  const start = Math.max(1, parseInt((await kv.get(key)) || "1", 10));
  const want = columns(form);
  for (let v = start; v < start + 4; v++) {
    const table = tableFor(form, v);
    const h = await call("GET", `/tables/${encodeURIComponent(table)}/headerRowRange?$select=values`);
    if (h.status === 404) {
      if (!(await makeTab(call, form, v))) {
        // The tab name is taken: another response may be making this very tab right now, else (an office tab
        // with that name) try the next one.
        await new Promise((ok) => setTimeout(ok, WAIT.ms));
        if (!(await call("GET", `/tables/${encodeURIComponent(table)}/headerRowRange?$select=values`)).ok) continue;
      }
    } else if (!h.ok) {
      return { ok: false, status: h.status };
    } else {
      const p = plan(((await h.json()).values || [[]])[0], want);
      if (p.start) continue;
      for (const name of p.add) {
        const c = await call("POST", `/tables/${encodeURIComponent(table)}/columns/add`, { index: null, name });
        if (!c.ok) return { ok: false, status: c.status };
      }
    }
    const add = await call("POST", `/tables/${encodeURIComponent(table)}/rows/add`, { index: null, values: [row] });
    if (!add.ok) return { ok: false, status: add.status };
    if (v !== start) await kv.put(key, String(v));
    return { ok: true, table, sheet: sheetFor(form, v) };
  }
  return { ok: false, status: 409 };
}

async function makeTab(call, form, version) {
  const sheet = sheetFor(form, version), cols = columns(form);
  const add = await call("POST", "/worksheets/add", { name: sheet });
  if (!add.ok) return false;
  const s = `/worksheets/${encodeURIComponent(sheet)}`, last = colLetter(cols.length - 1);
  await must(call("PATCH", `${s}/range(address='A1:${last}1')`, { values: [cols] }));
  const t = await must(call("POST", `${s}/tables/add`, { address: `A1:${last}1`, hasHeaders: true }));
  await must(call("PATCH", `/tables/${encodeURIComponent((await t.json()).id)}`, { name: tableFor(form, version) }));
  for (let i = 0; i < FIXED.length; i++) {
    await call("PATCH", `${s}/range(address='${colLetter(i)}:${colLetter(i)}')/format`, { columnWidth: FIXED_WIDTHS[i] });
  }
  await call("PATCH", `${s}/range(address='${colLetter(FIXED.length)}:${last}')/format`, { columnWidth: 190, wrapText: true });
  const note = colLetter(cols.length + 1);
  const lines = [
    `Responses to the "${form.name}" form on ${form.where || "The Kings Corner"}, one row per response.`,
    "Each person gets a confirmation email with their answers; iba@ and ibs@ get the details.",
    "Change Status as you go (for example Read, Done) and use Notes for anything worth remembering.",
    "⚠ NEVER rename, move or delete these columns. A question added in the website editor adds a column at the end; " +
      "renaming, removing or reordering questions starts a new tab.",
  ];
  await call("PATCH", `${s}/range(address='${note}1:${note}${lines.length}')`, { values: lines.map((l) => [l]) });
  await call("PATCH", `${s}/range(address='${note}1')/format/font`, { bold: true, color: "#1B2A4A" });
  await call("PATCH", `${s}/range(address='${note}${lines.length}')/format/font`, { bold: true, color: "#B00020" });
  return true;
}

export function confirmationMail(form, entry, id, env) {
  const e = escapeHtml;
  const content =
    `<p>Hello ${e(entry.name)},</p>` +
    `<p>Thank you! We received your response to <b>${e(form.name)}</b> (number ${e(id)}). Here is a copy of your answers:</p>` +
    answerList(form, entry) +
    `<p>If something needs to change, reply to this email or call the school office at (301) 856-1616.</p>` +
    `<p>Gratefully,<br>${e(env.SIGNATURE || "Independent Baptist Academy").replace(/\n/g, "<br>")}</p>`;
  const message = {
    subject: `${form.name}: response ${id} received`,
    body: { contentType: "HTML", content },
    toRecipients: [{ emailAddress: { address: entry.email, name: entry.name } }],
  };
  if (env.REPLY_TO) message.replyTo = [{ emailAddress: { address: env.REPLY_TO } }];
  return { message, saveToSentItems: true };
}

// A Teacher Portal form can name its own approvers (teachers.py routing): each is asked only when its question has
// its answer (e.g. a vehicle only when one is needed). Returns them with needed: true/false.
export function approvalsFor(form, entry) {
  return (form.approvals || []).map((a) => {
    if (!a.question) return { ...a, needed: true };
    const i = form.questions.findIndex((q) => q.key === a.question);
    const given = String(i < 0 ? "" : entry.answers[i] || "").toLowerCase().split("; ");
    return { ...a, needed: i < 0 || given.includes(String(a.answer).toLowerCase()) };
  });
}

// Who gets the plain notice: the form's own list (else FORM_NOTIFY), plus approvers who aren't needed this time,
// leaving out anyone who gets an approval request instead.
export function noticeRecipients(form, env, approvals = []) {
  const asked = new Set(approvals.filter((a) => a.needed).map((a) => a.email.toLowerCase()));
  const list = (form.notify && form.notify.length ? form.notify : String(env.FORM_NOTIFY || "").split(","))
    .concat(approvals.filter((a) => !a.needed).map((a) => a.email));
  return [...new Set(list.map((a) => a.trim().toLowerCase()).filter((a) => a && !asked.has(a)))];
}

function approvalLines(approvals) {
  const e = escapeHtml;
  if (!approvals.length) return "";
  return `<p><b>Approvals:</b><br>${approvals.map((a) => `${e(a.title)}${a.name ? ` (${e(a.name)})` : ""}: ` +
    (a.needed ? "asked to approve" : "not needed for this one")).join("<br>")}</p>`;
}

export function approvalMail(form, entry, id, a, replyTo) {
  const e = escapeHtml;
  const content =
    `<p>${a.name ? `${e(a.name)}, a` : "A"}pproval is needed for this <b>${e(form.name)}</b> from ${e(entry.name)} ` +
    `(${e(entry.email)}), sent on ${e(form.where || "The Kings Corner")}:</p>` +
    answerList(form, entry) +
    `<p><b>Please reply to this email to give your approval (${e(a.title)})</b>, or to say what needs to change. ` +
    `Your reply goes to ${e(replyTo)}.</p>`;
  return {
    message: {
      subject: `Approval needed (${a.title}): ${form.name} from ${entry.name} [${id}]`,
      body: { contentType: "HTML", content },
      toRecipients: [{ emailAddress: { address: a.email, name: a.name || undefined } }],
      replyTo: [{ emailAddress: { address: replyTo } }],
    },
    saveToSentItems: true,
  };
}

export function noticeMail(form, entry, id, env, sheet, approvals = []) {
  const e = escapeHtml;
  const content =
    `<p>A new response to <b>${e(form.name)}</b> came in on ${e(form.where || "The Kings Corner")}:</p>` +
    answerList(form, entry, true) + approvalLines(approvals) +
    `<p>Reply to this email to answer ${e(entry.name)}. Every response is on the "${e(sheet)}" tab of the office ` +
    `workbook (response ${e(id)}).</p>`;
  return {
    message: {
      subject: `${form.name}: ${entry.name}`,
      body: { contentType: "HTML", content },
      toRecipients: noticeRecipients(form, env, approvals).map((address) => ({ emailAddress: { address } })),
      replyTo: [{ emailAddress: { address: entry.email, name: entry.name } }],
    },
    saveToSentItems: true,
  };
}

function answerList(form, entry, withContact) {
  const e = escapeHtml;
  const rows = withContact ? [["Name", entry.name], ["Email", entry.email]] : [];
  form.questions.forEach((q, i) => rows.push([q.label, entry.answers[i] || "(no answer)"]));
  return `<table style="border-collapse:collapse;margin:8px 0 14px">${rows.map(([k, v]) =>
    `<tr><td style="padding:3px 14px 3px 0;vertical-align:top"><b>${e(k)}</b></td>` +
    `<td style="padding:3px 0;white-space:pre-wrap">${e(v)}</td></tr>`).join("")}</table>`;
}

// Removes one TEST response by its number (admin upkeep); only a row whose Name starts with "TEST". A tab whose
// name starts with "TEST" goes too once its last response is gone (a test form).
export async function deleteTestResponse(call, kv, id) {
  const where = JSON.parse((await kv.get(`formresp:${id}`)) || "null");
  if (!where) return { deleted: false, error: "no response with that number (or it is older than 60 days)" };
  const t = `/tables/${encodeURIComponent(where.table)}`;
  const r = await call("GET", `${t}/range?$select=text`);
  if (!r.ok) return { deleted: false, error: `table ${r.status}` };
  const [head, ...rows] = (await r.json()).text || [];
  const col = Object.fromEntries((head || []).map((h, i) => [String(h).trim(), i]));
  const i = rows.findIndex((row) => String(row[col["Response #"]] || "").trim() === id);
  if (i < 0) return { deleted: false, error: "no response with that number" };
  if (!/^TEST/i.test(String(rows[i][col.Name] || "").trim())) return { deleted: false, error: "not a TEST response" };
  let d = await call("DELETE", `${t}/rows/itemAt(index=${i})`);
  if (!d.ok) d = await call("DELETE", `${t}/rows/${i}`);
  if (!d.ok) return { deleted: false, error: `delete ${d.status}` };
  await kv.delete(`formresp:${id}`);
  const left = rows.filter((row, j) => j !== i && row.some((x) => String(x || "").trim() !== ""));
  if (/^TEST/i.test(where.sheet) && !left.length) {
    const w = await call("DELETE", `/worksheets/${encodeURIComponent(where.sheet)}`);
    return { deleted: true, tab_deleted: w.ok };
  }
  return { deleted: true };
}

// ------------------------------------------------------------------ helpers
export function colLetter(i) { return (i >= 26 ? colLetter(Math.floor(i / 26) - 1) : "") + String.fromCharCode(65 + (i % 26)); }

function text(s) { return /^[=+\-@]/.test(s) ? `'${s}` : s; }

async function must(p) {
  const r = await p;
  if (!r.ok) throw new Error(`workbook ${r.status} ${(await r.text()).slice(0, 200)}`);
  return r;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
