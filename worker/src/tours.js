// School tour sign-ups, on the same service and workbook as hot lunch and conferences.
//
// The tours (date, time, spots, deadline) come from the published list (site/tours/schedule.json, made
// by tours.py from what the office types in the editor). Each sign-up is one row of the "Signups" table
// on the Tours tab of the office workbook; the service adds that tab when it is missing. The table decides
// how full a tour is: a row the office deletes frees a spot, and a row the office types in (for a family
// who calls) takes one. They type the tour date like 1/28/2027 and the time like 1:00 PM.

import { normTime } from "./conferences.js";

export const SHEET = "Tours";
export const TABLE = "Signups";
export const COLUMNS = ["Signed up", "Tour date", "Time", "Parent", "Email", "Phone", "Children and grades", "Notes", "Signup #"];
const WIDTHS = [130, 80, 70, 150, 200, 110, 260, 260, 80];

export async function ensureTourTab(call) {
  const r = await call("GET", "/worksheets?$select=name");
  if (!r.ok) throw new Error(`worksheets ${r.status}`);
  if (((await r.json()).value || []).some((w) => w.name === SHEET)) return false;
  await must(call("POST", "/worksheets/add", { name: SHEET }));
  const s = `/worksheets/${SHEET}`;
  const last = String.fromCharCode(64 + COLUMNS.length);
  await must(call("PATCH", `${s}/range(address='B1:C2')`, { numberFormat: [["@", "@"], ["@", "@"]] }));
  await must(call("PATCH", `${s}/range(address='A1:${last}1')`, { values: [COLUMNS] }));
  const t = await must(call("POST", `${s}/tables/add`, { address: `A1:${last}1`, hasHeaders: true }));
  await must(call("PATCH", `/tables/${encodeURIComponent((await t.json()).id)}`, { name: TABLE }));
  for (let i = 0; i < COLUMNS.length; i++) {
    const c = String.fromCharCode(65 + i);
    await call("PATCH", `${s}/range(address='${c}:${c}')/format`, { columnWidth: WIDTHS[i] });
  }
  await call("PATCH", `${s}/range(address='K1:K4')`, {
    values: [["Tour sign-ups from the website land here by themselves, one row per family."],
      ["Each tour takes the number of families set in the website editor; then it shows as full."],
      ["Delete a row to cancel; that spot opens again. For a family who calls, add a row: Tour date like 1/28/2027, Time like 1:00 PM."],
      ["Keep the column names as they are."]],
  });
  return true;
}

// Every sign-up row as text, or null when the table isn't there yet.
export async function readSignups(call) {
  const r = await call("GET", `/tables/${TABLE}/range?$select=text`);
  if (r.status === 404) return null;
  if (!r.ok) throw new Error(`tours ${r.status}`);
  const [head, ...rows] = (await r.json()).text || [];
  const col = Object.fromEntries((head || []).map((h, i) => [String(h).trim(), i]));
  for (const need of ["Tour date", "Time", "Email"]) {
    if (!(need in col)) throw new Error(`Signups table has no "${need}" column`);
  }
  return rows.map((r) => ({
    date: normDate(r[col["Tour date"]]), time: normTime(r[col.Time]),
    email: String(r[col.Email] || "").trim().toLowerCase(),
  })).filter((s) => s.date && s.time);
}

// How many families have signed up for each tour: {"2027-01-28-1300": 3}.
export function takenCounts(schedule, signups) {
  const out = {};
  for (const t of schedule.tours || []) {
    const n = (signups || []).filter((s) => s.date === t.date && s.time === normTime(t.time)).length;
    if (n) out[t.id] = n;
  }
  return out;
}

export function validateSignup(body, schedule, now) {
  const str = (v, max) => String(v == null ? "" : v).replace(/[\u0000-\u001f]/g, " ").trim().slice(0, max);
  const tour = ((schedule && schedule.tours) || []).find((t) => t.id === str(body.tour, 40));
  if (!tour) return { error: "Please choose a tour; if the page is old, reload it" };
  if (now > Date.parse(tour.deadline) || now > Date.parse(tour.starts)) {
    return { error: "Online sign-up for this tour is closed; please call the office" };
  }
  const s = {
    tour, first: str(body.parent_first, 60), last: str(body.parent_last, 60),
    email: str(body.email, 120).toLowerCase(), phone: str(body.phone, 30),
    children: str(body.children, 400), notes: str(body.notes, 500),
  };
  if (!s.first || !s.last) return { error: "Please enter your first and last name" };
  if (!/^[^@\s<>"]+@[^@\s<>"]+\.[^@\s<>"]+$/.test(s.email)) return { error: "Please enter a valid email address" };
  if (!s.phone) return { error: "Please enter a phone number" };
  if (!s.children) return { error: "Please enter your children's names and grades" };
  return { signup: s };
}

// Why this family can't join the tour, or null if it can.
export function tourConflict(s, signups) {
  const mine = (signups || []).filter((x) => x.date === s.tour.date && x.time === normTime(s.tour.time));
  if (mine.some((x) => x.email === s.email)) {
    return { code: "already", error: `This email is already signed up for the tour on ${s.tour.label}. ` +
      "To change it, please call the office at (301) 856-1616." };
  }
  if (mine.length >= s.tour.spots) {
    return { code: "full", error: `The tour on ${s.tour.label} just filled up. Please pick another tour or call the office.` };
  }
  return null;
}

export function signupRow(s, id, submitted) {
  return [submitted, usDate(s.tour.date), s.tour.time, text(`${s.first} ${s.last}`), s.email, text(s.phone),
    text(s.children), text(s.notes), id];
}

export function familyMail(s, id, env) {
  const e = escapeHtml;
  const content =
    `<p>Dear ${e(s.first)} ${e(s.last)},</p>` +
    `<p>Thank you for signing up for a tour of Independent Baptist Academy. We look forward to meeting you!</p>` +
    `<p style="margin:12px 0"><b>${e(s.tour.label)}</b><br>9255 Piscataway Rd, Clinton, MD 20735</p>` +
    (s.tour.note ? `<p>${e(s.tour.note)}</p>` : "") +
    `<p>Please check in at the school office when you arrive. To change or cancel, please call the office at ` +
    `(301) 856-1616 or reply to this email. (Sign-up number ${e(id)}.)</p>` +
    `<p>Gratefully,<br>${e(env.SIGNATURE || "Independent Baptist Academy").replace(/\n/g, "<br>")}</p>`;
  const message = {
    subject: `School tour: ${s.tour.label}`,
    body: { contentType: "HTML", content },
    toRecipients: [{ emailAddress: { address: s.email, name: `${s.first} ${s.last}` } }],
  };
  if (env.REPLY_TO) message.replyTo = [{ emailAddress: { address: env.REPLY_TO } }];
  return { message, saveToSentItems: true };
}

export function officeMail(s, id, count, env) {
  const e = escapeHtml;
  const content =
    `<p>A family signed up for a school tour on the website:</p>` +
    `<p style="margin:12px 0"><b>${e(s.tour.label)}</b> (${count} of ${s.tour.spots} spots now taken)<br>` +
    `Parent: ${e(s.first)} ${e(s.last)}<br>Email: ${e(s.email)}<br>Phone: ${e(s.phone)}<br>` +
    `Children: ${e(s.children)}` + (s.notes ? `<br>Notes: ${e(s.notes)}` : "") + `</p>` +
    `<p>Every sign-up is on the Tours tab of the office workbook (sign-up ${e(id)}).</p>`;
  return {
    message: {
      subject: `Tour sign-up: ${s.first} ${s.last}, ${s.tour.label}`,
      body: { contentType: "HTML", content },
      toRecipients: String(env.TOUR_NOTIFY).split(",").map((a) => ({ emailAddress: { address: a.trim() } })),
      replyTo: [{ emailAddress: { address: s.email, name: `${s.first} ${s.last}` } }],
    },
    saveToSentItems: true,
  };
}

// Removes one TEST sign-up by its sign-up number (admin upkeep); only a row whose Parent starts with "TEST".
export async function deleteTestSignup(call, id) {
  const r = await call("GET", `/tables/${TABLE}/range?$select=text`);
  if (!r.ok) return { deleted: false, error: `table ${r.status}` };
  const [head, ...rows] = (await r.json()).text || [];
  const col = Object.fromEntries((head || []).map((h, i) => [String(h).trim(), i]));
  const i = rows.findIndex((row) => String(row[col["Signup #"]] || "").trim() === id);
  if (i < 0) return { deleted: false, error: "no sign-up with that number" };
  if (!/^TEST/i.test(String(rows[i][col.Parent] || "").trim())) return { deleted: false, error: "not a TEST sign-up" };
  let d = await call("DELETE", `/tables/${TABLE}/rows/itemAt(index=${i})`);
  if (!d.ok) d = await call("DELETE", `/tables/${TABLE}/rows/${i}`);
  return d.ok ? { deleted: true } : { deleted: false, error: `delete ${d.status}` };
}

// ------------------------------------------------------------------ helpers
function normDate(v) {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(v || "").trim());
  if (m) return `${m[3]}-${m[1].padStart(2, "0")}-${m[2].padStart(2, "0")}`;
  const iso = /^(\d{4}-\d{2}-\d{2})/.exec(String(v || "").trim());
  return iso ? iso[1] : "";
}

function text(s) { return /^[=+\-@]/.test(s) ? `'${s}` : s; }

function usDate(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return `${m}/${d}/${y}`;
}

async function must(p) {
  const r = await p;
  if (!r.ok) throw new Error(`workbook ${r.status} ${(await r.text()).slice(0, 200)}`);
  return r;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
