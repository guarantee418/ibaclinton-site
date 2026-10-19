// Report Student Absence (The Kings Corner's "Report an Absence" pop-up), on the same service and workbook as
// hot lunch, conferences, tours, transcripts, Request Info and feedback. It replaced a Microsoft Form.
//
// Each report is one row of the "Reports" table on the Absences tab of the office workbook (the service adds that
// tab when it is missing). The parent gets a confirmation; ABSENCE_NOTIFY (the office) gets the details, with the
// student's homeroom teacher (Teachers tab, matched by grade) Cc'd when the tab has an email for that grade.

import { GRADES as HOMEROOM_GRADES } from "./homerooms.js";

export const SHEET = "Absences";
export const TABLE = "Reports";
export const COLUMNS = ["Received", "Absence #", "First day", "Last day", "Student", "Grade", "Reason", "Details",
  "Parent", "Email", "Phone", "Teacher told", "Notes"];
const WIDTHS = [130, 85, 85, 85, 160, 120, 150, 260, 160, 200, 110, 140, 240];

export const GRADES = HOMEROOM_GRADES.map(([g]) => g); // the Teachers tab's grade labels, youngest first
export const REASONS = ["Illness", "Medical or dental appointment", "Family matter", "Travel", "Other"];
const DAY = 86400000;

export async function ensureAbsenceTab(call) {
  const r = await call("GET", "/worksheets?$select=name");
  if (!r.ok) throw new Error(`worksheets ${r.status}`);
  if (((await r.json()).value || []).some((w) => w.name === SHEET)) return false;
  await must(call("POST", "/worksheets/add", { name: SHEET }));
  const s = `/worksheets/${SHEET}`;
  const last = String.fromCharCode(64 + COLUMNS.length);
  await must(call("PATCH", `${s}/range(address='A1:${last}1')`, { values: [COLUMNS] }));
  const t = await must(call("POST", `${s}/tables/add`, { address: `A1:${last}1`, hasHeaders: true }));
  await must(call("PATCH", `/tables/${encodeURIComponent((await t.json()).id)}`, { name: TABLE }));
  for (let i = 0; i < COLUMNS.length; i++) {
    const c = String.fromCharCode(65 + i);
    await call("PATCH", `${s}/range(address='${c}:${c}')/format`, { columnWidth: WIDTHS[i] });
  }
  const note = String.fromCharCode(66 + COLUMNS.length);
  await call("PATCH", `${s}/range(address='${note}1:${note}3')`, {
    values: [["Absence reports from the website land here by themselves, one row per report."],
      ["The office gets an email for each one; the homeroom teacher is Cc'd when the Teachers tab has their email."],
      ["Keep the column names as they are."]],
  });
  return true;
}

export function validateAbsence(body, now) {
  const str = (v, max) => String(v == null ? "" : v).replace(/[\u0000-\u001f]/g, " ").trim().slice(0, max);
  const q = {
    first: str(body.student_first, 60), last: str(body.student_last, 60), grade: str(body.grade, 30),
    from: str(body.first_day, 10), to: str(body.last_day, 10), reason: str(body.reason, 40), details: str(body.details, 500),
    parent: str(body.parent_name, 120), email: str(body.email, 120).toLowerCase(), phone: str(body.phone, 30),
  };
  if (!q.first || !q.last) return { error: "Please enter the student's first and last name" };
  if (!GRADES.includes(q.grade)) return { error: "Please choose the student's grade" };
  const today = Date.parse(easternIso(now));
  const from = Date.parse(q.from);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(q.from) || !(from >= today - 14 * DAY) || !(from <= today + 365 * DAY)) {
    return { error: "Please enter the day of the absence (from two weeks ago to a year ahead)" };
  }
  if (q.to && q.to !== q.from) {
    const to = Date.parse(q.to);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(q.to) || !(to > from) || !(to <= from + 60 * DAY)) {
      return { error: "Please check the last day: on or after the first day, and within 60 days" };
    }
  } else q.to = "";
  if (!REASONS.includes(q.reason)) return { error: "Please choose a reason" };
  if (q.reason === "Other" && !q.details) return { error: "Please tell us the reason for the absence" };
  if (!q.parent) return { error: "Please enter your name" };
  if (!/^[^@\s<>"]+@[^@\s<>"]+\.[^@\s<>"]+$/.test(q.email)) return { error: "Please enter a valid email address" };
  if (!q.phone) return { error: "Please enter a phone number" };
  return { absence: q };
}

// The homeroom teacher for a grade, from the Teachers tab rows (readTeachers): { teacher, email } or null.
export function homeroomFor(teachers, grade) {
  const t = (teachers || []).find((x) => x.grade === grade && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x.email));
  return t ? { teacher: t.teacher, email: t.email } : null;
}

export function absenceRow(q, id, received, teacher) {
  return [received, id, usDate(q.from), q.to ? usDate(q.to) : "", text(`${q.first} ${q.last}`), q.grade, q.reason,
    text(q.details), text(q.parent), q.email, text(q.phone), teacher ? text(teacher.teacher) : "No email on Teachers tab", ""];
}

export function parentMail(q, id, env) {
  const e = escapeHtml;
  const content =
    `<p>Dear ${e(q.parent)},</p>` +
    `<p>Thank you for letting us know. We received your absence report for <b>${e(q.first)} ${e(q.last)}</b> ` +
    `(${e(q.grade)}): <b>${e(when(q))}</b>, ${e(q.reason.toLowerCase())}. Your report number is <b>${e(id)}</b>.</p>` +
    `<p>The school office and your child's homeroom teacher have been told. If anything changes, reply to this ` +
    `email or call the school office at (301) 856-1616.</p>` +
    `<p>Gratefully,<br>${e(env.SIGNATURE || "Independent Baptist Academy").replace(/\n/g, "<br>")}</p>`;
  const message = {
    subject: `Absence report ${id} received: ${q.first} ${q.last}`,
    body: { contentType: "HTML", content },
    toRecipients: [{ emailAddress: { address: q.email, name: q.parent } }],
  };
  if (env.REPLY_TO) message.replyTo = [{ emailAddress: { address: env.REPLY_TO } }];
  return { message, saveToSentItems: true };
}

export function officeMail(q, id, env, teacher) {
  const e = escapeHtml;
  const lines = [`Student: ${q.first} ${q.last}, ${q.grade}`, `Absent: ${when(q)}`,
    `Reason: ${q.reason}${q.details ? ` (${q.details})` : ""}`, `Parent: ${q.parent}, ${q.email}, ${q.phone}`];
  const content =
    `<p>A parent reported an absence on the website:</p>` +
    `<p style="margin:12px 0">${lines.map(e).join("<br>")}</p>` +
    `<p>Reply to this email to answer the parent. Every report is on the Absences tab of the office workbook ` +
    `(report ${e(id)}).${teacher ? "" : " No homeroom teacher email is on the Teachers tab for this grade, so the teacher was not told."}</p>`;
  const message = {
    subject: `Absence: ${q.first} ${q.last} (${q.grade}), ${when(q)}`,
    body: { contentType: "HTML", content },
    toRecipients: String(env.ABSENCE_NOTIFY).split(",").map((a) => ({ emailAddress: { address: a.trim() } })),
    replyTo: [{ emailAddress: { address: q.email, name: q.parent } }],
  };
  if (teacher) message.ccRecipients = [{ emailAddress: { address: teacher.email, name: teacher.teacher } }];
  return { message, saveToSentItems: true };
}

// Removes one TEST report by its number (admin upkeep); only a row whose Student starts with "TEST".
export async function deleteTestAbsence(call, id) {
  const r = await call("GET", `/tables/${TABLE}/range?$select=text`);
  if (!r.ok) return { deleted: false, error: `table ${r.status}` };
  const [head, ...rows] = (await r.json()).text || [];
  const col = Object.fromEntries((head || []).map((h, i) => [String(h).trim(), i]));
  const i = rows.findIndex((row) => String(row[col["Absence #"]] || "").trim() === id);
  if (i < 0) return { deleted: false, error: "no report with that number" };
  if (!/^TEST/i.test(String(rows[i][col.Student] || "").trim())) return { deleted: false, error: "not a TEST report" };
  let d = await call("DELETE", `/tables/${TABLE}/rows/itemAt(index=${i})`);
  if (!d.ok) d = await call("DELETE", `/tables/${TABLE}/rows/${i}`);
  return d.ok ? { deleted: true } : { deleted: false, error: `delete ${d.status}` };
}

// ------------------------------------------------------------------ helpers
function when(q) {
  const d = (iso) => new Date(`${iso}T12:00:00Z`).toLocaleDateString("en-US",
    { weekday: "long", month: "long", day: "numeric", year: "numeric", timeZone: "UTC" });
  return q.to ? `${d(q.from)} through ${d(q.to)}` : d(q.from);
}

function easternIso(now) {
  const p = Object.fromEntries(new Intl.DateTimeFormat("en-US", { timeZone: "America/New_York", year: "numeric",
    month: "2-digit", day: "2-digit" }).formatToParts(new Date(now)).map((x) => [x.type, x.value]));
  return `${p.year}-${p.month}-${p.day}`;
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
