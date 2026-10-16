// Parent-teacher conference sign-ups, on the same service and workbook as hot lunch.
//
// The day, the times and the teachers come from the published schedule (site/conferences/schedule.json,
// made by conferences.py from what the office types in the editor). Each booking is one row of the
// "Bookings" table on the Conferences tab of the office workbook; the service adds that tab when it
// is missing. The table is also what decides which times are taken, so a row the office deletes frees
// its time, and a row the office types in (for a parent who calls) blocks it: they type the teacher as
// the website shows it and the time like "9:30 AM".

export const SHEET = "Conferences";
export const TABLE = "Bookings";
export const COLUMNS = ["Booked", "Date", "Time", "Teacher", "Student", "Parent", "Email", "Phone", "Notes", "Booking #"];
const WIDTHS = [130, 80, 70, 130, 150, 150, 200, 110, 260, 80];

// ------------------------------------------------------------------ the workbook tab
export async function ensureConferenceTab(call) {
  const r = await call("GET", "/worksheets?$select=name");
  if (!r.ok) throw new Error(`worksheets ${r.status}`);
  if (((await r.json()).value || []).some((w) => w.name === SHEET)) return false;
  await must(call("POST", "/worksheets/add", { name: SHEET }));
  const s = `/worksheets/${SHEET}`;
  const last = String.fromCharCode(64 + COLUMNS.length);
  // Date and Time stay text, so "9:30 AM" reads back exactly as the website wrote it.
  await must(call("PATCH", `${s}/range(address='B1:C2')`, { numberFormat: [["@", "@"], ["@", "@"]] }));
  await must(call("PATCH", `${s}/range(address='A1:${last}1')`, { values: [COLUMNS] }));
  const t = await must(call("POST", `${s}/tables/add`, { address: `A1:${last}1`, hasHeaders: true }));
  const id = (await t.json()).id;
  await must(call("PATCH", `/tables/${encodeURIComponent(id)}`, { name: TABLE }));
  for (let i = 0; i < COLUMNS.length; i++) {
    const c = String.fromCharCode(65 + i);
    await call("PATCH", `${s}/range(address='${c}:${c}')/format`, { columnWidth: WIDTHS[i] });
  }
  await call("PATCH", `${s}/range(address='L1:L4')`, {
    values: [["Conference sign-ups from the website land here by themselves, one row per booking."],
      ["To see one teacher's day, filter the Teacher column. Delete a row to cancel; that time opens again."],
      ["For a parent who calls, add a row: Date like 11/16/2026, Time like 9:30 AM, Teacher as on the website."],
      ["Keep the column names as they are."]],
  });
  return true;
}

// Every booking row as text, or null when the table isn't there yet.
export async function readBookings(call) {
  const r = await call("GET", `/tables/${TABLE}/range?$select=text`);
  if (r.status === 404) return null;
  if (!r.ok) throw new Error(`conferences ${r.status}`);
  const [head, ...rows] = (await r.json()).text || [];
  const col = Object.fromEntries((head || []).map((h, i) => [String(h).trim(), i]));
  for (const need of ["Date", "Time", "Teacher", "Email"]) {
    if (!(need in col)) throw new Error(`Conferences table has no "${need}" column`);
  }
  return rows.map((r) => ({
    date: normDate(r[col.Date]), time: normTime(r[col.Time]), teacher: normName(r[col.Teacher]),
    email: String(r[col.Email] || "").trim().toLowerCase(),
  })).filter((b) => b.teacher && b.time);
}

// Which times are taken, by teacher id: {"mr-damm": ["0900", "0930"]}.
export function takenSlots(schedule, bookings) {
  const byName = new Map(schedule.teachers.map((t) => [normName(t.name), t.id]));
  const byTime = new Map(schedule.slots.map((s) => [normTime(s.label), s.id]));
  const out = {};
  for (const b of bookings || []) {
    if (b.date && b.date !== schedule.date) continue;
    const tid = byName.get(b.teacher), sid = byTime.get(b.time);
    if (tid && sid) (out[tid] = out[tid] || []).push(sid);
  }
  return out;
}

// ------------------------------------------------------------------ a booking
export function validateBooking(body, schedule, now) {
  if (!schedule || !schedule.date || !schedule.slots || !schedule.slots.length || !schedule.teachers || !schedule.teachers.length) {
    return { error: "Conference sign-ups aren't open right now" };
  }
  if (!schedule.deadline || now > Date.parse(schedule.deadline)) {
    return { error: "Online sign-up is closed; please call the office" };
  }
  if (body.date && body.date !== schedule.date) return { error: "The schedule changed; please reload the page and try again" };
  const str = (v, max) => String(v == null ? "" : v).replace(/[\u0000-\u001f]/g, " ").trim().slice(0, max);
  const teacher = schedule.teachers.find((t) => t.id === str(body.teacher, 80));
  if (!teacher) return { error: "Please choose a teacher; if the page is old, reload it" };
  const slot = schedule.slots.find((s) => s.id === str(body.slot, 8));
  if (!slot) return { error: "Please choose a time" };
  const b = {
    teacher, slot,
    first: str(body.parent_first, 60), last: str(body.parent_last, 60),
    email: str(body.email, 120).toLowerCase(), phone: str(body.phone, 30),
    student: str(body.student, 80), notes: str(body.notes, 500),
  };
  if (!b.first || !b.last) return { error: "Please enter your first and last name" };
  if (!/^[^@\s<>"]+@[^@\s<>"]+\.[^@\s<>"]+$/.test(b.email)) return { error: "Please enter a valid email address" };
  if (!b.phone) return { error: "Please enter a phone number" };
  if (!b.student) return { error: "Please enter your student's name" };
  return { booking: b };
}

// Why this booking can't have its time, or "" if it can.
export function conflict(b, schedule, bookings) {
  const name = normName(b.teacher.name), time = normTime(b.slot.label);
  const mine = (bookings || []).filter((x) => (!x.date || x.date === schedule.date) && x.teacher === name);
  if (mine.some((x) => x.time === time)) {
    return { code: "taken", error: `Someone just booked ${b.slot.label} with ${b.teacher.name}. Please pick another time.` };
  }
  const already = mine.find((x) => x.email === b.email);
  if (already) {
    return { code: "already", error: `This email already has a conference with ${b.teacher.name} at ${already.time}. ` +
      "To change it, please call the office at (301) 856-1616." };
  }
  return null;
}

export function bookingRow(b, schedule, id, submitted) {
  return [submitted, usDate(schedule.date), b.slot.label, b.teacher.name, text(b.student), text(`${b.first} ${b.last}`),
    b.email, text(b.phone), text(b.notes), id];
}

export function parentMail(b, schedule, id, env) {
  const e = escapeHtml;
  const content =
    `<p>Dear ${e(b.first)} ${e(b.last)},</p>` +
    `<p>Thank you for signing up. Your parent-teacher conference is booked:</p>` +
    `<p style="margin:12px 0"><b>${e(b.teacher.name)}</b>${b.teacher.subject ? ` (${e(b.teacher.subject)})` : ""}<br>` +
    `${e(schedule.date_label)}, ${e(b.slot.label)} to ${e(b.slot.end)}<br>` +
    `About: ${e(b.student)}</p>` +
    (b.notes ? `<p>You mentioned: ${e(b.notes)}</p>` : "") +
    `<p>Booking number ${e(id)}. Please check in at the school office when you arrive. To change or cancel, ` +
    `please call the office at (301) 856-1616 or reply to this email.</p>` +
    `<p>Gratefully,<br>${e(env.SIGNATURE || "Independent Baptist Academy").replace(/\n/g, "<br>")}</p>`;
  const message = {
    subject: `Conference booked: ${b.teacher.name}, ${shortDate(schedule.date)} at ${b.slot.label}`,
    body: { contentType: "HTML", content },
    toRecipients: [{ emailAddress: { address: b.email, name: `${b.first} ${b.last}` } }],
  };
  if (env.REPLY_TO) message.replyTo = [{ emailAddress: { address: env.REPLY_TO } }];
  return { message, saveToSentItems: true };
}

export function teacherMail(b, schedule, id, to, name, env) {
  const e = escapeHtml;
  const shared = (b.teacher.members || []).length > 1;
  const content =
    `<p>${e(name)},</p>` +
    `<p>A parent signed up for a conference with you on the school website` +
    (shared ? ` (${e(b.teacher.name)}: ${e(b.teacher.members.join(", "))})` : "") + `:</p>` +
    `<p style="margin:12px 0"><b>${e(schedule.date_label)}, ${e(b.slot.label)} to ${e(b.slot.end)}</b><br>` +
    `Student: ${e(b.student)}<br>Parent: ${e(b.first)} ${e(b.last)}<br>Email: ${e(b.email)}<br>Phone: ${e(b.phone)}` +
    (b.notes ? `<br>Wants to talk about: ${e(b.notes)}` : "") + `</p>` +
    `<p>Every booking is on the Conferences tab of the office workbook (booking ${e(id)}). ` +
    `If something changes, the office will let you know.</p>`;
  const message = {
    subject: `Conference booked: ${b.student}, ${shortDate(schedule.date)} at ${b.slot.label}`,
    body: { contentType: "HTML", content },
    toRecipients: [{ emailAddress: { address: to, name } }],
  };
  if (env.REPLY_TO) message.replyTo = [{ emailAddress: { address: env.REPLY_TO } }];
  return { message, saveToSentItems: true };
}

// The teacher's email from the Teachers tab (grade, teacher, email), matched by name.
export function teacherEmail(teachers, name) {
  const n = normName(name);
  const t = (teachers || []).find((x) => normName(x.teacher) === n && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x.email));
  return t ? t.email : "";
}

// ------------------------------------------------------------------ helpers
export function normName(v) { return String(v || "").toLowerCase().replace(/\./g, "").replace(/\s+/g, " ").trim(); }

export function normTime(v) {
  const m = /^\s*(\d{1,2}):(\d{2})(?::\d{2})?\s*([ap])\.?\s*m?\.?\s*$/i.exec(String(v || ""));
  if (m) return `${Number(m[1])}:${m[2]} ${m[3].toUpperCase()}M`;
  const h24 = /^\s*(\d{1,2}):(\d{2})(?::\d{2})?\s*$/.exec(String(v || ""));
  if (h24) {
    const h = Number(h24[1]);
    return `${(h + 11) % 12 + 1}:${h24[2]} ${h < 12 ? "AM" : "PM"}`;
  }
  return String(v || "").trim().toUpperCase();
}

function normDate(v) {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(v || "").trim());
  if (m) return `${m[3]}-${m[1].padStart(2, "0")}-${m[2].padStart(2, "0")}`;
  const iso = /^(\d{4}-\d{2}-\d{2})/.exec(String(v || "").trim());
  return iso ? iso[1] : "";
}

// Text the sheet must not read as a formula; a leading apostrophe isn't shown.
function text(s) { return /^[=+\-@]/.test(s) ? `'${s}` : s; }

function usDate(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return `${m}/${d}/${y}`;
}

function shortDate(iso) {
  return new Date(`${iso}T12:00:00Z`).toLocaleDateString("en-US", { timeZone: "UTC", month: "long", day: "numeric" });
}

async function must(p) {
  const r = await p;
  if (!r.ok) throw new Error(`workbook ${r.status} ${(await r.text()).slice(0, 200)}`);
  return r;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
