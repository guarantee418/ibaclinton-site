// Hot lunch order service for ibaclinton.com (Cloudflare Worker).
//
//   POST /order     an order from the Parent Portal form: checked against the published menu,
//                   added to the "Orders" table of the office workbook, confirmation emailed
//   GET  /conferences  which conference times are taken, by teacher (no names)
//   POST /conference   a parent-teacher conference booking (src/conferences.js): checked against the
//                   published schedule, added to the Conferences tab of the office workbook, emailed
//   GET  /tours        how many families have signed up for each tour (no names)
//   POST /tour         a school tour sign-up (src/tours.js): checked against the published tours,
//                   refused when the tour is full, added to the Tours tab of the office workbook, emailed
//   POST /inquiry      a Request Info form (src/inquiries.js): added to the Inquiries tab of the office
//                   workbook, emailed to the family (confirmation) and the office
//   POST /absence      a Report Student Absence form (src/absences.js): added to the Absences tab of the office
//                   workbook, emailed to the parent (confirmation) and the office, Cc the homeroom teacher
//   POST /feedback     a Feedback for Administration form from the Parent Portal (src/feedback.js): added to
//                   the Feedback tab of the office workbook (no email)
//   GET  /transcripts  the transcript fees and the payment link (no names)
//   POST /transcript   a transcript request (src/transcripts.js): added to the Transcripts tab of the office
//                   workbook, emailed to the requester and the office
//   GET  /health    {"ok":true,"connected":true|false}
//   GET  /connect   one-time Microsoft sign-in by the office mailbox (needs ?key=SETUP_KEY)
//   GET  /callback  Microsoft sends the sign-in back here
//   POST /admin     upkeep, needs header X-Setup-Key: {"action":"setup"} adds the Teachers and
//                   Homerooms tabs (and the Conferences tab) to the workbook if missing (and, with
//                   "teacher_emails", fills blank teacher emails); {"action":"homerooms","week":"YYYY-MM-DD",
//                   "test_to":"someone@..."} sends that week's homeroom emails to test_to only;
//                   {"action":"delete_test_booking","id":"ABCDE"} removes one conference booking whose
//                   student is "TEST..."; {"action":"delete_test_order","id":"ABCDE"} removes a lunch order
//                   whose rows are all "TEST..." (Actions -> "Delete a TEST booking or lunch order");
//                   {"action":"delete_test_transcript","id":"ABCDE"} removes one transcript request whose
//                   student is "TEST..."; {"action":"delete_test_inquiry","id":"ABCDE"} likewise one inquiry;
//                   {"action":"delete_test_feedback","id":"ABCDE"} one feedback response whose name is "TEST...";
//                   {"action":"delete_test_absence","id":"ABCDE"} one absence report whose student is "TEST..."
//   cron            Mondays 7:00 AM Eastern: homeroom lunch lists to teachers (src/homerooms.js)
//
// Microsoft 365 access is delegated: the Worker acts only as the mailbox in MAILBOX
// (send mail as that user, edit that user's OneDrive files). It never sees a password;
// it keeps Microsoft's refresh token in KV. See worker/README.md for setup.

import { TABLE as CONF_TABLE, bookingRow, conflict, deleteTestBooking, ensureConferenceTab, parentMail, readBookings, takenSlots, teacherEmail,
  teacherMail, validateBooking } from "./conferences.js";
import { TABLE as TOUR_TABLE, deleteTestSignup, ensureTourTab, familyMail, officeMail, readSignups, signupRow, takenCounts,
  tourConflict, validateSignup } from "./tours.js";
import { TABLE as INQUIRY_TABLE, deleteTestInquiry, ensureInquiryTab, familyMail as inquiryFamilyMail, inquiryRow,
  officeMail as inquiryOfficeMail, validateInquiry } from "./inquiries.js";
import { TABLE as ABSENCE_TABLE, absenceRow, deleteTestAbsence, ensureAbsenceTab, homeroomFor, officeMail as absenceOfficeMail,
  parentMail as absenceParentMail, validateAbsence } from "./absences.js";
import { TABLE as FEEDBACK_TABLE, deleteTestFeedback, ensureFeedbackTab, feedbackRow, validateFeedback } from "./feedback.js";
import { TABLE as TRANSCRIPT_TABLE, deleteTestRequest, ensureTranscriptTab, officeMail as transcriptOfficeMail, requestRow,
  requesterMail, settings as transcriptSettings, validateRequest } from "./transcripts.js";
import { STEPS as ORGANIZE_STEPS, inspect as inspectWorkbook, organize as organizeWorkbook } from "./guide.js";
import { checkHomerooms, ensureTabs, fillTeacherEmails, weekCounts, homeroomMails, mondayToSend, readTeachers, readWeek } from "./homerooms.js";

const SCOPES = "offline_access User.Read Mail.Send Files.ReadWrite";
const DAY_KEYS = ["mon", "tue", "wed", "thu", "fri"];
const MAX_CHILDREN = 8;
const GRAPH = "https://graph.microsoft.com/v1.0";

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const cors = corsHeaders(request, env);
    try {
      if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: cors });
      if (url.pathname === "/health") {
        return json({ ok: true, connected: !!(await env.LUNCH_KV.get("ms_refresh")) }, 200, cors);
      }
      if (url.pathname === "/connect" && request.method === "GET") return connect(url, env);
      if (url.pathname === "/callback" && request.method === "GET") return callback(url, env);
      if (url.pathname === "/order" && request.method === "POST") return await order(request, env, cors);
      if (url.pathname === "/conferences" && request.method === "GET") return await conferences(env, cors);
      if (url.pathname === "/conference" && request.method === "POST") return await conference(request, env, cors);
      if (url.pathname === "/tours" && request.method === "GET") return await tours(env, cors);
      if (url.pathname === "/tour" && request.method === "POST") return await tour(request, env, cors);
      if (url.pathname === "/inquiry" && request.method === "POST") return await inquiry(request, env, cors);
      if (url.pathname === "/absence" && request.method === "POST") return await absence(request, env, cors);
      if (url.pathname === "/feedback" && request.method === "POST") return await feedback(request, env, cors);
      if (url.pathname === "/transcripts" && request.method === "GET") return transcripts(env, cors);
      if (url.pathname === "/transcript" && request.method === "POST") return await transcript(request, env, cors);
      if (url.pathname === "/admin" && request.method === "POST") return await admin(request, env);
      return json({ ok: false, error: "Not found" }, 404, cors);
    } catch (err) {
      console.error(err && err.stack || err);
      return json({ ok: false, error: "Something went wrong on our end" }, 500, cors);
    }
  },

  async scheduled(event, env, ctx) {
    const week = mondayToSend(new Date(event.scheduledTime));
    if (week) ctx.waitUntil(homeroomRun(env, week).then((r) => console.log("homerooms", JSON.stringify(r))));
  },
};

// ------------------------------------------------------------------ homeroom lists and upkeep
function workbook(token, env) {
  const base = `/me/drive/root:/${encodePath(env.WORKBOOK_PATH)}:/workbook`;
  return (method, sub, body) => graph(token, method, base + sub, body);
}

export async function homeroomRun(env, week, testTo = "") {
  const token = await accessToken(env);
  if (!token) return { ok: false, error: "not connected" };
  const call = workbook(token, env);
  const added = await ensureTabs(call);
  if (!testTo && await env.LUNCH_KV.get(`homerooms:${week}`)) return { ok: true, week, already_sent: true };
  const students = await readWeek(call, week);
  const teachers = await readTeachers(call);
  const mails = homeroomMails(students, teachers, week, env, testTo);
  let sent = 0;
  for (const m of mails) {
    const r = await graph(token, "POST", "/me/sendMail", m.mail);
    if (r.ok) sent++; else console.error("homeroom mail", r.status, await r.text());
  }
  if (!testTo) await env.LUNCH_KV.put(`homerooms:${week}`, String(sent), { expirationTtl: 60 * 86400 });
  return { ok: sent === mails.length, week, tabs_added: added, students: students.length, emails: mails.length, sent,
    grades_without_email: teachers.filter((t) => !t.email).length };
}

async function admin(request, env) {
  if (!env.SETUP_KEY || request.headers.get("X-Setup-Key") !== env.SETUP_KEY) return json({ ok: false, error: "Not allowed" }, 403, {});
  let body = {};
  try { body = JSON.parse(await request.text() || "{}"); } catch { return json({ ok: false, error: "Bad request" }, 400, {}); }
  if (body.action === "setup") {
    const token = await accessToken(env);
    if (!token) return json({ ok: false, error: "not connected" }, 503, {});
    const call = workbook(token, env);
    const added = await ensureTabs(call);
    if (await ensureConferenceTab(call)) added.push("Conferences");
    if (await ensureTourTab(call)) added.push("Tours");
    if (await ensureTranscriptTab(call)) added.push("Transcripts");
    if (await ensureInquiryTab(call)) added.push("Inquiries");
    if (await ensureFeedbackTab(call)) added.push("Feedback");
    if (await ensureAbsenceTab(call)) added.push("Absences");
    const filled = body.teacher_emails ? await fillTeacherEmails(call, body.teacher_emails) : 0;
    return json({ ok: true, tabs_added: added, teacher_emails_filled: filled, homerooms_tab: await checkHomerooms(call) }, 200, {});
  }
  if (body.action === "weeks") { // how many order rows per week (no names), for checking the workbook
    const token = await accessToken(env);
    if (!token) return json({ ok: false, error: "not connected" }, 503, {});
    return json({ ok: true, weeks: await weekCounts(workbook(token, env)) }, 200, {});
  }
  if (body.action === "inspect") { // the workbook's tabs, tables and notes (no rows), see src/guide.js
    const token = await accessToken(env);
    if (!token) return json({ ok: false, error: "not connected" }, 503, {});
    return json({ ok: true, sheets: await inspectWorkbook(workbook(token, env)) }, 200, {});
  }
  if (body.action === "organize") { // one step of tidying the workbook ("step": order, guide, ...), see src/guide.js
    if (!ORGANIZE_STEPS.includes(body.step)) return json({ ok: false, error: `step must be one of ${ORGANIZE_STEPS.join(", ")}` }, 400, {});
    const token = await accessToken(env);
    if (!token) return json({ ok: false, error: "not connected" }, 503, {});
    return json({ ok: true, step: body.step, done: await organizeWorkbook(workbook(token, env), body.step) }, 200, {});
  }
  if (body.action === "menu") { // how fetching the menu goes, hop by hop (no secrets)
    const hops = [];
    let next = env.MENU_URL;
    for (let i = 0; i < 5 && next; i++) {
      try {
        const r = await fetch(next, { redirect: "manual" });
        const loc = r.headers.get("Location");
        hops.push({ url: next, status: r.status, type: r.headers.get("Content-Type"), location: loc });
        next = loc ? new URL(loc, next).toString() : "";
      } catch (err) { hops.push({ url: next, error: String(err && err.message || err) }); next = ""; }
    }
    return json({ ok: true, hops }, 200, {});
  }
  if (body.action === "delete_test_booking") { // removes one conference booking whose student is "TEST..."
    const id = String(body.id || "").trim().toUpperCase();
    if (!/^[A-Z0-9]{5}$/.test(id)) return json({ ok: false, error: "id (the 5-letter booking number) is required" }, 400, {});
    const token = await accessToken(env);
    if (!token) return json({ ok: false, error: "not connected" }, 503, {});
    const r = await deleteTestBooking(workbook(token, env), id);
    return json({ ok: r.deleted, id, ...r }, r.deleted ? 200 : 400, {});
  }
  if (body.action === "delete_test_signup") { // removes one tour sign-up whose parent is "TEST..."
    const id = String(body.id || "").trim().toUpperCase();
    if (!/^[A-Z0-9]{5}$/.test(id)) return json({ ok: false, error: "id (the 5-letter sign-up number) is required" }, 400, {});
    const token = await accessToken(env);
    if (!token) return json({ ok: false, error: "not connected" }, 503, {});
    const r = await deleteTestSignup(workbook(token, env), id);
    return json({ ok: r.deleted, id, ...r }, r.deleted ? 200 : 400, {});
  }
  if (body.action === "delete_test_transcript") { // removes one transcript request whose student is "TEST..."
    const id = String(body.id || "").trim().toUpperCase();
    if (!/^[A-Z0-9]{5}$/.test(id)) return json({ ok: false, error: "id (the 5-letter request number) is required" }, 400, {});
    const token = await accessToken(env);
    if (!token) return json({ ok: false, error: "not connected" }, 503, {});
    const r = await deleteTestRequest(workbook(token, env), id);
    return json({ ok: r.deleted, id, ...r }, r.deleted ? 200 : 400, {});
  }
  if (body.action === "delete_test_inquiry") { // removes one Request Info inquiry whose student is "TEST..."
    const id = String(body.id || "").trim().toUpperCase();
    if (!/^[A-Z0-9]{5}$/.test(id)) return json({ ok: false, error: "id (the 5-letter request number) is required" }, 400, {});
    const token = await accessToken(env);
    if (!token) return json({ ok: false, error: "not connected" }, 503, {});
    const r = await deleteTestInquiry(workbook(token, env), id);
    return json({ ok: r.deleted, id, ...r }, r.deleted ? 200 : 400, {});
  }
  if (body.action === "delete_test_absence") { // removes one absence report whose student is "TEST..."
    const id = String(body.id || "").trim().toUpperCase();
    if (!/^[A-Z0-9]{5}$/.test(id)) return json({ ok: false, error: "id (the 5-letter report number) is required" }, 400, {});
    const token = await accessToken(env);
    if (!token) return json({ ok: false, error: "not connected" }, 503, {});
    const r = await deleteTestAbsence(workbook(token, env), id);
    return json({ ok: r.deleted, id, ...r }, r.deleted ? 200 : 400, {});
  }
  if (body.action === "delete_test_feedback") { // removes one feedback response whose name is "TEST..."
    const id = String(body.id || "").trim().toUpperCase();
    if (!/^[A-Z0-9]{5}$/.test(id)) return json({ ok: false, error: "id (the 5-letter feedback number) is required" }, 400, {});
    const token = await accessToken(env);
    if (!token) return json({ ok: false, error: "not connected" }, 503, {});
    const r = await deleteTestFeedback(workbook(token, env), id);
    return json({ ok: r.deleted, id, ...r }, r.deleted ? 200 : 400, {});
  }
  if (body.action === "delete_test_order") { // removes a lunch order whose rows are all "TEST..."
    const id = String(body.id || "").trim().toUpperCase();
    if (!/^[A-Z0-9]{5}$/.test(id)) return json({ ok: false, error: "id (the 5-letter order number) is required" }, 400, {});
    const token = await accessToken(env);
    if (!token) return json({ ok: false, error: "not connected" }, 503, {});
    const r = await deleteTestOrder(workbook(token, env), env.TABLE || "Orders", id);
    return json({ ok: r.deleted > 0, id, ...r }, r.deleted > 0 ? 200 : 400, {});
  }
  if (body.action === "homerooms") {
    const week = String(body.week || ""), to = String(body.test_to || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(week) || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)) {
      return json({ ok: false, error: "week (YYYY-MM-DD) and test_to are required" }, 400, {});
    }
    return json(await homeroomRun(env, week, to), 200, {});
  }
  return json({ ok: false, error: "Unknown action" }, 400, {});
}

// Removes every row of one lunch order, by order number, only if each of its rows has a Student or
// Parent starting with "TEST" (admin upkeep). Bottom row first, so the row numbers stay right.
export async function deleteTestOrder(call, table, id) {
  const r = await call("GET", `/tables/${encodeURIComponent(table)}/range?$select=text`);
  if (!r.ok) return { deleted: 0, error: `table ${r.status}` };
  const [head, ...rows] = (await r.json()).text || [];
  const col = Object.fromEntries((head || []).map((h, i) => [String(h).trim(), i]));
  const hits = rows.map((row, i) => [row, i]).filter(([row]) => String(row[col["Order #"]] || "").trim() === id);
  if (!hits.length) return { deleted: 0, error: "no order with that number" };
  const test = (v) => /^TEST/i.test(String(v || "").trim());
  if (!hits.every(([row]) => test(row[col.Student]) || test(row[col.Parent]))) return { deleted: 0, error: "not a TEST order" };
  let deleted = 0;
  for (const [, i] of hits.reverse()) {
    let d = await call("DELETE", `/tables/${encodeURIComponent(table)}/rows/itemAt(index=${i})`);
    if (!d.ok) d = await call("DELETE", `/tables/${encodeURIComponent(table)}/rows/${i}`);
    if (!d.ok) return { deleted, error: `delete ${d.status}` };
    deleted++;
  }
  return { deleted };
}

// ------------------------------------------------------------------ orders
async function order(request, env, cors) {
  if (!cors["Access-Control-Allow-Origin"]) return json({ ok: false, error: "Not allowed" }, 403, cors);
  const text = await request.text();
  if (text.length > 20000) return json({ ok: false, error: "Order is too large" }, 413, cors);
  let body;
  try { body = JSON.parse(text); } catch { return json({ ok: false, error: "Bad request" }, 400, cors); }
  if (body.website) return json({ ok: true, order_id: "0", total: 0, email_sent: true }, 200, cors); // spam trap

  const ip = request.headers.get("CF-Connecting-IP") || "unknown";
  if (await tooMany(env, ip)) {
    return json({ ok: false, error: "Too many orders from this connection; please call the office" }, 429, cors);
  }
  const menu = await loadMenu(env);
  const checked = validateOrder(body, menu, Date.now());
  if (checked.error) return json({ ok: false, error: checked.error }, 400, cors);

  const orderId = newOrderId();
  const token = await accessToken(env);
  if (!token) return json({ ok: false, error: "Online ordering isn't connected yet" }, 503, cors);

  const rows = orderRows(checked.order, menu, orderId, new Date());
  const add = await graph(token, "POST",
    `/me/drive/root:/${encodePath(env.WORKBOOK_PATH)}:/workbook/tables/${encodeURIComponent(env.TABLE || "Orders")}/rows/add`,
    { index: null, values: rows });
  if (!add.ok) {
    console.error("workbook", add.status, await add.text());
    return json({ ok: false, error: "We couldn't save your order" }, 502, cors);
  }
  const mail = await graph(token, "POST", "/me/sendMail", confirmationMail(checked.order, menu, orderId, env));
  if (!mail.ok) console.error("mail", mail.status, await mail.text());
  return json({ ok: true, order_id: orderId, total: checked.order.total, email_sent: mail.ok }, 200, cors);
}

export function validateOrder(body, menu, now) {
  if (!menu || !menu.week_start || !menu.deadline || !menu.days || !menu.days.length) {
    return { error: "Online ordering isn't open right now" };
  }
  if (now > Date.parse(menu.deadline)) return { error: "Ordering for this week is closed; please call the office" };
  if (body.week_start && body.week_start !== menu.week_start) {
    return { error: "The menu changed; please reload the page and order again" };
  }
  const str = (v, max) => String(v == null ? "" : v).replace(/[\u0000-\u001f]/g, " ").trim().slice(0, max);
  const parent = {
    first: str(body.parent_first, 60), last: str(body.parent_last, 60),
    email: str(body.email, 120), phone: str(body.phone, 30),
  };
  if (!parent.first || !parent.last) return { error: "Please enter your first and last name" };
  if (!/^[^@\s<>"]+@[^@\s<>"]+\.[^@\s<>"]+$/.test(parent.email)) return { error: "Please enter a valid email address" };
  if (!parent.phone) return { error: "Please enter a phone number" };
  const kids = Array.isArray(body.children) ? body.children : [];
  if (!kids.length) return { error: "Please add a student" };
  if (kids.length > MAX_CHILDREN) return { error: "Too many students in one order" };
  const open = new Map(menu.days.map((d) => [d.key, d]));
  const grades = new Map(menu.grades.map((g) => [g.value, g]));
  const children = [];
  let total = 0;
  for (const k of kids) {
    const name = str(k && k.name, 80);
    const grade = grades.get(str(k && k.grade, 10));
    const days = [...new Set(Array.isArray(k && k.days) ? k.days.map((d) => str(d, 3)) : [])];
    if (!name) return { error: "Please enter each student's name" };
    if (!grade) return { error: `Please choose a grade for ${name}` };
    if (!days.length) return { error: `Please choose at least one day for ${name}` };
    if (days.some((d) => !open.has(d))) return { error: "That day isn't on this week's menu; please reload the page" };
    const price = grade.staff ? menu.prices.staff : grade.preschool ? menu.prices.preschool : menu.prices.regular;
    if (!(price > 0)) return { error: `Please choose a grade for ${name}` };
    const ordered = DAY_KEYS.filter((d) => days.includes(d));
    total += price * ordered.length;
    children.push({ name, grade, price, days: ordered });
  }
  return { order: { parent, children, total: Math.round(total * 100) / 100 } };
}

export function orderRows(order, menu, orderId, now) {
  const parent = `${order.parent.first} ${order.parent.last}`;
  return order.children.map((c) => [
    eastern(now, true), usDate(menu.week_start), orderId, parent, order.parent.email, order.parent.phone,
    c.name, c.grade.label, c.grade.preschool ? "Yes" : "No", c.price,
    ...DAY_KEYS.map((d) => (c.days.includes(d) ? "X" : "")),
    '=COUNTIF(Orders[@[Mon]:[Fri]],"X")', "=[@Lunches]*[@[Price per lunch]]", "",
  ]);
}

export function confirmationMail(order, menu, orderId, env) {
  const e = escapeHtml;
  const byKey = new Map(menu.days.map((d) => [d.key, d]));
  const kids = order.children.map((c) => {
    const lines = c.days.map((d) => {
      const day = byKey.get(d);
      return `<li>${e(day.day)} ${e(day.date_label)}: ${e(day.meal)}</li>`;
    }).join("");
    return `<p style="margin:12px 0 4px"><b>${e(c.name)}</b> (${e(c.grade.label)}), ${c.days.length} lunch` +
      `${c.days.length === 1 ? "" : "es"} at ${money(c.price)}: ${money(c.price * c.days.length)}</p><ul style="margin:0">${lines}</ul>`;
  }).join("");
  const content =
    `<p>Dear ${e(order.parent.first)} ${e(order.parent.last)},</p>` +
    `<p>Thank you for your hot lunch order for the week of ${e(menu.week_label)}. Here is what we received:</p>` +
    kids +
    `<p style="margin-top:14px"><b>Total: ${money(order.total)}</b> (order number ${e(orderId)})</p>` +
    `<p>Lunch charges are added to your Gradelink account. To change or cancel this order, please call the ` +
    `school office at (301) 856-1616 or reply to this email.</p>` +
    `<p>Gratefully,<br>${e(env.SIGNATURE || "Independent Baptist Academy").replace(/\n/g, "<br>")}</p>`;
  const message = {
    subject: `Hot lunch order received: week of ${menu.week_label}`,
    body: { contentType: "HTML", content },
    toRecipients: [{ emailAddress: { address: order.parent.email, name: `${order.parent.first} ${order.parent.last}` } }],
  };
  if (env.REPLY_TO) message.replyTo = [{ emailAddress: { address: env.REPLY_TO } }];
  return { message, saveToSentItems: true };
}

// ------------------------------------------------------------------ conferences
async function conferences(env, cors) {
  if (!cors["Access-Control-Allow-Origin"]) return json({ ok: false, error: "Not allowed" }, 403, cors);
  const schedule = await loadJson(env.CONF_URL);
  const token = await accessToken(env);
  if (!token) return json({ ok: false, error: "Online sign-up isn't connected yet" }, 503, cors);
  const bookings = await bookingsOrTab(workbook(token, env));
  return json({ ok: true, date: schedule.date, taken: takenSlots(schedule, bookings) }, 200, cors);
}

async function conference(request, env, cors) {
  if (!cors["Access-Control-Allow-Origin"]) return json({ ok: false, error: "Not allowed" }, 403, cors);
  const text = await request.text();
  if (text.length > 5000) return json({ ok: false, error: "Too large" }, 413, cors);
  let body;
  try { body = JSON.parse(text); } catch { return json({ ok: false, error: "Bad request" }, 400, cors); }
  if (body.website) return json({ ok: true, booking_id: "0", time: "", email_sent: true }, 200, cors); // spam trap

  const ip = request.headers.get("CF-Connecting-IP") || "unknown";
  if (await tooMany(env, ip, "rlc")) {
    return json({ ok: false, error: "Too many sign-ups from this connection; please call the office" }, 429, cors);
  }
  const schedule = await loadJson(env.CONF_URL);
  const checked = validateBooking(body, schedule, Date.now());
  if (checked.error) return json({ ok: false, error: checked.error }, 400, cors);
  const b = checked.booking;
  const token = await accessToken(env);
  if (!token) return json({ ok: false, error: "Online sign-up isn't connected yet" }, 503, cors);

  // Two parents tapping the same time at once: the first one holds it for a minute while it is saved.
  const lock = `conf-lock:${schedule.date}:${b.teacher.id}:${b.slot.id}`;
  const taken = { ok: false, code: "taken", error: `Someone just booked ${b.slot.label} with ${b.teacher.name}. Please pick another time.` };
  if (await env.LUNCH_KV.get(lock)) return json(taken, 409, cors);
  await env.LUNCH_KV.put(lock, "1", { expirationTtl: 60 });

  const call = workbook(token, env);
  const clash = conflict(b, schedule, await bookingsOrTab(call));
  if (clash) {
    if (clash.code !== "taken") await env.LUNCH_KV.delete(lock);
    return json({ ok: false, ...clash }, clash.code === "taken" ? 409 : 400, cors);
  }
  const id = newOrderId();
  const add = await call("POST", `/tables/${CONF_TABLE}/rows/add`,
    { index: null, values: [bookingRow(b, schedule, id, eastern(new Date(), true))] });
  if (!add.ok) {
    console.error("conference workbook", add.status, await add.text());
    await env.LUNCH_KV.delete(lock);
    return json({ ok: false, error: "We couldn't save your booking" }, 502, cors);
  }
  const mail = await graph(token, "POST", "/me/sendMail", parentMail(b, schedule, id, env));
  if (!mail.ok) console.error("conference mail", mail.status, await mail.text());
  try {
    // Every teacher in the panel (all three for High School) whose email is on the Teachers tab.
    const teachers = await readTeachers(call);
    for (const name of b.teacher.members && b.teacher.members.length ? b.teacher.members : [b.teacher.name]) {
      const to = teacherEmail(teachers, name);
      if (!to) continue;
      const t = await graph(token, "POST", "/me/sendMail", teacherMail(b, schedule, id, to, name, env));
      if (!t.ok) console.error("teacher mail", t.status, await t.text());
    }
  } catch (err) { console.error("teacher mail", err && err.message); }
  return json({ ok: true, booking_id: id, time: b.slot.label, email_sent: mail.ok }, 200, cors);
}

// The bookings, adding the Conferences tab first if the workbook doesn't have it yet.
async function bookingsOrTab(call) {
  const rows = await readBookings(call);
  if (rows) return rows;
  await ensureConferenceTab(call);
  return [];
}

// ------------------------------------------------------------------ school tours
async function tours(env, cors) {
  if (!cors["Access-Control-Allow-Origin"]) return json({ ok: false, error: "Not allowed" }, 403, cors);
  const schedule = await loadJson(env.TOUR_URL);
  const token = await accessToken(env);
  if (!token) return json({ ok: false, error: "Online sign-up isn't connected yet" }, 503, cors);
  return json({ ok: true, taken: takenCounts(schedule, await signupsOrTab(workbook(token, env))) }, 200, cors);
}

async function tour(request, env, cors) {
  if (!cors["Access-Control-Allow-Origin"]) return json({ ok: false, error: "Not allowed" }, 403, cors);
  const text = await request.text();
  if (text.length > 5000) return json({ ok: false, error: "Too large" }, 413, cors);
  let body;
  try { body = JSON.parse(text); } catch { return json({ ok: false, error: "Bad request" }, 400, cors); }
  if (body.website) return json({ ok: true, signup_id: "0", email_sent: true }, 200, cors); // spam trap

  const ip = request.headers.get("CF-Connecting-IP") || "unknown";
  if (await tooMany(env, ip, "rlt")) {
    return json({ ok: false, error: "Too many sign-ups from this connection; please call the office" }, 429, cors);
  }
  const schedule = await loadJson(env.TOUR_URL);
  const checked = validateSignup(body, schedule, Date.now());
  if (checked.error) return json({ ok: false, error: checked.error }, 400, cors);
  const s = checked.signup;
  const token = await accessToken(env);
  if (!token) return json({ ok: false, error: "Online sign-up isn't connected yet" }, 503, cors);

  // One sign-up at a time per tour, so the last spot can't go to two families.
  const lock = `tour-lock:${s.tour.id}`;
  if (await env.LUNCH_KV.get(lock)) {
    return json({ ok: false, code: "busy", error: "Another family is signing up for this tour right now. Please try again in a minute." }, 409, cors);
  }
  await env.LUNCH_KV.put(lock, "1", { expirationTtl: 60 });
  try {
    const call = workbook(token, env);
    const signups = await signupsOrTab(call);
    const clash = tourConflict(s, signups);
    if (clash) return json({ ok: false, ...clash }, clash.code === "full" ? 409 : 400, cors);
    const id = newOrderId();
    const add = await call("POST", `/tables/${TOUR_TABLE}/rows/add`,
      { index: null, values: [signupRow(s, id, eastern(new Date(), true))] });
    if (!add.ok) {
      console.error("tour workbook", add.status, await add.text());
      return json({ ok: false, error: "We couldn't save your sign-up" }, 502, cors);
    }
    const count = takenCounts(schedule, signups)[s.tour.id] + 1 || 1;
    const mail = await graph(token, "POST", "/me/sendMail", familyMail(s, id, env));
    if (!mail.ok) console.error("tour mail", mail.status, await mail.text());
    if (env.TOUR_NOTIFY) {
      const o = await graph(token, "POST", "/me/sendMail", officeMail(s, id, count, env));
      if (!o.ok) console.error("tour office mail", o.status, await o.text());
    }
    return json({ ok: true, signup_id: id, email_sent: mail.ok }, 200, cors);
  } finally {
    await env.LUNCH_KV.delete(lock);
  }
}

// The sign-ups, adding the Tours tab first if the workbook doesn't have it yet.
async function signupsOrTab(call) {
  const rows = await readSignups(call);
  if (rows) return rows;
  await ensureTourTab(call);
  return [];
}

// ------------------------------------------------------------------ Request Info inquiries
async function inquiry(request, env, cors) {
  if (!cors["Access-Control-Allow-Origin"]) return json({ ok: false, error: "Not allowed" }, 403, cors);
  const text = await request.text();
  if (text.length > 6000) return json({ ok: false, error: "Too large" }, 413, cors);
  let body;
  try { body = JSON.parse(text); } catch { return json({ ok: false, error: "Bad request" }, 400, cors); }
  if (body.website) return json({ ok: true, request_id: "0", email_sent: true }, 200, cors); // spam trap

  const ip = request.headers.get("CF-Connecting-IP") || "unknown";
  if (await tooMany(env, ip, "rlin")) {
    return json({ ok: false, error: "Too many requests from this connection; please call the office" }, 429, cors);
  }
  const checked = validateInquiry(body, Date.now());
  if (checked.error) return json({ ok: false, error: checked.error }, 400, cors);
  const q = checked.inquiry;
  const token = await accessToken(env);
  if (!token) return json({ ok: false, error: "Online requests aren't connected yet" }, 503, cors);

  const call = workbook(token, env);
  const id = newOrderId();
  const row = { index: null, values: [inquiryRow(q, id, eastern(new Date(), true))] };
  let add = await call("POST", `/tables/${INQUIRY_TABLE}/rows/add`, row);
  if (add.status === 404 && await ensureInquiryTab(call)) add = await call("POST", `/tables/${INQUIRY_TABLE}/rows/add`, row);
  if (!add.ok) {
    console.error("inquiry workbook", add.status, await add.text());
    return json({ ok: false, error: "We couldn't save your request" }, 502, cors);
  }
  const mail = await graph(token, "POST", "/me/sendMail", inquiryFamilyMail(q, id, env));
  if (!mail.ok) console.error("inquiry mail", mail.status, await mail.text());
  if (env.INQUIRY_NOTIFY) {
    const o = await graph(token, "POST", "/me/sendMail", inquiryOfficeMail(q, id, env));
    if (!o.ok) console.error("inquiry office mail", o.status, await o.text());
  }
  return json({ ok: true, request_id: id, email_sent: mail.ok }, 200, cors);
}

// ------------------------------------------------------------------ Report Student Absence
async function absence(request, env, cors) {
  if (!cors["Access-Control-Allow-Origin"]) return json({ ok: false, error: "Not allowed" }, 403, cors);
  const text = await request.text();
  if (text.length > 6000) return json({ ok: false, error: "Too large" }, 413, cors);
  let body;
  try { body = JSON.parse(text); } catch { return json({ ok: false, error: "Bad request" }, 400, cors); }
  if (body.website) return json({ ok: true, absence_id: "0", email_sent: true }, 200, cors); // spam trap

  const ip = request.headers.get("CF-Connecting-IP") || "unknown";
  if (await tooMany(env, ip, "rlab")) {
    return json({ ok: false, error: "Too many reports from this connection; please call the office" }, 429, cors);
  }
  const checked = validateAbsence(body, Date.now());
  if (checked.error) return json({ ok: false, error: checked.error }, 400, cors);
  const q = checked.absence;
  const token = await accessToken(env);
  if (!token) return json({ ok: false, error: "Online absence reports aren't connected yet" }, 503, cors);

  const call = workbook(token, env);
  let teacher = null;
  try { teacher = homeroomFor(await readTeachers(call), q.grade); } catch (err) { console.error("absence teachers", err); }
  const id = newOrderId();
  const row = { index: null, values: [absenceRow(q, id, eastern(new Date(), true), teacher)] };
  let add = await call("POST", `/tables/${ABSENCE_TABLE}/rows/add`, row);
  if (add.status === 404 && await ensureAbsenceTab(call)) add = await call("POST", `/tables/${ABSENCE_TABLE}/rows/add`, row);
  if (!add.ok) {
    console.error("absence workbook", add.status, await add.text());
    return json({ ok: false, error: "We couldn't save your report" }, 502, cors);
  }
  const mail = await graph(token, "POST", "/me/sendMail", absenceParentMail(q, id, env));
  if (!mail.ok) console.error("absence mail", mail.status, await mail.text());
  if (env.ABSENCE_NOTIFY) {
    const o = await graph(token, "POST", "/me/sendMail", absenceOfficeMail(q, id, env, teacher));
    if (!o.ok) console.error("absence office mail", o.status, await o.text());
  }
  return json({ ok: true, absence_id: id, email_sent: mail.ok }, 200, cors);
}

// ------------------------------------------------------------------ Feedback for Administration
async function feedback(request, env, cors) {
  if (!cors["Access-Control-Allow-Origin"]) return json({ ok: false, error: "Not allowed" }, 403, cors);
  const text = await request.text();
  if (text.length > 8000) return json({ ok: false, error: "Too large" }, 413, cors);
  let body;
  try { body = JSON.parse(text); } catch { return json({ ok: false, error: "Bad request" }, 400, cors); }
  if (body.website) return json({ ok: true, feedback_id: "0" }, 200, cors); // spam trap

  const ip = request.headers.get("CF-Connecting-IP") || "unknown";
  if (await tooMany(env, ip, "rlfb")) {
    return json({ ok: false, error: "Too many responses from this connection; please try again later" }, 429, cors);
  }
  const checked = validateFeedback(body);
  if (checked.error) return json({ ok: false, error: checked.error }, 400, cors);
  const token = await accessToken(env);
  if (!token) return json({ ok: false, error: "Online feedback isn't connected yet" }, 503, cors);

  const call = workbook(token, env);
  const id = newOrderId();
  const row = { index: null, values: [feedbackRow(checked.feedback, id, eastern(new Date(), true))] };
  let add = await call("POST", `/tables/${FEEDBACK_TABLE}/rows/add`, row);
  if (add.status === 404 && await ensureFeedbackTab(call)) add = await call("POST", `/tables/${FEEDBACK_TABLE}/rows/add`, row);
  if (!add.ok) {
    console.error("feedback workbook", add.status, await add.text());
    return json({ ok: false, error: "We couldn't save your feedback" }, 502, cors);
  }
  return json({ ok: true, feedback_id: id }, 200, cors);
}

// ------------------------------------------------------------------ transcript requests
function transcripts(env, cors) {
  if (!cors["Access-Control-Allow-Origin"]) return json({ ok: false, error: "Not allowed" }, 403, cors);
  return json({ ok: true, ...transcriptSettings(env) }, 200, cors);
}

async function transcript(request, env, cors) {
  if (!cors["Access-Control-Allow-Origin"]) return json({ ok: false, error: "Not allowed" }, 403, cors);
  const text = await request.text();
  if (text.length > 8000) return json({ ok: false, error: "Too large" }, 413, cors);
  let body;
  try { body = JSON.parse(text); } catch { return json({ ok: false, error: "Bad request" }, 400, cors); }
  if (body.website) return json({ ok: true, request_id: "0", email_sent: true }, 200, cors); // spam trap

  const ip = request.headers.get("CF-Connecting-IP") || "unknown";
  if (await tooMany(env, ip, "rltr")) {
    return json({ ok: false, error: "Too many requests from this connection; please call the office" }, 429, cors);
  }
  const checked = validateRequest(body, Date.now());
  if (checked.error) return json({ ok: false, error: checked.error }, 400, cors);
  const q = checked.request;
  const token = await accessToken(env);
  if (!token) return json({ ok: false, error: "Online requests aren't connected yet" }, 503, cors);

  const call = workbook(token, env);
  const id = newOrderId();
  const row = { index: null, values: [requestRow(q, id, eastern(new Date(), true), env)] };
  let add = await call("POST", `/tables/${TRANSCRIPT_TABLE}/rows/add`, row);
  if (add.status === 404 && await ensureTranscriptTab(call)) add = await call("POST", `/tables/${TRANSCRIPT_TABLE}/rows/add`, row);
  if (!add.ok) {
    console.error("transcript workbook", add.status, await add.text());
    return json({ ok: false, error: "We couldn't save your request" }, 502, cors);
  }
  const mail = await graph(token, "POST", "/me/sendMail", requesterMail(q, id, env));
  if (!mail.ok) console.error("transcript mail", mail.status, await mail.text());
  if (env.TRANSCRIPT_NOTIFY) {
    const o = await graph(token, "POST", "/me/sendMail", transcriptOfficeMail(q, id, env));
    if (!o.ok) console.error("transcript office mail", o.status, await o.text());
  }
  return json({ ok: true, request_id: id, email_sent: mail.ok, ...transcriptSettings(env) }, 200, cors);
}

async function loadJson(url) {
  // As for the menu: the minute in the address skips GitHub's 5-minute cache.
  const r = await fetch(`${url}${url.includes("?") ? "&" : "?"}v=${Math.floor(Date.now() / 60000)}`, { cf: { cacheTtl: 30 } });
  if (!r.ok) throw new Error(`${url} ${r.status}`);
  return r.json();
}

async function loadMenu(env) {
  // The minute in the address skips GitHub's 5-minute cache, so a new menu counts within a minute.
  const url = `${env.MENU_URL}${env.MENU_URL.includes("?") ? "&" : "?"}v=${Math.floor(Date.now() / 60000)}`;
  const r = await fetch(url, { cf: { cacheTtl: 30 } });
  if (!r.ok) throw new Error(`menu ${r.status}`);
  return r.json();
}

async function tooMany(env, ip, prefix = "rl") {
  const key = `${prefix}:${ip}`;
  const n = parseInt((await env.LUNCH_KV.get(key)) || "0", 10);
  if (n >= 10) return true;
  await env.LUNCH_KV.put(key, String(n + 1), { expirationTtl: 3600 });
  return false;
}

// ------------------------------------------------------------------ Microsoft sign-in
let cached = { token: "", expires: 0 }; // access token for this Worker instance

function redirectUri(url) { return `${url.origin}/callback`; }

async function connect(url, env) {
  if (!env.SETUP_KEY || url.searchParams.get("key") !== env.SETUP_KEY) return page("Not allowed.", 403);
  const state = newOrderId() + newOrderId();
  await env.LUNCH_KV.put(`state:${state}`, "1", { expirationTtl: 600 });
  const auth = new URL(`https://login.microsoftonline.com/${env.TENANT_ID}/oauth2/v2.0/authorize`);
  auth.search = new URLSearchParams({
    client_id: env.CLIENT_ID, response_type: "code", redirect_uri: redirectUri(url),
    response_mode: "query", scope: SCOPES, state, login_hint: env.MAILBOX, prompt: "select_account",
  });
  return Response.redirect(auth.toString(), 302);
}

async function callback(url, env) {
  const state = url.searchParams.get("state") || "";
  if (!state || !(await env.LUNCH_KV.get(`state:${state}`))) return page("This sign-in link expired. Please start again.", 400);
  await env.LUNCH_KV.delete(`state:${state}`);
  const code = url.searchParams.get("code");
  if (!code) return page(`Microsoft sign-in didn't finish: ${escapeHtml(url.searchParams.get("error_description") || "no code")}`, 400);
  const tok = await tokenRequest(env, { grant_type: "authorization_code", code, redirect_uri: redirectUri(url) });
  if (!tok.access_token || !tok.refresh_token) return page("Microsoft sign-in failed. Please try again.", 400);
  const me = await (await graph(tok.access_token, "GET", "/me?$select=userPrincipalName,mail")).json();
  const who = String(me.userPrincipalName || me.mail || "").toLowerCase();
  if (who !== String(env.MAILBOX).toLowerCase()) {
    return page(`Signed in as ${escapeHtml(who)}, but online ordering must be connected with ${escapeHtml(env.MAILBOX)}. ` +
      "Please sign out of Microsoft and try again.", 403);
  }
  await env.LUNCH_KV.put("ms_refresh", tok.refresh_token);
  cached = { token: "", expires: 0 };
  return page(`Connected. Orders and sign-ups will be saved to ${escapeHtml(env.WORKBOOK_PATH)} and confirmations ` +
    `sent from ${escapeHtml(env.MAILBOX)}. You can close this page.`);
}

async function accessToken(env) {
  if (cached.token && Date.now() < cached.expires) return cached.token;
  const refresh = await env.LUNCH_KV.get("ms_refresh");
  if (!refresh) return "";
  const tok = await tokenRequest(env, { grant_type: "refresh_token", refresh_token: refresh, scope: SCOPES });
  if (!tok.access_token) {
    console.error("token refresh failed", tok.error, tok.error_description);
    return "";
  }
  if (tok.refresh_token && tok.refresh_token !== refresh) await env.LUNCH_KV.put("ms_refresh", tok.refresh_token);
  cached = { token: tok.access_token, expires: Date.now() + (Number(tok.expires_in || 3600) - 300) * 1000 };
  return cached.token;
}

async function tokenRequest(env, params) {
  const r = await fetch(`https://login.microsoftonline.com/${env.TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: env.CLIENT_ID, client_secret: env.CLIENT_SECRET, ...params }),
  });
  return r.json();
}

function graph(token, method, path, body) {
  return fetch(GRAPH + path, {
    method,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
}

// ------------------------------------------------------------------ helpers
function corsHeaders(request, env) {
  const origin = request.headers.get("Origin") || "";
  const allowed = String(env.ALLOWED_ORIGINS || "").split(",").map((s) => s.trim()).filter(Boolean);
  const h = { "Access-Control-Allow-Methods": "GET, POST, OPTIONS", "Access-Control-Allow-Headers": "Content-Type", Vary: "Origin" };
  if (allowed.includes(origin)) h["Access-Control-Allow-Origin"] = origin;
  return h;
}

function json(data, status, headers) {
  return new Response(JSON.stringify(data), { status, headers: { ...headers, "Content-Type": "application/json" } });
}

function page(message, status = 200) {
  return new Response(`<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">` +
    `<title>Hot lunch ordering</title><body style="font:17px/1.5 system-ui,sans-serif;max-width:36em;margin:3em auto;padding:0 1em">` +
    `<h1 style="font-size:1.4em">Hot lunch ordering</h1><p>${message}</p></body>`,
  { status, headers: { "Content-Type": "text/html; charset=utf-8" } });
}

function newOrderId() {
  const a = new Uint8Array(5);
  crypto.getRandomValues(a);
  return Array.from(a, (b) => "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"[b % 32]).join("");
}

function encodePath(p) { return String(p).split("/").map(encodeURIComponent).join("/"); }

export function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function money(n) { return "$" + Number(n).toFixed(2); }

function usDate(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return `${m}/${d}/${y}`;
}

function eastern(date, withTime) {
  const parts = Object.fromEntries(new Intl.DateTimeFormat("en-US", {
    timeZone: "America/New_York", year: "numeric", month: "numeric", day: "numeric",
    hour: "numeric", minute: "2-digit", hour12: true,
  }).formatToParts(date).map((p) => [p.type, p.value]));
  const day = `${parts.month}/${parts.day}/${parts.year}`;
  return withTime ? `${day} ${parts.hour}:${parts.minute} ${parts.dayPeriod}` : day;
}
