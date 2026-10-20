// Run: node --test worker/test/worker.test.mjs  (no network: Microsoft and the menu are faked)
import { test } from "node:test";
import assert from "node:assert/strict";
import worker, { validateOrder, orderRows, confirmationMail } from "../src/index.js";

const MENU = {
  week_start: "2026-10-19", week_label: "October 19 – 23",
  deadline: "2026-10-16T19:00:00+00:00", deadline_label: "Friday, October 16 at 3:00 PM",
  prices: { regular: 8, preschool: 5 }, note: "",
  days: [
    { key: "mon", day: "Monday", meal: "Chicken parmesan", date: "2026-10-19", date_label: "Oct 19" },
    { key: "wed", day: "Wednesday", meal: "Hot dogs <b>", date: "2026-10-21", date_label: "Oct 21" },
  ],
  grades: [{ value: "K4", label: "Preschool (K4)", preschool: true }, { value: "3", label: "3rd Grade", preschool: false }],
};
const BEFORE = Date.parse("2026-10-15T12:00:00Z");
const ORDER = {
  parent_first: "Pat", parent_last: "Smith", email: "pat@example.com", phone: "301-555-0100", week_start: "2026-10-19",
  children: [{ name: "Sam Smith", grade: "3", days: ["mon", "wed"] }, { name: "Ally Smith", grade: "K4", days: ["wed"] }],
};
const ENV0 = {
  MENU_URL: "https://site.test/lunch/menu.json", ALLOWED_ORIGINS: "https://guarantee418.github.io",
  MAILBOX: "ibs@ibaclinton.com", REPLY_TO: "ibs@ibcministries.org", SIGNATURE: "Mrs. Napierala\nIndependent Baptist <Academy>",
  WORKBOOK_PATH: "IBA Portal Forms.xlsx", TABLE: "Orders", TENANT_ID: "tenant", CLIENT_ID: "client",
  CLIENT_SECRET: "secret", SETUP_KEY: "setup",
};

function kv(initial = {}) {
  const m = new Map(Object.entries(initial));
  return { get: async (k) => (m.has(k) ? m.get(k) : null), put: async (k, v) => { m.set(k, v); }, delete: async (k) => { m.delete(k); }, m };
}

test("valid order: prices from the menu, days in week order", () => {
  const r = validateOrder({ ...ORDER, children: [{ name: "Sam", grade: "3", days: ["wed", "mon", "mon"] }] }, MENU, BEFORE);
  assert.equal(r.error, undefined);
  assert.deepEqual(r.order.children[0].days, ["mon", "wed"]);
  assert.equal(r.order.total, 16);
  assert.equal(validateOrder(ORDER, MENU, BEFORE).order.total, 21); // 2 x $8 + 1 x $5
});

test("rejects late orders, unknown days and grades, bad email, stale menus", () => {
  assert.match(validateOrder(ORDER, MENU, Date.parse("2026-10-16T19:00:01Z")).error, /closed/);
  assert.match(validateOrder({ ...ORDER, children: [{ name: "Sam", grade: "3", days: ["tue"] }] }, MENU, BEFORE).error, /isn't on this week/);
  assert.match(validateOrder({ ...ORDER, children: [{ name: "Sam", grade: "13", days: ["mon"] }] }, MENU, BEFORE).error, /grade/);
  assert.match(validateOrder({ ...ORDER, children: [{ name: "Sam", grade: "3", days: [] }] }, MENU, BEFORE).error, /at least one day/);
  assert.match(validateOrder({ ...ORDER, email: "nope" }, MENU, BEFORE).error, /email/);
  assert.match(validateOrder({ ...ORDER, week_start: "2026-10-12" }, MENU, BEFORE).error, /menu changed/);
  assert.match(validateOrder(ORDER, { ...MENU, days: [] }, BEFORE).error, /isn't open/);
  assert.match(validateOrder({ ...ORDER, children: Array(9).fill(ORDER.children[0]) }, MENU, BEFORE).error, /Too many/);
});

test("workbook rows: one per student, X per day, formulas for count and total", () => {
  const rows = orderRows(validateOrder(ORDER, MENU, BEFORE).order, MENU, "ABCDE", new Date("2026-10-15T18:05:00Z"));
  assert.equal(rows.length, 2);
  assert.deepEqual(rows[0].slice(0, 15), ["10/15/2026 2:05 PM", "10/19/2026", "ABCDE", "Pat Smith", "pat@example.com",
    "301-555-0100", "Sam Smith", "3rd Grade", "No", 8, "X", "", "X", "", ""]);
  assert.equal(rows[1][8], "Yes");
  assert.equal(rows[1][9], 5);
  assert.equal(rows[0].length, 18);
});

test("confirmation email escapes everything typed in and replies to the office", () => {
  const order = validateOrder({ ...ORDER, parent_first: "<script>" }, MENU, BEFORE).order;
  const mail = confirmationMail(order, MENU, "ABCDE", ENV0);
  const html = mail.message.body.content;
  assert.ok(!html.includes("<script>") && html.includes("&lt;script&gt;"));
  assert.ok(html.includes("Hot dogs &lt;b&gt;") && html.includes("$21.00") && html.includes("ABCDE"));
  assert.equal(mail.message.replyTo[0].emailAddress.address, "ibs@ibcministries.org");
  assert.equal(mail.message.toRecipients[0].emailAddress.address, "pat@example.com");
  assert.equal(mail.saveToSentItems, true);
  assert.ok(html.includes("Mrs. Napierala<br>Independent Baptist &lt;Academy&gt;"));
});

function fakeMicrosoft({ workbookOk = true, me = "ibs@ibaclinton.com" } = {}) {
  const calls = [];
  globalThis.fetch = async (url, init = {}) => {
    url = String(url); calls.push({ url, init });
    const ok = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { "Content-Type": "application/json" } });
    if (url.split("?")[0] === ENV0.MENU_URL) return ok(MENU);
    if (url.includes("/oauth2/v2.0/token")) return ok({ access_token: "at", refresh_token: "rt2", expires_in: 3600 });
    if (url.includes("/workbook/tables/Orders/rows/add")) return workbookOk ? ok({}, 201) : ok({ error: "locked" }, 409);
    if (url.endsWith("/me/sendMail")) return new Response(null, { status: 202 });
    if (url.includes("/me?")) return ok({ userPrincipalName: me });
    return ok({ error: "unexpected " + url }, 500);
  };
  return calls;
}

function post(body, origin = "https://guarantee418.github.io") {
  return new Request("https://lunch.test/order", {
    method: "POST", body: JSON.stringify(body),
    headers: { "Content-Type": "application/json", Origin: origin, "CF-Connecting-IP": "1.2.3.4" },
  });
}

test("POST /order end to end: saves rows, sends mail, rotates the token", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: BEFORE });
  const calls = fakeMicrosoft();
  const env = { ...ENV0, LUNCH_KV: kv({ ms_refresh: "rt1" }) };
  const res = await worker.fetch(post(ORDER), env);
  const data = await res.json();
  assert.equal(res.status, 200, JSON.stringify(data));
  assert.equal(data.ok, true); assert.equal(data.total, 21); assert.equal(data.email_sent, true);
  assert.equal(res.headers.get("Access-Control-Allow-Origin"), "https://guarantee418.github.io");
  const add = calls.find((c) => c.url.includes("rows/add"));
  assert.ok(add.url.includes("/me/drive/root:/IBA%20Portal%20Forms.xlsx:/workbook/tables/Orders/rows/add"));
  assert.equal(JSON.parse(add.init.body).values.length, 2);
  assert.ok(calls.some((c) => c.url.endsWith("/me/sendMail")));
  assert.equal(await env.LUNCH_KV.get("ms_refresh"), "rt2");
});

test("POST /order refuses other websites, traps spam, and reports a workbook failure", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: BEFORE });
  fakeMicrosoft({ workbookOk: false });
  const env = { ...ENV0, LUNCH_KV: kv({ ms_refresh: "rt1" }) };
  assert.equal((await worker.fetch(post(ORDER, "https://evil.example"), env)).status, 403);
  const spam = await (await worker.fetch(post({ ...ORDER, website: "x" }), env)).json();
  assert.equal(spam.ok, true);
  const fail = await worker.fetch(post(ORDER), env);
  assert.equal(fail.status, 502);
  assert.match((await fail.json()).error, /couldn't save/);
});

test("/connect needs the setup key; /callback only accepts the office mailbox", async () => {
  fakeMicrosoft({ me: "someone.else@ibaclinton.com" });
  const env = { ...ENV0, LUNCH_KV: kv() };
  assert.equal((await worker.fetch(new Request("https://lunch.test/connect?key=wrong"), env)).status, 403);
  const go = await worker.fetch(new Request("https://lunch.test/connect?key=setup"), env);
  assert.equal(go.status, 302);
  const auth = new URL(go.headers.get("Location"));
  assert.equal(auth.searchParams.get("redirect_uri"), "https://lunch.test/callback");
  assert.equal(auth.searchParams.get("login_hint"), "ibs@ibaclinton.com");
  const state = auth.searchParams.get("state");
  const cb = await worker.fetch(new Request(`https://lunch.test/callback?code=c&state=${state}`), env);
  assert.equal(cb.status, 403);
  assert.equal(await env.LUNCH_KV.get("ms_refresh"), null);
  // Right mailbox, fresh state
  fakeMicrosoft();
  const go2 = await worker.fetch(new Request("https://lunch.test/connect?key=setup"), env);
  const state2 = new URL(go2.headers.get("Location")).searchParams.get("state");
  const ok = await worker.fetch(new Request(`https://lunch.test/callback?code=c&state=${state2}`), env);
  assert.equal(ok.status, 200);
  assert.equal(await env.LUNCH_KV.get("ms_refresh"), "rt2");
  // A used state can't be replayed
  assert.equal((await worker.fetch(new Request(`https://lunch.test/callback?code=c&state=${state2}`), env)).status, 400);
});

// ------------------------------------------------------------------ homeroom lunch lists
import { excelDate, homeroomMails, mondayToSend, readWeek, ensureTabs } from "../src/homerooms.js";

const HEAD = ["Submitted", "Week of", "Order #", "Parent", "Email", "Phone", "Student", "Grade", "Preschool",
  "Price per lunch", "Mon", "Tue", "Wed", "Thu", "Fri", "Lunches", "Total", "Notes"];
const row = (week, student, grade, days) => ["", week, "A", "P", "p@x", "1", student, grade, "No", 8,
  ...["Mon", "Tue", "Wed", "Thu", "Fri"].map((d) => (days.includes(d) ? "X" : "")), 0, 0, ""];
const TEACHERS = [
  { grade: "3rd Grade", teacher: "Mrs. Lundberg", email: "lund@school.test" },
  { grade: "7th Grade", teacher: "Mr. Napierala", email: "nap@school.test" },
  { grade: "8th Grade", teacher: "Mr. Napierala", email: "NAP@school.test" },
  { grade: "1st Grade", teacher: "Miss Ryans", email: "" },
];
const fakeCall = (values) => async () => new Response(JSON.stringify({ values }), { status: 200 });

test("readWeek keeps one week, merges a student ordered twice, reads Excel dates", async () => {
  const kids = await readWeek(fakeCall([HEAD,
    row(46321, "Sam <b>", "3rd Grade", ["Mon"]),            // 10/26/2026 as an Excel day number
    row("10/26/2026", "Sam <b>", "3rd Grade", ["Wed"]),     // same student, second order
    row(46314, "Old Week", "3rd Grade", ["Mon"]),           // 10/19
    row(46321, "Nobody", "3rd Grade", []),
    row(null, "", "", [])]), "2026-10-26");
  assert.equal(kids.length, 1);
  assert.deepEqual([...kids[0].days].sort(), ["Mon", "Wed"]);
  assert.equal(excelDate(46321), "2026-10-26");
});

test("homeroom emails: one per teacher, Cc the office, escaped, none without email or orders", () => {
  const students = [
    { name: "Zed <b>", grade: "8th Grade", days: new Set(["Mon"]) },
    { name: "Amy", grade: "7th Grade", days: new Set(["Mon", "Fri"]) },
    { name: "Lu", grade: "1st Grade", days: new Set(["Tue"]) },
  ];
  const env = { ...ENV0, HOMEROOM_CC: "ibs@ibcministries.org" };
  const mails = homeroomMails(students, TEACHERS, "2026-10-26", env);
  assert.equal(mails.length, 1); // Lundberg has no orders, Ryans has no email
  const m = mails[0].mail.message;
  assert.equal(m.toRecipients[0].emailAddress.address, "nap@school.test");
  assert.equal(m.ccRecipients[0].emailAddress.address, "ibs@ibcministries.org");
  assert.equal(m.subject, "Hot lunch this week: 7th & 8th (week of October 26)");
  const html = m.body.content;
  assert.ok(html.includes("Zed &lt;b&gt;") && !html.includes("Zed <b>"));
  assert.ok(html.indexOf("Amy") < html.indexOf("Zed")); // 7th before 8th
  // Test mode: everything to the tester, no Cc, teachers without email included
  const test = homeroomMails(students, TEACHERS, "2026-10-26", env, "boss@school.test");
  assert.equal(test.length, 2);
  assert.ok(test.every((t) => t.mail.message.toRecipients[0].emailAddress.address === "boss@school.test"));
  assert.ok(test.every((t) => !t.mail.message.ccRecipients && t.mail.message.subject.startsWith("[TEST]")));
});

test("Monday 7 AM Eastern only, in summer and winter time", () => {
  assert.equal(mondayToSend(new Date("2026-10-26T11:00:00Z")), "2026-10-26"); // EDT
  assert.equal(mondayToSend(new Date("2026-10-26T12:00:00Z")), "");
  assert.equal(mondayToSend(new Date("2026-11-02T12:00:00Z")), "2026-11-02"); // EST
  assert.equal(mondayToSend(new Date("2026-11-02T11:00:00Z")), "");
  assert.equal(mondayToSend(new Date("2026-10-27T11:00:00Z")), "");
});

test("ensureTabs adds only missing tabs", async () => {
  const calls = [];
  const call = async (method, path, body) => {
    calls.push({ method, path, body });
    if (path.startsWith("/worksheets?")) return new Response(JSON.stringify({ value: [{ name: "Orders" }, { name: "Teachers" }] }));
    return new Response("{}", { status: 200 });
  };
  assert.deepEqual(await ensureTabs(call), ["Homerooms"]);
  assert.ok(!calls.some((c) => c.path.startsWith("/worksheets/Teachers")));
  assert.ok(calls.some((c) => c.body && c.body.formulas && String(c.body.formulas[0][0]).startsWith("=LET(")));
});

test("the Gradelink tab is added once, with the billing list read from the Orders table", async () => {
  const { ensureGradelinkTab, BILLING_FORMULA, BILL_WEEK } = await import("../src/gradelink.js");
  const book = (names) => {
    const calls = [];
    const call = async (method, path, body) => {
      calls.push({ method, path, body });
      if (path.startsWith("/worksheets?")) return new Response(JSON.stringify({ value: names.map((name) => ({ name })) }));
      return new Response("{}", { status: 200 });
    };
    return { call, calls };
  };
  let b = book(["Orders", "Gradelink"]);
  assert.equal(await ensureGradelinkTab(b.call), false);
  assert.equal(b.calls.length, 1);
  b = book(["Orders"]);
  assert.equal(await ensureGradelinkTab(b.call), true);
  assert.ok(b.calls.length < 30);
  const a7 = b.calls.find((c) => c.path.endsWith("range(address='A7')"));
  assert.equal(a7.body.formulas[0][0], BILLING_FORMULA);
  // Balanced brackets, every Orders column it reads exists, and the empty week has a plain message.
  const f = BILLING_FORMULA;
  assert.equal((f.match(/"/g) || []).length % 2, 0);
  const code = f.replace(/"[^"]*"/g, '""');
  assert.equal((code.match(/\(/g) || []).length, (code.match(/\)/g) || []).length);
  for (const col of f.match(/Orders\[[^\]]+\]/g)) {
    assert.ok(["Week of", "Student", "Grade", "Parent", "Lunches", "Total", "Notes"].includes(col.slice(7, -1)), col);
  }
  assert.match(f, /m,"No lunch orders for this week yet\.",IF\(SUM\(f\)=0,m,/);
  assert.ok(!/IFERROR\(LET/.test(f));
  // Billing week: Monday-Thursday this week, Friday-Sunday the coming week (as Excel's WEEKDAY(d,3): Monday = 0).
  const monday = (iso) => { const d = new Date(`${iso}T12:00:00Z`); d.setUTCDate(d.getUTCDate() + 3); d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() + 6) % 7)); return d.toISOString().slice(0, 10); };
  assert.equal(BILL_WEEK, "=TODAY()+3-WEEKDAY(TODAY()+3,3)");
  for (const day of ["2026-10-19", "2026-10-21", "2026-10-22"]) assert.equal(monday(day), "2026-10-19", day);
  for (const day of ["2026-10-23", "2026-10-24", "2026-10-25"]) assert.equal(monday(day), "2026-10-26", day);
});

test("POST /admin needs the setup key", async () => {
  const env = { ...ENV0, LUNCH_KV: kv() };
  const req = (key) => new Request("https://lunch.test/admin", { method: "POST", body: '{"action":"setup"}', headers: { "X-Setup-Key": key } });
  assert.equal((await worker.fetch(req("wrong"), env)).status, 403);
  assert.equal((await worker.fetch(req("setup"), env)).status, 503); // right key, not connected
});

// ------------------------------------------------------------------ parent-teacher conferences
import { conflict, normTime, takenSlots, teacherEmail, validateBooking, bookingRow, parentMail, ensureConferenceTab } from "../src/conferences.js";

const SCHED = {
  date: "2026-11-16", date_label: "Monday, November 16, 2026", minutes: 25,
  deadline: "2026-11-13T20:00:00+00:00", deadline_label: "Friday, November 13 at 3:00 PM", note: "",
  slots: [{ id: "0900", label: "9:00 AM", end: "9:25 AM" }, { id: "0930", label: "9:30 AM", end: "9:55 AM" },
    { id: "1230", label: "12:30 PM", end: "12:55 PM" }],
  teachers: [{ id: "mr-damm", name: "Mr. Damm", subject: "9th & 10th Grade", members: ["Mr. Damm"] },
    { id: "miss-ryans", name: "Miss Ryans", subject: "1st Grade", members: ["Miss Ryans"] },
    { id: "high-school-teachers", name: "High School Teachers", subject: "7th–12th Grade", members: ["Mr. Napierala", "Mr. Damm", "Mr. Greene"] }],
};
const CONF_URL = "https://site.test/conferences/schedule.json";
const BOOK = { teacher: "mr-damm", slot: "0930", date: "2026-11-16", parent_first: "Pat", parent_last: "Smith",
  email: "Pat@Example.com", phone: "301-555-0100", student: "Sam Smith", notes: "Reading" };
const CHEAD = ["Booked", "Date", "Time", "Teacher", "Student", "Parent", "Email", "Phone", "Notes", "Booking #"];
const crow = (date, time, teacher, email = "x@y.z") => ["", date, time, teacher, "S", "P", email, "", "", "A"];

test("conference bookings are checked against the schedule", () => {
  const ok = validateBooking(BOOK, SCHED, BEFORE);
  assert.equal(ok.error, undefined);
  assert.equal(ok.booking.email, "pat@example.com");
  assert.equal(ok.booking.slot.label, "9:30 AM");
  assert.match(validateBooking(BOOK, SCHED, Date.parse("2026-11-13T20:00:01Z")).error, /closed/);
  assert.match(validateBooking({ ...BOOK, teacher: "nobody" }, SCHED, BEFORE).error, /teacher/);
  assert.match(validateBooking({ ...BOOK, slot: "0915" }, SCHED, BEFORE).error, /time/);
  assert.match(validateBooking({ ...BOOK, date: "2026-11-17" }, SCHED, BEFORE).error, /schedule changed/);
  assert.match(validateBooking({ ...BOOK, student: " " }, SCHED, BEFORE).error, /student/);
  assert.match(validateBooking({ ...BOOK, email: "nope" }, SCHED, BEFORE).error, /email/);
  assert.match(validateBooking(BOOK, { ...SCHED, teachers: [] }, BEFORE).error, /aren't open/);
});

test("taken times come from the workbook, however the office typed them", () => {
  const rows = [
    { date: "2026-11-16", time: normTime("9:30:00 AM"), teacher: "mr damm", email: "a@b.c" },
    { date: "", time: normTime("12:30"), teacher: "mr damm", email: "" },        // typed 24-hour, no date
    { date: "2026-11-16", time: normTime("9:00 am"), teacher: "miss ryans", email: "" },
    { date: "2026-04-01", time: "9:00 AM", teacher: "mr damm", email: "" },     // another conference day
    { date: "2026-11-16", time: "9:00 AM", teacher: "mrs nobody", email: "" },
  ];
  assert.deepEqual(takenSlots(SCHED, rows), { "mr-damm": ["0930", "1230"], "miss-ryans": ["0900"] });
  const b = validateBooking(BOOK, SCHED, BEFORE).booking;
  assert.equal(conflict(b, SCHED, rows).code, "taken");
  const other = validateBooking({ ...BOOK, slot: "0900", email: "a@b.c" }, SCHED, BEFORE).booking;
  assert.equal(conflict(other, SCHED, rows).code, "already");
  assert.equal(conflict(validateBooking({ ...BOOK, slot: "0900" }, SCHED, BEFORE).booking, SCHED, rows), null);
});

test("conference row and emails: formula-safe, escaped, teacher found by name", () => {
  const b = validateBooking({ ...BOOK, student: "=HYPERLINK()", parent_first: "<b>" }, SCHED, BEFORE).booking;
  const row = bookingRow(b, SCHED, "ABCDE", "11/1/2026 9:00 AM");
  assert.deepEqual(row.slice(1, 5), ["11/16/2026", "9:30 AM", "Mr. Damm", "'=HYPERLINK()"]);
  assert.equal(row.length, CHEAD.length);
  const html = parentMail(b, SCHED, "ABCDE", ENV0).message.body.content;
  assert.ok(html.includes("&lt;b&gt;") && !html.includes("<b><b>"));
  assert.ok(html.includes("9:30 AM to 9:55 AM") && html.includes("Monday, November 16, 2026"));
  assert.equal(teacherEmail([{ grade: "9th Grade", teacher: "Mr Damm", email: "damm@school.test" }], "Mr. Damm"), "damm@school.test");
  assert.equal(teacherEmail([{ grade: "1st Grade", teacher: "Miss Ryans", email: "" }], "Miss Ryans"), "");
});

function fakeConferences({ rows = [], table = true } = {}) {
  const calls = [];
  globalThis.fetch = async (url, init = {}) => {
    url = String(url); calls.push({ url, init });
    const ok = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { "Content-Type": "application/json" } });
    if (url.startsWith(CONF_URL)) return ok(SCHED);
    if (url.includes("/oauth2/v2.0/token")) return ok({ access_token: "at", refresh_token: "rt1", expires_in: 3600 });
    if (url.includes("/workbook/tables/Bookings/range")) return table ? ok({ text: [CHEAD, ...rows] }) : ok({ error: "ItemNotFound" }, 404);
    if (url.includes("/workbook/tables/Bookings/rows/add")) return ok({}, 201);
    if (url.includes("/workbook/worksheets?")) return ok({ value: [{ name: "Orders" }] });
    if (url.includes("/workbook/worksheets/Teachers/usedRange")) {
      return ok({ values: [["Grade", "Teacher", "Email"], ["9th Grade", "Mr. Damm", "damm@school.test"],
        ["7th Grade", "Mr. Napierala", "nap@school.test"], ["11th Grade", "Mr. Greene", ""]] });
    }
    if (url.includes("/tables/add")) return ok({ id: "{T1}" });
    if (url.includes("/workbook/")) return ok({});
    if (url.endsWith("/me/sendMail")) return new Response(null, { status: 202 });
    return ok({ error: "unexpected " + url }, 500);
  };
  return calls;
}
const CENV = () => ({ ...ENV0, CONF_URL, LUNCH_KV: kv({ ms_refresh: "rt1" }) });
const cpost = (body) => new Request("https://lunch.test/conference", {
  method: "POST", body: JSON.stringify(body),
  headers: { "Content-Type": "application/json", Origin: "https://guarantee418.github.io", "CF-Connecting-IP": "5.6.7.8" },
});

test("POST /conference end to end: saves the row, emails the parent and the teacher", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: BEFORE });
  const calls = fakeConferences({ rows: [crow("11/16/2026", "9:00 AM", "Mr. Damm")] });
  const res = await worker.fetch(cpost(BOOK), CENV());
  const data = await res.json();
  assert.equal(res.status, 200, JSON.stringify(data));
  assert.equal(data.time, "9:30 AM"); assert.equal(data.email_sent, true);
  const add = calls.find((c) => c.url.includes("/tables/Bookings/rows/add"));
  assert.equal(JSON.parse(add.init.body).values[0][3], "Mr. Damm");
  const mails = calls.filter((c) => c.url.endsWith("/me/sendMail")).map((c) => JSON.parse(c.init.body).message.toRecipients[0].emailAddress.address);
  assert.deepEqual(mails, ["pat@example.com", "damm@school.test"]);
});

test("POST /conference: a taken time is refused, and a second tap at once is held off", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: BEFORE });
  fakeConferences({ rows: [crow("11/16/2026", "9:30 AM", "Mr. Damm")] });
  const taken = await worker.fetch(cpost(BOOK), CENV());
  assert.equal(taken.status, 409);
  assert.equal((await taken.json()).code, "taken");
  fakeConferences();
  const env = CENV();
  assert.equal((await worker.fetch(cpost({ ...BOOK, slot: "0900" }), env)).status, 200);
  const again = await worker.fetch(cpost({ ...BOOK, slot: "0900", email: "other@example.com" }), env);
  assert.equal(again.status, 409); // the workbook fake doesn't change; the lock catches it
});

test("GET /conferences lists taken times and adds the tab when the workbook has none", async () => {
  fakeConferences({ rows: [crow("11/16/2026", "12:30 PM", "Miss Ryans"), crow("", "", "")] });
  const res = await worker.fetch(new Request("https://lunch.test/conferences", { headers: { Origin: "https://guarantee418.github.io" } }), CENV());
  assert.deepEqual((await res.json()).taken, { "miss-ryans": ["1230"] });
  const calls = fakeConferences({ table: false });
  const empty = await (await worker.fetch(new Request("https://lunch.test/conferences", { headers: { Origin: "https://guarantee418.github.io" } }), CENV())).json();
  assert.deepEqual(empty.taken, {});
  assert.ok(calls.some((c) => c.url.includes("/worksheets/add")));
  assert.ok(calls.some((c) => c.url.includes("/tables/add")));
  assert.ok(calls.some((c) => c.url.includes("/tables/%7BT1%7D") && JSON.parse(c.init.body).name === "Bookings"));
  assert.equal((await worker.fetch(new Request("https://lunch.test/conferences", { headers: { Origin: "https://evil.example" } }), CENV())).status, 403);
});

test("POST /conference for a shared panel emails each of its teachers that has an email", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: BEFORE });
  const calls = fakeConferences();
  const res = await worker.fetch(cpost({ ...BOOK, teacher: "high-school-teachers" }), CENV());
  assert.equal(res.status, 200);
  const mails = calls.filter((c) => c.url.endsWith("/me/sendMail")).map((c) => JSON.parse(c.init.body).message);
  assert.deepEqual(mails.map((m) => m.toRecipients[0].emailAddress.address), ["pat@example.com", "nap@school.test", "damm@school.test"]);
  assert.ok(mails[1].body.content.startsWith("<p>Mr. Napierala,</p>") && mails[1].body.content.includes("Mr. Napierala, Mr. Damm, Mr. Greene"));
});

test("staff lunches use the staff price; a menu without one refuses staff", () => {
  const menu = { ...MENU, prices: { ...MENU.prices, staff: 4 },
    grades: [...MENU.grades, { value: "STAFF", label: "Staff", preschool: false, staff: true }] };
  const kid = { name: "Mr. Teacher", grade: "STAFF", days: ["mon", "wed"] };
  const ok = validateOrder({ ...ORDER, children: [kid] }, menu, BEFORE);
  assert.equal(ok.order.total, 8);
  const rows = orderRows(ok.order, menu, "ABCDE", new Date(BEFORE));
  assert.equal(rows[0][7], "Staff"); assert.equal(rows[0][8], "No"); assert.equal(rows[0][9], 4);
  const noPrice = { ...menu, prices: { regular: 8, preschool: 5 } };
  assert.match(validateOrder({ ...ORDER, children: [kid] }, noPrice, BEFORE).error, /grade/);
});

test("fillTeacherEmails fills only blank emails, by grade code or label", async () => {
  const { fillTeacherEmails } = await import("../src/homerooms.js");
  let written;
  const call = async (method, path, body) => {
    if (method === "GET") return new Response(JSON.stringify({ values: [["Grade", "Teacher", "Email"],
      ["Preschool (K3)", "Miss Dove", ""], ["1st Grade", "Miss Ryans", "kept@school.test"], ["12th Grade", "Mr. Greene", ""]] }));
    written = body; return new Response("{}");
  };
  const n = await fillTeacherEmails(call, "K3=dove@school.test, 1=new@school.test,12th Grade=g@school.test,9=bad");
  assert.equal(n, 2);
  assert.deepEqual(written.values, [["Email"], ["dove@school.test"], ["kept@school.test"], ["g@school.test"]]);
});

test("deleteTestBooking removes only a TEST row, by booking number", async () => {
  const { deleteTestBooking } = await import("../src/conferences.js");
  const deleted = [];
  const call = async (method, path) => {
    if (method === "GET") return new Response(JSON.stringify({ text: [CHEAD,
      ["", "11/16/2026", "9:00 AM", "Mr. Damm", "Real Kid", "P", "a@b.c", "", "", "REAL1"],
      ["", "11/16/2026", "12:40 PM", "TEST Teacher", "TEST Student", "TEST", "g@x", "", "", "MBG7S"]] }));
    deleted.push(path); return new Response(null, { status: 204 });
  };
  assert.deepEqual(await deleteTestBooking(call, "REAL1"), { deleted: false, error: "not a TEST booking" });
  assert.deepEqual(await deleteTestBooking(call, "NOPE1"), { deleted: false, error: "no booking with that number" });
  assert.deepEqual(await deleteTestBooking(call, "MBG7S"), { deleted: true });
  assert.deepEqual(deleted, ["/tables/Bookings/rows/itemAt(index=1)"]);
});

test("deleteTestOrder removes all rows of a TEST order, bottom first, and nothing real", async () => {
  const { deleteTestOrder } = await import("../src/index.js");
  const deleted = [];
  const call = async (method, path) => {
    if (method === "GET") return new Response(JSON.stringify({ text: [HEAD,
      row("", "Real Kid", "3rd Grade", ["Mon"]).map((v, i) => (i === 2 ? "REAL1" : v)),
      ["", "10/26/2026", "25MQU", "TEST Parent", "", "", "TEST Student", "3rd Grade"],
      ["", "", "", "", "", "", "", ""],
      ["", "10/26/2026", "25MQU", "TEST Parent", "", "", "TEST Staff", "Staff"],
      ["", "10/26/2026", "MIXED", "Pat", "", "", "TEST Student", "3rd Grade"],
      ["", "10/26/2026", "MIXED", "Pat", "", "", "Real Kid", "3rd Grade"]] }));
    deleted.push(path); return new Response(null, { status: 204 });
  };
  assert.deepEqual(await deleteTestOrder(call, "Orders", "REAL1"), { deleted: 0, error: "not a TEST order" });
  assert.deepEqual(await deleteTestOrder(call, "Orders", "MIXED"), { deleted: 0, error: "not a TEST order" });
  assert.deepEqual(await deleteTestOrder(call, "Orders", "NOPE1"), { deleted: 0, error: "no order with that number" });
  assert.deepEqual(await deleteTestOrder(call, "Orders", "25MQU"), { deleted: 2 });
  assert.deepEqual(deleted, ["/tables/Orders/rows/itemAt(index=3)", "/tables/Orders/rows/itemAt(index=1)"]);
});

// ------------------------------------------------------------------ school tours
import { takenCounts, tourConflict, validateSignup, signupRow, familyMail, officeMail } from "../src/tours.js";

const TOURS = { note: "", tours: [
  { id: "2027-01-28-1300", date: "2027-01-28", time: "1:00 PM", label: "Thursday, January 28, 2027 at 1:00 PM",
    starts: "2027-01-28T18:00:00+00:00", spots: 2, deadline: "2027-01-27T20:00:00+00:00", note: "" },
  { id: "2027-02-25-1300", date: "2027-02-25", time: "1:00 PM", label: "Thursday, February 25, 2027 at 1:00 PM",
    starts: "2027-02-25T18:00:00+00:00", spots: 5, deadline: "2027-02-24T20:00:00+00:00", note: "" }] };
const TOUR_URL = "https://site.test/tours/schedule.json";
const SIGN = { tour: "2027-01-28-1300", parent_first: "Pat", parent_last: "Smith", email: "Pat@Example.com",
  phone: "301-555-0100", children: "Sam, 3rd grade", notes: "" };
const THEAD = ["Signed up", "Tour date", "Time", "Parent", "Email", "Phone", "Children and grades", "Notes", "Signup #"];
const trow = (date, time, email) => ["", date, time, "P", email, "", "", "", "A"];

test("tour sign-ups are checked against the published tours", () => {
  const ok = validateSignup(SIGN, TOURS, BEFORE);
  assert.equal(ok.error, undefined);
  assert.equal(ok.signup.email, "pat@example.com");
  assert.match(validateSignup({ ...SIGN, tour: "nope" }, TOURS, BEFORE).error, /choose a tour/);
  assert.match(validateSignup(SIGN, TOURS, Date.parse("2027-01-27T20:00:01Z")).error, /closed/);
  assert.match(validateSignup({ ...SIGN, children: "" }, TOURS, BEFORE).error, /children/);
  assert.match(validateSignup({ ...SIGN, phone: "" }, TOURS, BEFORE).error, /phone/);
});

test("a tour is full at its number of families, counting rows the office typed", () => {
  const rows = [{ date: "2027-01-28", time: normTime("1:00:00 PM"), email: "a@b.c" },
    { date: "2027-01-28", time: normTime("13:00"), email: "d@e.f" },
    { date: "2027-02-25", time: "1:00 PM", email: "pat@example.com" }];
  assert.deepEqual(takenCounts(TOURS, rows), { "2027-01-28-1300": 2, "2027-02-25-1300": 1 });
  const s = validateSignup(SIGN, TOURS, BEFORE).signup;
  assert.equal(tourConflict(s, rows).code, "full");
  assert.equal(tourConflict(s, rows.slice(1)), null);
  const feb = validateSignup({ ...SIGN, tour: "2027-02-25-1300" }, TOURS, BEFORE).signup;
  assert.equal(tourConflict(feb, rows).code, "already");
});

test("tour row and emails: formula-safe and escaped; the office hears who signed up", () => {
  const s = validateSignup({ ...SIGN, children: "=SUM(1)", parent_first: "<i>" }, TOURS, BEFORE).signup;
  const row = signupRow(s, "TOUR1", "1/5/2027 9:00 AM");
  assert.deepEqual(row.slice(1, 3), ["1/28/2027", "1:00 PM"]);
  assert.equal(row[6], "'=SUM(1)"); assert.equal(row.length, THEAD.length);
  const fam = familyMail(s, "TOUR1", ENV0).message;
  assert.ok(fam.body.content.includes("&lt;i&gt;") && fam.subject.includes("January 28, 2027"));
  const off = officeMail(s, "TOUR1", 2, { TOUR_NOTIFY: "office@school.test" }).message;
  assert.equal(off.toRecipients[0].emailAddress.address, "office@school.test");
  assert.equal(off.replyTo[0].emailAddress.address, "pat@example.com");
  assert.ok(off.body.content.includes("2 of 2 spots"));
});

function fakeTours({ rows = [], table = true } = {}) {
  const calls = [];
  globalThis.fetch = async (url, init = {}) => {
    url = String(url); calls.push({ url, init });
    const ok = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { "Content-Type": "application/json" } });
    if (url.startsWith(TOUR_URL)) return ok(TOURS);
    if (url.includes("/oauth2/v2.0/token")) return ok({ access_token: "at", refresh_token: "rt1", expires_in: 3600 });
    if (url.includes("/workbook/tables/Signups/range")) return table ? ok({ text: [THEAD, ...rows] }) : ok({ error: "ItemNotFound" }, 404);
    if (url.includes("/workbook/tables/Signups/rows/add")) return ok({}, 201);
    if (url.includes("/workbook/worksheets?")) return ok({ value: [{ name: "Orders" }] });
    if (url.includes("/tables/add")) return ok({ id: "{T2}" });
    if (url.includes("/workbook/")) return ok({});
    if (url.endsWith("/me/sendMail")) return new Response(null, { status: 202 });
    return ok({ error: "unexpected " + url }, 500);
  };
  return calls;
}
const TENV = () => ({ ...ENV0, TOUR_URL, TOUR_NOTIFY: "office@school.test", LUNCH_KV: kv({ ms_refresh: "rt1" }) });
const tpost = (body) => new Request("https://lunch.test/tour", {
  method: "POST", body: JSON.stringify(body),
  headers: { "Content-Type": "application/json", Origin: "https://guarantee418.github.io", "CF-Connecting-IP": "9.9.9.9" },
});

test("POST /tour end to end: saves the row, emails the family and the office, frees the lock", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: BEFORE });
  const calls = fakeTours({ rows: [trow("1/28/2027", "1:00 PM", "a@b.c")] });
  const env = TENV();
  const res = await worker.fetch(tpost(SIGN), env);
  assert.equal(res.status, 200, await res.clone().text());
  const mails = calls.filter((c) => c.url.endsWith("/me/sendMail")).map((c) => JSON.parse(c.init.body).message);
  assert.deepEqual(mails.map((m) => m.toRecipients[0].emailAddress.address), ["pat@example.com", "office@school.test"]);
  assert.ok(mails[1].body.content.includes("2 of 2 spots"));
  assert.equal(await env.LUNCH_KV.get("tour-lock:2027-01-28-1300"), null);
});

test("POST /tour refuses a full tour; GET /tours counts and adds the tab when missing", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: BEFORE });
  fakeTours({ rows: [trow("1/28/2027", "1:00 PM", "a@b.c"), trow("1/28/2027", "1:00 PM", "d@e.f")] });
  const full = await worker.fetch(tpost(SIGN), TENV());
  assert.equal(full.status, 409);
  assert.equal((await full.json()).code, "full");
  const get = (env) => worker.fetch(new Request("https://lunch.test/tours", { headers: { Origin: "https://guarantee418.github.io" } }), env);
  assert.deepEqual((await (await get(TENV())).json()).taken, { "2027-01-28-1300": 2 });
  const calls = fakeTours({ table: false });
  assert.deepEqual((await (await get(TENV())).json()).taken, {});
  assert.ok(calls.some((c) => c.url.includes("/worksheets/add")) && calls.some((c) => c.url.includes("/tables/%7BT2%7D")));
});

test("POST /tour: a second family at the same moment is asked to try again", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: BEFORE });
  fakeTours();
  const env = TENV();
  await env.LUNCH_KV.put("tour-lock:2027-01-28-1300", "1");
  const busy = await worker.fetch(tpost(SIGN), env);
  assert.equal(busy.status, 409);
  assert.equal((await busy.json()).code, "busy");
});

// ------------------------------------------------------------------ transcript requests
import { COLUMNS as TR_COLUMNS, validateRequest, requestRow, requesterMail, officeMail as trOfficeMail, startStatus } from "../src/transcripts.js";

const TR = {
  who: "alumni", student_first: "Pat", student_last: "Smith", former_name: "Pat Jones", birth_date: "1995-04-02",
  class_year: "2013", email: "Pat@Example.com", phone: "301-555-0100", kind: "official", copies: "2",
  reason: "College/University", delivery: "mail", send_to: "State University", attention: "Admissions",
  street: "1 College Ave", city: "College Park", state: "md", zip: "20742", consent: true, signature: "Pat Smith",
};
const TRENV = { ...ENV0, TRANSCRIPT_NOTIFY: "office@school.test", TRANSCRIPT_FEE_ALUMNI: "10", TRANSCRIPT_FEE_STUDENT: "10", TRANSCRIPT_PAY_URL: "" };

test("transcript requests are checked on the server", () => {
  const ok = validateRequest(TR, BEFORE);
  assert.equal(ok.error, undefined);
  assert.equal(ok.request.state, "MD"); assert.equal(ok.request.copies, 2); assert.equal(ok.request.relation, "student");
  assert.match(validateRequest({ ...TR, who: "" }, BEFORE).error, /current student or alumni/);
  assert.match(validateRequest({ ...TR, birth_date: "2030-01-01" }, BEFORE).error, /date of birth/);
  assert.match(validateRequest({ ...TR, class_year: "1900" }, BEFORE).error, /graduated/);
  assert.match(validateRequest({ ...TR, zip: "" }, BEFORE).error, /mailing address/);
  assert.match(validateRequest({ ...TR, delivery: "email" }, BEFORE).error, /email address the transcript/);
  assert.equal(validateRequest({ ...TR, delivery: "pickup", send_to: "" }, BEFORE).error, undefined);
  assert.match(validateRequest({ ...TR, consent: "yes" }, BEFORE).error, /release/);
  assert.match(validateRequest({ ...TR, copies: "9" }, BEFORE).error, /copies/);
  assert.match(validateRequest({ ...TR, reason: "Other" }, BEFORE).error, /reason/);
  // A parent can ask for an enrolled student's transcript, not a former student's.
  const kid = { ...TR, who: "student", grade: "11", relation: "parent", requester_name: "Lee Smith" };
  assert.equal(validateRequest(kid, BEFORE).request.relation, "parent");
  assert.equal(validateRequest({ ...TR, relation: "parent" }, BEFORE).request.relation, "student");
  assert.match(validateRequest({ ...kid, requester_name: "" }, BEFORE).error, /your name/);
  assert.match(validateRequest({ ...kid, grade: "13" }, BEFORE).error, /grade/);
});

test("transcript row, status and emails: formula-safe, escaped, payment told the right way", () => {
  const q = validateRequest({ ...TR, notes: "=HYPERLINK(1)", student_first: "<b>" }, BEFORE).request;
  const row = requestRow(q, "TRAN1", "1/5/2027 9:00 AM", TRENV);
  assert.equal(row.length, TR_COLUMNS.length);
  assert.deepEqual(row.slice(1, 5), ["TRAN1", "Waiting for payment", "$10.00", "Alumni"]);
  assert.equal(row[7], "4/2/1995"); assert.equal(row[18], "1 College Ave, College Park, MD 20742");
  assert.equal(row[TR_COLUMNS.indexOf("Notes")], "'=HYPERLINK(1)");
  const mine = requesterMail(q, "TRAN1", TRENV).message;
  assert.ok(mine.body.content.includes("&lt;b&gt;") && /office will contact you/.test(mine.body.content));
  const paid = requesterMail(q, "TRAN1", { ...TRENV, TRANSCRIPT_PAY_URL: "https://pay.test/form" }).message.body.content;
  assert.ok(paid.includes('href="https://pay.test/form"') && /pay the \$10\.00/.test(paid));
  const kid = validateRequest({ ...TR, who: "student", grade: "11", relation: "parent", requester_name: "Lee Smith" }, BEFORE).request;
  assert.equal(startStatus(kid, TRENV), "Fee to post in Gradelink");
  assert.equal(startStatus(kid, { ...TRENV, TRANSCRIPT_FEE_STUDENT: "0" }), "New");
  assert.ok(/Gradelink account/.test(requesterMail(kid, "TRAN2", TRENV).message.body.content));
  const off = trOfficeMail(kid, "TRAN2", TRENV).message;
  assert.equal(off.toRecipients[0].emailAddress.address, "office@school.test");
  assert.equal(off.replyTo[0].emailAddress.address, "pat@example.com");
  assert.ok(/Lee Smith \(parent\/guardian\)/.test(off.body.content) && /Post the \$10\.00/.test(off.body.content));
});

function fakeTranscripts({ table = true } = {}) {
  const calls = [];
  let made = false;
  globalThis.fetch = async (url, init = {}) => {
    url = String(url); calls.push({ url, init });
    const ok = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { "Content-Type": "application/json" } });
    if (url.includes("/oauth2/v2.0/token")) return ok({ access_token: "at", refresh_token: "rt1", expires_in: 3600 });
    if (url.includes("/workbook/tables/Requests/rows/add")) return table || made ? ok({}, 201) : ok({ error: "ItemNotFound" }, 404);
    if (url.includes("/workbook/worksheets?")) return ok({ value: [{ name: "Orders" }] });
    if (url.includes("/tables/add")) { made = true; return ok({ id: "{T3}" }); }
    if (url.includes("/workbook/")) return ok({});
    if (url.endsWith("/me/sendMail")) return new Response(null, { status: 202 });
    return ok({ error: "unexpected " + url }, 500);
  };
  return calls;
}
const trpost = (body) => new Request("https://lunch.test/transcript", {
  method: "POST", body: JSON.stringify(body),
  headers: { "Content-Type": "application/json", Origin: "https://guarantee418.github.io", "CF-Connecting-IP": "9.9.9.9" },
});

test("POST /transcript end to end: saves the row (adding the tab), emails the requester and the office", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: BEFORE });
  const calls = fakeTranscripts({ table: false });
  const res = await worker.fetch(trpost(TR), { ...TRENV, LUNCH_KV: kv({ ms_refresh: "rt1" }) });
  const out = await res.json();
  assert.equal(res.status, 200, JSON.stringify(out));
  assert.match(out.request_id, /^[A-Z0-9]{5}$/); assert.equal(out.fee_alumni, 10); assert.equal(out.pay_url, "");
  assert.ok(calls.some((c) => c.url.includes("/worksheets/add")));
  const mails = calls.filter((c) => c.url.endsWith("/me/sendMail")).map((c) => JSON.parse(c.init.body).message);
  assert.deepEqual(mails.map((m) => m.toRecipients[0].emailAddress.address), ["pat@example.com", "office@school.test"]);
  const bad = await worker.fetch(trpost({ ...TR, signature: "" }), { ...TRENV, LUNCH_KV: kv({ ms_refresh: "rt1" }) });
  assert.equal(bad.status, 400);
  const get = await worker.fetch(new Request("https://lunch.test/transcripts", { headers: { Origin: "https://guarantee418.github.io" } }), TRENV);
  assert.deepEqual(await get.json(), { ok: true, fee_alumni: 10, fee_student: 10, pay_url: "" });
  const stranger = await worker.fetch(new Request("https://lunch.test/transcripts", { headers: { Origin: "https://evil.test" } }), TRENV);
  assert.equal(stranger.status, 403);
});

// ------------------------------------------------------------------ Request Info inquiries
import { COLUMNS as IN_COLUMNS, validateInquiry, inquiryRow, familyMail as inFamilyMail, officeMail as inOfficeMail } from "../src/inquiries.js";

const INQ = {
  parent_first: "Pat", parent_last: "Smith", parent2_name: "", email: "Pat@Example.com", phone: "301-555-0100",
  student_first: "Sam", student_last: "Smith", student_age: "8", student_grade: "3rd Grade", student_birthday: "2018-03-04",
  other_students: "", current_school: "Clinton Elementary", heard_about: "A friend from church",
};
const INENV = { ...ENV0, INQUIRY_NOTIFY: "office@school.test, principal@school.test" };

test("Request Info inquiries are checked on the server", () => {
  const ok = validateInquiry(INQ, BEFORE);
  assert.equal(ok.error, undefined);
  assert.equal(ok.inquiry.email, "pat@example.com"); assert.equal(ok.inquiry.age, 8);
  assert.match(validateInquiry({ ...INQ, parent_last: "" }, BEFORE).error, /your first and last name/);
  assert.match(validateInquiry({ ...INQ, email: "nope" }, BEFORE).error, /email/);
  assert.match(validateInquiry({ ...INQ, student_age: "40" }, BEFORE).error, /age/);
  assert.match(validateInquiry({ ...INQ, student_grade: "13th" }, BEFORE).error, /grade/);
  assert.match(validateInquiry({ ...INQ, student_birthday: "2030-01-01" }, BEFORE).error, /birthday/);
  assert.match(validateInquiry({ ...INQ, current_school: " " }, BEFORE).error, /school/);
  assert.match(validateInquiry({ ...INQ, heard_about: "" }, BEFORE).error, /heard/);
});

test("inquiry row and emails: formula-safe, escaped, office can reply to the family", () => {
  const q = validateInquiry({ ...INQ, heard_about: "=cmd()", student_first: "<i>" }, BEFORE).inquiry;
  const row = inquiryRow(q, "INQ01", "1/5/2027 9:00 AM");
  assert.equal(row.length, IN_COLUMNS.length);
  assert.deepEqual(row.slice(1, 3), ["INQ01", "New"]);
  assert.equal(row[IN_COLUMNS.indexOf("Birthday")], "3/4/2018");
  assert.equal(row[IN_COLUMNS.indexOf("Heard about IBA")], "'=cmd()");
  const fam = inFamilyMail(q, "INQ01", INENV).message;
  assert.ok(fam.body.content.includes("&lt;i&gt;") && fam.toRecipients[0].emailAddress.address === "pat@example.com");
  const off = inOfficeMail(q, "INQ01", INENV).message;
  assert.deepEqual(off.toRecipients.map((r) => r.emailAddress.address), ["office@school.test", "principal@school.test"]);
  assert.equal(off.replyTo[0].emailAddress.address, "pat@example.com");
});

test("POST /inquiry end to end: saves the row (adding the tab), emails the family and the office", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: BEFORE });
  const calls = [];
  let made = false;
  globalThis.fetch = async (url, init = {}) => {
    url = String(url); calls.push({ url, init });
    const ok = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { "Content-Type": "application/json" } });
    if (url.includes("/oauth2/v2.0/token")) return ok({ access_token: "at", refresh_token: "rt1", expires_in: 3600 });
    if (url.includes("/workbook/tables/InfoRequests/rows/add")) return made ? ok({}, 201) : ok({ error: "ItemNotFound" }, 404);
    if (url.includes("/workbook/worksheets?")) return ok({ value: [{ name: "Orders" }] });
    if (url.includes("/tables/add")) { made = true; return ok({ id: "{T4}" }); }
    if (url.includes("/workbook/")) return ok({});
    if (url.endsWith("/me/sendMail")) return new Response(null, { status: 202 });
    return ok({ error: "unexpected " + url }, 500);
  };
  const post = (body) => worker.fetch(new Request("https://lunch.test/inquiry", {
    method: "POST", body: JSON.stringify(body),
    headers: { "Content-Type": "application/json", Origin: "https://guarantee418.github.io", "CF-Connecting-IP": "8.8.8.8" },
  }), { ...INENV, LUNCH_KV: kv({ ms_refresh: "rt1" }) });
  const res = await post(INQ);
  const out = await res.json();
  assert.equal(res.status, 200, JSON.stringify(out));
  assert.match(out.request_id, /^[A-Z0-9]{5}$/);
  const mails = calls.filter((c) => c.url.endsWith("/me/sendMail")).map((c) => JSON.parse(c.init.body).message);
  assert.deepEqual(mails.map((m) => m.toRecipients[0].emailAddress.address), ["pat@example.com", "office@school.test"]);
  assert.equal((await post({ ...INQ, phone: "" })).status, 400);
  const trap = await (await post({ ...INQ, website: "spam" })).json();
  assert.equal(trap.request_id, "0");
});

// ------------------------------------------------------------------ workbook layout (admin)
import { organize, inspect, STEPS, GUIDE, NOTES } from "../src/guide.js";

function fakeBook({ orders = [], b3 = "=MAX(Orders[Week of])" } = {}) {
  const calls = [];
  const sheets = ["Orders", "Homerooms", "Gradelink", "Kitchen", "How to use", "Teachers", "Staff Families", "Conferences", "Tours", "Transcripts", "Inquiries", "Feedback", "Absences"];
  const ok = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { "Content-Type": "application/json" } });
  const call = async (method, path, body) => {
    calls.push({ method, path, body });
    if (path.startsWith("/worksheets?")) return ok({ value: sheets.map((name, position) => ({ name, position, visibility: "Visible" })) });
    if (path === "/tables/Orders/range?$select=values") return ok({ values: [["Order #", "Student", "Lunches", "Total"], ...orders] });
    if (path.startsWith("/tables/") && path.endsWith("range?$select=values")) return ok({ values: [["A", "B"], ["", ""]] });
    if (path.includes("range(address='B3')?$select=formulas")) return ok({ formulas: [[b3]] });
    return ok({});
  };
  return { call, calls };
}

test("organize: every step stays under Cloudflare's 50 outside calls and never touches a table's rows with data", async () => {
  for (const step of STEPS) {
    const { call, calls } = fakeBook({ orders: [["ABCDE", "Sam", 2, 16], ["", "", 0, 0], ["", "", 0, 0]] });
    const done = await organize(call, step);
    assert.ok(calls.length < 48, `${step}: ${calls.length} calls`);
    assert.ok(!calls.some((c) => c.method === "PATCH" && /\/tables\/[^/]+\/range|headerRowRange/.test(c.path)), step);
    if (step === "rows") {
      // Orders: the two empty rows go (bottom first), the real one stays; other tables keep their only row.
      const del = calls.filter((c) => c.method === "DELETE").map((c) => c.path);
      assert.deepEqual(del, ["/tables/Orders/rows/itemAt(index=2)", "/tables/Orders/rows/itemAt(index=1)"]);
      assert.ok(done.includes("Orders: 2 removed"));
    }
    if (step === "order") {
      const moves = calls.filter((c) => c.method === "PATCH").map((c) => [decodeURIComponent(c.path.slice(12)), c.body.position]);
      assert.deepEqual(moves[0], ["How to use", 0]); assert.deepEqual(moves[1], ["Orders", 1]);
    }
  }
});

test("organize: the Kitchen week box is fixed only while it still holds the original formula", async () => {
  let b = fakeBook();
  await organize(b.call, "kitchen");
  assert.match(b.calls.find((c) => c.method === "PATCH").body.formulas[0][0], /^=IF\(COUNT\(Orders\[Week of\]\)=0,TODAY\(\)/);
  b = fakeBook({ b3: "10/26/2026" });
  assert.deepEqual(await organize(b.call, "kitchen"), ["week box left as it is", "choices table"]);
  assert.equal(b.calls.filter((c) => c.method === "PATCH" && c.path.includes("B3")).length, 0);
  // The Choices table goes below the day counts, with a formula per day (Monday at row 16).
  const t = b.calls.find((c) => c.method === "PATCH" && c.path.includes("range(address='A14:E40')"));
  assert.equal(t.body.formulas.length, 27);
  assert.match(t.body.formulas[2][1], /Orders\[Mon\]/); assert.match(t.body.formulas[7][1], /Orders\[Tue\]/);
});

test("organize: the warnings about the Orders headings are bold red", async () => {
  let b = fakeBook();
  await organize(b.call, "notes");
  const n = NOTES.Orders[1].findIndex((l) => l.startsWith("⚠")) + 1;
  assert.ok(n > 0 && /Price per lunch, Lunches, Total/.test(NOTES.Orders[1][n - 1]));
  assert.ok(b.calls.some((c) => c.path === `/worksheets/Orders/range(address='T${n}')/format/font` && c.body.bold && c.body.color === "#B00020"));
  b = fakeBook();
  await organize(b.call, "guide");
  const warns = GUIDE.map(([k], i) => (k === "warn" ? i + 1 : 0)).filter(Boolean);
  assert.equal(warns.length, 2);
  assert.match(GUIDE[warns[0] - 1][2], /Price per lunch, Mon, Tue, Wed, Thu, Fri, Lunches, Total, Notes/);
  for (const r of warns) {
    assert.ok(b.calls.some((c) => c.path.endsWith(`range(address='A${r}:B${r}')/format/font`) && c.body.bold && c.body.color === "#B00020"), `row ${r}`);
  }
  const values = b.calls.find((c) => c.method === "PATCH" && c.body && c.body.values).body.values;
  assert.equal(values[warns[0] - 1][0], "⚠");
});

test("guide and reminders cover every tab", () => {
  const tabs = GUIDE.filter(([k]) => k === "tab").map(([, a]) => a);
  assert.deepEqual(tabs, ["Orders", "Kitchen", "Homerooms", "Gradelink", "Teachers", "Staff Families", "Conferences", "Tours", "Transcripts", "Inquiries", "Feedback", "Absences"]);
  assert.deepEqual(Object.keys(NOTES).sort(), [...tabs].sort());
  // Reminders start to the right of each table (Orders A:R, Conferences A:J, Tours A:I, Transcripts A:V, Inquiries A:N, Feedback A:J, Absences A:M).
  assert.deepEqual([NOTES.Orders[0], NOTES.Conferences[0], NOTES.Tours[0], NOTES.Transcripts[0], NOTES.Inquiries[0], NOTES.Feedback[0],
    NOTES.Absences[0]], ["T1", "L1", "K1", "X1", "P1", "L1", "O1"]);
});

// ------------------------------------------------------------------ Feedback for Administration
import { COLUMNS as FB_COLUMNS, validateFeedback, feedbackRow } from "../src/feedback.js";

const FB = { name: "", email: "", phone: "", grades: "3rd", topic: "Suggestion", message: "More field trips, please.\nThank you!" };

test("feedback is checked on the server; name and contact are optional", () => {
  const ok = validateFeedback(FB);
  assert.equal(ok.error, undefined);
  assert.equal(ok.feedback.message, "More field trips, please.\nThank you!");
  assert.match(validateFeedback({ ...FB, email: "nope" }).error, /email/);
  assert.match(validateFeedback({ ...FB, topic: "Gossip" }).error, /topic/);
  assert.match(validateFeedback({ ...FB, message: "  " }).error, /feedback/);
  const row = feedbackRow(validateFeedback({ ...FB, message: "=HYPERLINK(1)" }).feedback, "FB001", "1/5/2027 9:00 AM");
  assert.equal(row.length, FB_COLUMNS.length);
  assert.equal(row[3], "(no name)"); assert.equal(row[2], "New"); assert.equal(row[8], "'=HYPERLINK(1)");
});

test("POST /feedback end to end: saves the row (adding the tab), sends no email", async () => {
  const calls = [];
  let made = false;
  globalThis.fetch = async (url, init = {}) => {
    url = String(url); calls.push({ url, init });
    const ok = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { "Content-Type": "application/json" } });
    if (url.includes("/oauth2/v2.0/token")) return ok({ access_token: "at", refresh_token: "rt1", expires_in: 3600 });
    if (url.includes("/workbook/tables/Responses/rows/add")) return made ? ok({}, 201) : ok({ error: "ItemNotFound" }, 404);
    if (url.includes("/workbook/worksheets?")) return ok({ value: [{ name: "Orders" }] });
    if (url.includes("/tables/add")) { made = true; return ok({ id: "{T5}" }); }
    if (url.includes("/workbook/")) return ok({});
    return ok({ error: "unexpected " + url }, 500);
  };
  const post = (body) => worker.fetch(new Request("https://lunch.test/feedback", {
    method: "POST", body: JSON.stringify(body),
    headers: { "Content-Type": "application/json", Origin: "https://guarantee418.github.io", "CF-Connecting-IP": "8.8.4.4" },
  }), { ...ENV0, LUNCH_KV: kv({ ms_refresh: "rt1" }) });
  const res = await post(FB);
  const out = await res.json();
  assert.equal(res.status, 200, JSON.stringify(out));
  assert.match(out.feedback_id, /^[A-Z0-9]{5}$/);
  assert.ok(made);
  assert.ok(!calls.some((c) => c.url.endsWith("/me/sendMail")));
  const added = calls.filter((c) => c.url.includes("/tables/Responses/rows/add")).pop();
  assert.equal(JSON.parse(added.init.body).values[0][7], "Suggestion");
  assert.equal((await post({ ...FB, message: "" })).status, 400);
  assert.equal((await (await post({ ...FB, website: "spam" })).json()).feedback_id, "0");
});

// ------------------------------------------------------------------ Report Student Absence
import { COLUMNS as AB_COLUMNS, validateAbsence, absenceRow, homeroomFor, officeMail as abOfficeMail } from "../src/absences.js";

const ABS = { student_first: "Sam", student_last: "Smith", grade: "3rd Grade", first_day: "2026-10-16", last_day: "",
  reason: "Illness", details: "", parent_name: "Pat Smith", email: "Pat@Example.com", phone: "301-555-0100" };
const AB_TEACHERS = [{ grade: "3rd Grade", teacher: "Mrs. Lundberg", email: "lundberg@school.test" }, { grade: "4th Grade", teacher: "Mrs. Kirby", email: "" }];

test("absence reports are checked on the server", () => {
  const ok = validateAbsence(ABS, BEFORE);
  assert.equal(ok.error, undefined); assert.equal(ok.absence.email, "pat@example.com"); assert.equal(ok.absence.to, "");
  assert.match(validateAbsence({ ...ABS, student_last: "" }, BEFORE).error, /first and last name/);
  assert.match(validateAbsence({ ...ABS, grade: "13th Grade" }, BEFORE).error, /grade/);
  assert.match(validateAbsence({ ...ABS, first_day: "2026-09-01" }, BEFORE).error, /day of the absence/);
  assert.match(validateAbsence({ ...ABS, last_day: "2026-10-14" }, BEFORE).error, /last day/);
  assert.equal(validateAbsence({ ...ABS, last_day: "2026-10-20" }, BEFORE).absence.to, "2026-10-20");
  assert.equal(validateAbsence({ ...ABS, last_day: "2026-10-16" }, BEFORE).absence.to, "");
  assert.match(validateAbsence({ ...ABS, reason: "Other" }, BEFORE).error, /reason/);
  assert.match(validateAbsence({ ...ABS, email: "nope" }, BEFORE).error, /email/);
});

test("absence row and office email: homeroom teacher by grade, formula-safe", () => {
  const q = validateAbsence({ ...ABS, details: "=cmd()", last_day: "2026-10-19" }, BEFORE).absence;
  const t = homeroomFor(AB_TEACHERS, q.grade);
  assert.deepEqual(t, { teacher: "Mrs. Lundberg", email: "lundberg@school.test" });
  assert.equal(homeroomFor(AB_TEACHERS, "4th Grade"), null);
  const row = absenceRow(q, "AB001", "10/15/2026 8:00 AM", t);
  assert.equal(row.length, AB_COLUMNS.length);
  assert.deepEqual(row.slice(2, 8), ["10/16/2026", "10/19/2026", "Sam Smith", "3rd Grade", "Illness", "'=cmd()"]);
  assert.equal(row[11], "Mrs. Lundberg");
  const m = abOfficeMail(q, "AB001", { ABSENCE_NOTIFY: "office@school.test" }, t).message;
  assert.equal(m.toRecipients[0].emailAddress.address, "office@school.test");
  assert.equal(m.ccRecipients[0].emailAddress.address, "lundberg@school.test");
  assert.equal(m.replyTo[0].emailAddress.address, "pat@example.com");
  assert.match(m.subject, /Friday, October 16, 2026 through Monday, October 19, 2026/);
  assert.equal(abOfficeMail(q, "AB001", { ABSENCE_NOTIFY: "o@s.test" }, null).message.ccRecipients, undefined);
  const cc = abOfficeMail(q, "AB001", { ABSENCE_NOTIFY: "o@s.test", ABSENCE_CC: "ibs@school.test" }, t).message.ccRecipients;
  assert.deepEqual(cc.map((r) => r.emailAddress.address), ["lundberg@school.test", "ibs@school.test"]);
  assert.deepEqual(abOfficeMail(q, "AB001", { ABSENCE_NOTIFY: "o@s.test", ABSENCE_CC: "ibs@school.test" }, null)
    .message.ccRecipients.map((r) => r.emailAddress.address), ["ibs@school.test"]);
});

test("POST /absence end to end: saves the row (adding the tab), emails the parent and the office Cc the teacher", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: BEFORE });
  const calls = [];
  let made = false;
  globalThis.fetch = async (url, init = {}) => {
    url = String(url); calls.push({ url, init });
    const ok = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { "Content-Type": "application/json" } });
    if (url.includes("/oauth2/v2.0/token")) return ok({ access_token: "at", refresh_token: "rt1", expires_in: 3600 });
    if (url.includes("/worksheets/Teachers/usedRange")) return ok({ values: [["Grade", "Teacher", "Email"], ["3rd Grade", "Mrs. Lundberg", "lundberg@school.test"]] });
    if (url.includes("/workbook/tables/Reports/rows/add")) return made ? ok({}, 201) : ok({ error: "ItemNotFound" }, 404);
    if (url.includes("/workbook/worksheets?")) return ok({ value: [{ name: "Orders" }] });
    if (url.includes("/tables/add")) { made = true; return ok({ id: "{T6}" }); }
    if (url.includes("/workbook/")) return ok({});
    if (url.endsWith("/me/sendMail")) return new Response(null, { status: 202 });
    return ok({ error: "unexpected " + url }, 500);
  };
  const post = (body) => worker.fetch(new Request("https://lunch.test/absence", {
    method: "POST", body: JSON.stringify(body),
    headers: { "Content-Type": "application/json", Origin: "https://guarantee418.github.io", "CF-Connecting-IP": "9.9.9.9" },
  }), { ...ENV0, ABSENCE_NOTIFY: "office@school.test", LUNCH_KV: kv({ ms_refresh: "rt1" }) });
  const res = await post(ABS);
  const out = await res.json();
  assert.equal(res.status, 200, JSON.stringify(out));
  assert.match(out.absence_id, /^[A-Z0-9]{5}$/);
  const mails = calls.filter((c) => c.url.endsWith("/me/sendMail")).map((c) => JSON.parse(c.init.body).message);
  assert.deepEqual(mails.map((m) => m.toRecipients[0].emailAddress.address), ["pat@example.com", "office@school.test"]);
  assert.equal(mails[1].ccRecipients[0].emailAddress.address, "lundberg@school.test");
  assert.equal((await post({ ...ABS, phone: "" })).status, 400);
  assert.equal((await (await post({ ...ABS, website: "x" })).json()).absence_id, "0");
});

// ------------------------------------------------------------------ lunch choices (e.g. cheese or pepperoni pizza)
import { choiceFormula } from "../src/guide.js";

const CMENU = { ...MENU, days: [MENU.days[0], { ...MENU.days[1], meal: "Pizza", options: ["Cheese pizza", "Pepperoni pizza"] }] };
const CORDER = { ...ORDER, children: [{ name: "Sam Smith", grade: "3", days: ["mon", "wed"], choices: { wed: "pepperoni PIZZA" } }] };

test("a day with choices needs one of them; the sheet gets the choice, the email shows it", () => {
  const ok = validateOrder(CORDER, CMENU, BEFORE);
  assert.equal(ok.error, undefined);
  assert.deepEqual(ok.order.children[0].choices, { wed: "Pepperoni pizza" });
  assert.match(validateOrder({ ...CORDER, children: [{ ...CORDER.children[0], choices: {} }] }, CMENU, BEFORE).error, /Wednesday's lunch for Sam Smith/);
  assert.match(validateOrder({ ...CORDER, children: [{ ...CORDER.children[0], choices: { wed: "Sushi" } }] }, CMENU, BEFORE).error, /Wednesday's lunch/);
  // Monday has no choices, so nothing is asked for it.
  assert.equal(validateOrder({ ...CORDER, children: [{ ...CORDER.children[0], days: ["mon"], choices: {} }] }, CMENU, BEFORE).error, undefined);
  const row = orderRows(ok.order, CMENU, "ABCDE", new Date("2026-10-15T18:05:00Z"))[0];
  assert.deepEqual(row.slice(10, 16), ["X", "", "Pepperoni pizza", "", "", '=COUNTIF(Orders[@[Mon]:[Fri]],"?*")']);
  const mail = confirmationMail(ok.order, CMENU, "ABCDE", ENV0).message.body.content;
  assert.match(mail, /Wednesday Oct 21: Pepperoni pizza/);
});

test("Kitchen choices formula counts each choice per day for the Kitchen's week", () => {
  const f = choiceFormula("Wed");
  assert.match(f, /^=IFERROR\(LET\(u,SORT\(UNIQUE\(FILTER\(Orders\[Wed\],/);
  assert.match(f, /Orders\[Wed\]<>"X"/);
  assert.match(f, /COUNTIFS\(Orders\[Week of\],\$B\$3,Orders\[Wed\],u,Orders\[Preschool\],"Yes"\)/);
});

test("homeroom lists show the lunch choice instead of a check mark", async () => {
  const r = row(46321, "Sam", "3rd Grade", ["Mon"]); r[12] = "Pepperoni <pizza>";
  const kids = await readWeek(fakeCall([HEAD, r]), "2026-10-26");
  assert.deepEqual([...kids[0].days].sort(), ["Mon", "Wed"]);
  assert.deepEqual(kids[0].choices, { Wed: "Pepperoni <pizza>" });
  const html = homeroomMails(kids, TEACHERS, "2026-10-26", ENV0)[0].mail.message.body.content;
  assert.ok(html.includes("Pepperoni &lt;pizza&gt;") && html.includes("&#10003;"));
});

// ------------------------------------------------------------------ several weeks on the menu
import { pickWeek } from "../src/index.js";

test("an order is checked against the week the parent picked", () => {
  const next = { ...MENU, week_start: "2026-10-26", week_label: "October 26 – 30", deadline: "2026-10-23T13:00:00+00:00",
    days: [{ key: "mon", day: "Monday", meal: "Tacos", options: [], date: "2026-10-26", date_label: "Oct 26" }] };
  const menu = { ...MENU, weeks: [{ ...MENU, prices: undefined, grades: undefined }, next] };
  delete menu.weeks[0].prices; delete menu.weeks[0].grades;
  const picked = pickWeek(menu, "2026-10-26");
  assert.equal(picked.week_label, "October 26 – 30"); assert.ok(picked.prices && picked.grades);
  const order = { ...ORDER, week_start: "2026-10-26", children: [{ name: "Sam", grade: "3", days: ["mon"] }] };
  assert.equal(validateOrder(order, picked, BEFORE).error, undefined);
  // The first week's order still works, an unknown week falls back to the top level and is refused as changed.
  assert.equal(validateOrder(ORDER, pickWeek(menu, "2026-10-19"), BEFORE).error, undefined);
  assert.match(validateOrder({ ...ORDER, week_start: "2026-11-02" }, pickWeek(menu, "2026-11-02"), BEFORE).error, /menu changed/);
  // Past its deadline a week is closed even if a later one is open.
  assert.match(validateOrder(order, picked, Date.parse("2026-10-24T12:00:00Z")).error, /closed/);
});

test("staff lunches: the confirmation email has their prices and total", () => {
  const menu = { ...MENU, grades: [...MENU.grades, { value: "STAFF", label: "Staff", preschool: false, staff: true }], prices: { ...MENU.prices, staff: 4 } };
  const ok = validateOrder({ ...ORDER, children: [{ name: "Mr. T", grade: "STAFF", days: ["mon", "wed"] }] }, menu, BEFORE);
  assert.equal(ok.error, undefined); assert.equal(ok.order.total, 8);
  const html = confirmationMail(ok.order, menu, "ABCDE", ENV0).message.body.content;
  assert.match(html, /Mr\. T<\/b> \(Staff\), 2 lunches at \$4\.00: \$8\.00<\/p>/);
  assert.match(html, /Total: \$8\.00<\/b> \(order number ABCDE\)/);
  assert.match(confirmationMail(validateOrder(ORDER, menu, BEFORE).order, menu, "ABCDE", ENV0).message.body.content, /Total: \$/);
});

// ------------------------------------------------------------------ staff families (Staff Families tab)
import { isStaffFamily } from "../src/staff.js";

test("staff families: matched by parent or child last name or email; every child at the staff price, total in the email", () => {
  const fams = [{ name: "kirby", email: "" }, { name: "buck", email: "mom@buck.test" }];
  assert.ok(isStaffFamily(fams, { last: "Kirby", email: "x@y.z" }, []));
  assert.ok(isStaffFamily(fams, { last: "Smith", email: "x@y.z" }, ["Levi Buck"]));
  assert.ok(isStaffFamily(fams, { last: "Smith", email: "MOM@buck.test" }, ["Sam Smith"]));
  assert.ok(!isStaffFamily(fams, { last: "Smith", email: "x@y.z" }, ["Sam Smith"]));
  assert.ok(!isStaffFamily([], { last: "Kirby", email: "" }, []));
  const menu = { ...MENU, prices: { ...MENU.prices, staff: 4 } };
  const ok = validateOrder(ORDER, menu, BEFORE, true);
  assert.deepEqual(ok.order.children.map((c) => c.price), [4, 4]); // 3rd grade and K4 alike
  assert.equal(ok.order.total, 12); assert.ok(ok.order.staffPrice);
  const rows = orderRows(ok.order, menu, "ABCDE", new Date(BEFORE));
  assert.equal(rows[0][9], 4); assert.equal(rows[0][17], "Staff family price");
  assert.match(confirmationMail(ok.order, menu, "ABCDE", ENV0).message.body.content, /Total: \$12\.00/);
});

// ------------------------------------------------------------------ Monday.com copies
import { BOARDS, itemFor, setupMonday } from "../src/monday.js";

test("monday item: name from the row, values by column title, formula guard dropped", () => {
  const cols = [{ id: "status", title: "Status", type: "status" }, { id: "t1", title: "Email", type: "text" },
    { id: "l1", title: "Notes", type: "long_text" }, { id: "t2", title: "Signup #", type: "text" }];
  const row = ["1/5/2027 9:00 AM", "1/28/2027", "1:00 PM", "'=Pat Smith", "pat@example.com", "301", "Sam (3rd)", "Hi", "ABCDE"];
  const item = itemFor("tour", row, cols);
  assert.equal(item.name, "=Pat Smith");
  assert.deepEqual(item.values, { t1: "pat@example.com", l1: { text: "Hi" }, t2: "ABCDE", status: { label: "Signed up" } });
  assert.equal(BOARDS.inquiry.columns[2], "Status");
});

test("monday setup makes a board once and only the missing columns", async () => {
  const calls = [];
  globalThis.fetch = async (url, init = {}) => {
    const q = JSON.parse(init.body); calls.push(q);
    assert.equal(String(url), "https://api.monday.com/v2");
    assert.equal(init.headers.Authorization, "tok");
    const ok = (data) => new Response(JSON.stringify({ data }), { status: 200 });
    if (q.query.includes("create_board")) return ok({ create_board: { id: 77 } });
    if (q.query.includes("create_column")) return ok({ create_column: { id: "c" } });
    if (q.query.includes("users(emails")) return ok({ users: [{ id: 5, email: "IBA@school.test" }] });
    if (q.query.includes("add_users_to_board")) return ok({ add_users_to_board: [{ id: 5 }] });
    if (q.query.includes("boards(ids")) return ok({ boards: [{ id: "77", columns: [{ id: "status", title: "Status", type: "status" }] }] });
    return new Response("{}", { status: 500 });
  };
  const env = { MONDAY_TOKEN: "tok", MONDAY_SHARE: "iba@school.test, ibs@school.test", LUNCH_KV: kv() };
  const made1 = await setupMonday(env, "inquiry");
  assert.equal(made1.board, "77");
  assert.equal(made1.columns_added, BOARDS.inquiry.columns.length - 2); // not the name or the board's own Status
  assert.ok(calls.length < 50); // a Worker may make only 50 outside requests at a time
  assert.deepEqual(made1.shared_with, ["iba@school.test"]);
  assert.deepEqual(made1.not_on_monday, ["ibs@school.test"]);
  assert.deepEqual(calls.find((c) => c.query.includes("add_users_to_board")).variables.u, ["5"]);
  assert.equal(calls.filter((c) => c.query.includes("create_board")).length, 1);
  const made = calls.filter((c) => c.query.includes("create_column")).map((c) => c.variables.t);
  assert.ok(!made.includes("Status") && !made.includes("Parent") && made.includes("Heard about IBA"));
  assert.equal(calls.find((c) => c.variables.t === "Heard about IBA").variables.c, "long_text");
  calls.length = 0;
  await setupMonday(env, "inquiry");
  assert.equal(calls.filter((c) => c.query.includes("create_board")).length, 0);
});

test("POST /inquiry copies the row to Monday, and a Monday failure doesn't stop the form", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: BEFORE });
  for (const mondayOk of [true, false]) {
    const items = [];
    globalThis.fetch = async (url, init = {}) => {
      url = String(url);
      const ok = (o, s = 200) => new Response(JSON.stringify(o), { status: s, headers: { "Content-Type": "application/json" } });
      if (url.includes("/oauth2/v2.0/token")) return ok({ access_token: "at", refresh_token: "rt1", expires_in: 3600 });
      if (url.includes("/workbook/")) return ok({}, 201);
      if (url.endsWith("/me/sendMail")) return new Response(null, { status: 202 });
      if (url === "https://api.monday.com/v2") {
        if (!mondayOk) return ok({ errors: [{ message: "down" }] });
        const q = JSON.parse(init.body);
        if (q.query.includes("boards(ids")) {
          return ok({ data: { boards: [{ id: "9", columns: [{ id: "st", title: "Status", type: "status" }, { id: "s1", title: "Student", type: "text" }] }] } });
        }
        items.push(q.variables); return ok({ data: { create_item: { id: 1 } } });
      }
      return ok({ error: "unexpected " + url }, 500);
    };
    const res = await worker.fetch(new Request("https://lunch.test/inquiry", {
      method: "POST", body: JSON.stringify(INQ),
      headers: { "Content-Type": "application/json", Origin: "https://guarantee418.github.io", "CF-Connecting-IP": `8.8.4.${mondayOk ? 1 : 2}` },
    }), { ...INENV, MONDAY_TOKEN: "tok", LUNCH_KV: kv({ ms_refresh: "rt1", "monday:board:inquiry": "9" }) });
    assert.equal(res.status, 200);
    if (mondayOk) {
      assert.equal(items.length, 1);
      assert.equal(items[0].b, "9");
      const v = JSON.parse(items[0].v);
      assert.deepEqual(v.st, { label: "New" });
      assert.match(v.s1, /\S/);
    }
  }
});
