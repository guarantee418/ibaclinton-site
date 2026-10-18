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
