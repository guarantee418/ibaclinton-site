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
  WORKBOOK_PATH: "Hot Lunch Orders.xlsx", TABLE: "Orders", TENANT_ID: "tenant", CLIENT_ID: "client",
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
    if (url === ENV0.MENU_URL) return ok(MENU);
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
  assert.ok(add.url.includes("/me/drive/root:/Hot%20Lunch%20Orders.xlsx:/workbook/tables/Orders/rows/add"));
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
