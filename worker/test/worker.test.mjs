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
