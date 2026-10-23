// A copy of each Request Info inquiry, tour sign-up and transcript request on Monday.com (Garrett, 2026-10-19).
//
// The workbook stays the record: the forms save there first, and the Monday item is added after the family
// already has their answer, so Monday being down or misconfigured never stops a form. It is one way only:
// changes made on Monday don't reach the workbook.
//
// Needs the MONDAY_TOKEN secret (a Monday personal API token; items show as made by its owner). The "monday"
// admin action makes one board per form when missing (MONDAY_BOARD_KIND, "private" by default: only the token's
// owner sees it until they share it) and adds any missing column; the board ids are kept in KV.
// The deploy calls it once per board: a Worker may make only 50 outside requests per call.
// Columns are matched by title, so the office can add their own columns or rename the board.

import { COLUMNS as INQUIRY_COLUMNS } from "./inquiries.js";
import { COLUMNS as TOUR_COLUMNS } from "./tours.js";
import { COLUMNS as TRANSCRIPT_COLUMNS } from "./transcripts.js";

const API = "https://api.monday.com/v2";

// Each form: its board name, the workbook column used as the item name, the number column, its first Status
// (when the workbook row has none), and the columns that hold longer text.
export const BOARDS = {
  inquiry: { name: "Website: Info Requests", columns: INQUIRY_COLUMNS, title: "Parent", number: "Inquiry #",
    long: ["Other children", "Current school", "Heard about IBA"] },
  tour: { name: "Website: Tour Sign-ups", columns: TOUR_COLUMNS, title: "Parent", number: "Signup #",
    status: "Signed up", long: ["Children and grades", "Notes"] },
  transcript: { name: "Website: Transcript Requests", columns: TRANSCRIPT_COLUMNS, title: "Student", number: "Request #",
    long: ["Reason", "Mailing address", "Notes"] },
};

function columnType(form, title) {
  if (title === "Status") return "status";
  return form.long.includes(title) ? "long_text" : "text";
}

async function api(env, query, variables = {}) {
  const r = await fetch(API, {
    method: "POST",
    headers: { Authorization: env.MONDAY_TOKEN, "Content-Type": "application/json", "API-Version": "2026-04" },
    body: JSON.stringify({ query, variables }),
  });
  const out = await r.json().catch(() => ({}));
  if (!r.ok || out.errors || out.error_message) {
    throw new Error(`monday ${r.status} ${JSON.stringify(out.errors || out.error_message || out).slice(0, 300)}`);
  }
  return out.data;
}

async function columns(env, board) {
  const d = await api(env, "query ($b: [ID!]) { boards(ids: $b) { id columns { id title type } } }", { b: [board] });
  const b = (d.boards || [])[0];
  return b ? b.columns : null;
}

// Makes one form's board and the columns it is missing; returns {board: "123", columns_added: n}.
export async function setupMonday(env, kind) {
  const form = BOARDS[kind];
  let board = await env.LUNCH_KV.get(`monday:board:${kind}`);
  let have = board ? await columns(env, board) : null;
  if (!have) {
    const d = await api(env, "mutation ($n: String!, $k: BoardKind!) { create_board(board_name: $n, board_kind: $k, empty: true) { id } }",
      { n: form.name, k: env.MONDAY_BOARD_KIND || "private" });
    board = String(d.create_board.id);
    await env.LUNCH_KV.put(`monday:board:${kind}`, board);
    have = (await columns(env, board)) || [];
  }
  const titles = new Set(have.map((c) => c.title));
  let added = 0;
  for (const title of form.columns) {
    if (title === form.title || titles.has(title)) continue;
    await api(env, "mutation ($b: ID!, $t: String!, $c: ColumnType!) { create_column(board_id: $b, title: $t, column_type: $c) { id } }",
      { b: board, t: title, c: columnType(form, title) });
    added++;
  }
  return { board, columns_added: added, ...(await share(env, board)) };
}

// Adds the people in MONDAY_SHARE (Monday account emails, comma-separated) to the board as members.
async function share(env, board) {
  const emails = String(env.MONDAY_SHARE || "").split(",").map((e) => e.trim().toLowerCase()).filter(Boolean);
  if (!emails.length) return {};
  const d = await api(env, "query ($e: [String]) { users(emails: $e) { id email } }", { e: emails });
  const users = (d.users || []).filter((u) => emails.includes(String(u.email).toLowerCase()));
  if (users.length) {
    await api(env, "mutation ($b: ID!, $u: [ID!]!) { add_users_to_board(board_id: $b, user_ids: $u, kind: subscriber) { id } }",
      { b: board, u: users.map((u) => String(u.id)) });
  }
  const found = users.map((u) => String(u.email).toLowerCase());
  return { shared_with: found, not_on_monday: emails.filter((e) => !found.includes(e)) };
}

// The item for one workbook row (the same values, in the workbook's column order).
export function itemFor(kind, row, cols) {
  const form = BOARDS[kind];
  const byTitle = Object.fromEntries(cols.map((c) => [c.title, c]));
  const plain = (v) => String(v == null ? "" : v).replace(/^'/, ""); // the workbook's formula guard
  const values = {};
  let name = "";
  form.columns.forEach((title, i) => {
    const v = plain(row[i]);
    if (title === form.title) { name = v; return; }
    const c = byTitle[title];
    if (!c || !v) return;
    if (c.type === "status") values[c.id] = { label: v.slice(0, 40) };
    else if (c.type === "long_text") values[c.id] = { text: v };
    else if (c.type === "text") values[c.id] = v;
  });
  if (form.status && byTitle.Status && !values[byTitle.Status.id]) values[byTitle.Status.id] = { label: form.status };
  return { name: name || plain(row[form.columns.indexOf(form.number)]) || "Website form", values };
}

// Adds the item; never throws (the form already succeeded), logs instead.
export async function addToMonday(env, kind, row) {
  if (!env.MONDAY_TOKEN) return null;
  try {
    const board = await env.LUNCH_KV.get(`monday:board:${kind}`);
    if (!board) { console.error("monday: no board yet for", kind, "(run the deploy workflow)"); return null; }
    const item = itemFor(kind, row, (await columns(env, board)) || []);
    const d = await api(env, "mutation ($b: ID!, $n: String!, $v: JSON!) { create_item(board_id: $b, item_name: $n, " +
      "column_values: $v, create_labels_if_missing: true) { id } }", { b: board, n: item.name, v: JSON.stringify(item.values) });
    return String(d.create_item.id);
  } catch (err) {
    console.error("monday", kind, err && err.message || err);
    return null;
  }
}

// Removes the item with this number (used after a TEST row was deleted from the workbook).
export async function deleteFromMonday(env, kind, id) {
  if (!env.MONDAY_TOKEN) return 0;
  try {
    const board = await env.LUNCH_KV.get(`monday:board:${kind}`);
    if (!board) return 0;
    const col = ((await columns(env, board)) || []).find((c) => c.title === BOARDS[kind].number);
    if (!col) return 0;
    const d = await api(env, "query ($b: ID!, $c: String!, $v: String!) { items_page_by_column_values(board_id: $b, limit: 5, " +
      "columns: [{column_id: $c, column_values: [$v]}]) { items { id } } }", { b: board, c: col.id, v: id });
    const items = (d.items_page_by_column_values || {}).items || [];
    for (const it of items) await api(env, "mutation ($i: ID!) { delete_item(item_id: $i) { id } }", { i: it.id });
    return items.length;
  } catch (err) {
    console.error("monday delete", kind, err && err.message || err);
    return 0;
  }
}

// ------------------------------------------------------------------ forms with "Copy to monday.com" (Teacher Portal)
// Each such form gets its own board ("Teacher Portal: Activity Request"), made by its first response and kept in KV.
// Columns: the response number, received time, email, Status, one status column per approval ("Waiting for reply",
// or "Not needed" when its answer wasn't given; the office marks Approved), then one per question. A question
// added later adds its column on the next response. All missing columns are made in one request, so a response
// stays well under the Worker's 50 outside requests.
export const FORM_STATUS = "New", WAITING = "Waiting for reply", NOT_NEEDED = "Not needed";

export function formBoardColumns(form) {
  return [
    { title: "Response #", type: "text" }, { title: "Received", type: "text" }, { title: "Email", type: "text" },
    { title: "Status", type: "status" },
    ...(form.approvals || []).map((a) => ({ title: a.title, type: "status" })),
    ...form.questions.map((q) => ({ title: q.label, type: q.type === "long" ? "long_text" : "text" })),
  ];
}

async function formBoard(env, form) {
  const key = `monday:form:${form.id}`;
  let board = await env.LUNCH_KV.get(key);
  let have = board ? await columns(env, board) : null;
  if (!have) {
    const d = await api(env, "mutation ($n: String!, $k: BoardKind!) { create_board(board_name: $n, board_kind: $k, empty: true) { id } }",
      { n: `${String(form.id || "").startsWith("teacher-") ? "Teacher Portal" : "Website"}: ${form.name}`.slice(0, 80),
        k: env.FORM_BOARD_KIND || env.MONDAY_BOARD_KIND || "private" });
    board = String(d.create_board.id);
    await env.LUNCH_KV.put(key, board);
    await share(env, board);
    have = [];
  }
  const titles = new Set(have.map((c) => c.title.toLowerCase()));
  const missing = formBoardColumns(form).filter((c) => !titles.has(c.title.toLowerCase()));
  if (missing.length) {
    const vars = { b: board }, args = [];
    const calls = missing.map((c, i) => {
      vars[`t${i}`] = c.title.slice(0, 255);
      args.push(`$t${i}: String!`);
      return `c${i}: create_column(board_id: $b, title: $t${i}, column_type: ${c.type}) { id }`;
    });
    await api(env, `mutation ($b: ID!, ${args.join(", ")}) { ${calls.join(" ")} }`, vars);
    have = (await columns(env, board)) || [];
  }
  return { board, cols: have };
}

export function formItem(form, entry, id, received, approvals, cols) {
  const byTitle = Object.fromEntries(cols.map((c) => [c.title.toLowerCase(), c]));
  const values = {};
  const put = (title, v) => {
    const c = byTitle[String(title).toLowerCase()];
    v = String(v == null ? "" : v);
    if (!c || !v) return;
    if (c.type === "status") values[c.id] = { label: v.slice(0, 40) };
    else if (c.type === "long_text") values[c.id] = { text: v };
    else values[c.id] = v;
  };
  put("Response #", id); put("Received", received); put("Email", entry.email); put("Status", FORM_STATUS);
  for (const a of approvals) put(a.title, a.needed ? WAITING : NOT_NEEDED);
  form.questions.forEach((q, i) => put(q.label, entry.answers[i]));
  return { name: entry.name, values };
}

// Adds the response; never throws (the form already succeeded). Returns the item id or null.
export async function addFormToMonday(env, form, entry, id, received, approvals = []) {
  if (!env.MONDAY_TOKEN) return null;
  try {
    const { board, cols } = await formBoard(env, form);
    const item = formItem(form, entry, id, received, approvals, cols);
    const d = await api(env, "mutation ($b: ID!, $n: String!, $v: JSON!) { create_item(board_id: $b, item_name: $n, " +
      "column_values: $v, create_labels_if_missing: true) { id } }", { b: board, n: item.name, v: JSON.stringify(item.values) });
    return String(d.create_item.id);
  } catch (err) {
    console.error("monday form", form.id, err && err.message || err);
    return null;
  }
}

export async function deleteFormItem(env, item) {
  if (!env.MONDAY_TOKEN) return false;
  try {
    await api(env, "mutation ($i: ID!) { delete_item(item_id: $i) { id } }", { i: String(item) });
    return true;
  } catch (err) {
    console.error("monday form delete", err && err.message || err);
    return false;
  }
}
