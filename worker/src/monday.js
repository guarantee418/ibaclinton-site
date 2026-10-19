// A copy of each Request Info inquiry, tour sign-up and transcript request on Monday.com (Garrett, 2026-10-19).
//
// The workbook stays the record: the forms save there first, and the Monday item is added after the family
// already has their answer, so Monday being down or misconfigured never stops a form. It is one way only:
// changes made on Monday don't reach the workbook.
//
// Needs the MONDAY_TOKEN secret (a Monday personal API token; items show as made by its owner). The setup
// action makes one board per form when missing (MONDAY_BOARD_KIND, "private" by default: only the token's
// owner sees it until they share it) and adds any missing column; the board ids are kept in KV.
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

// Makes the boards and columns that are missing; returns {inquiry: "123", ...}.
export async function setupMonday(env) {
  const out = {};
  for (const [kind, form] of Object.entries(BOARDS)) {
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
    for (const title of form.columns) {
      if (title === form.title || titles.has(title)) continue;
      await api(env, "mutation ($b: ID!, $t: String!, $c: ColumnType!) { create_column(board_id: $b, title: $t, column_type: $c) { id } }",
        { b: board, t: title, c: columnType(form, title) });
    }
    out[kind] = board;
  }
  return out;
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
