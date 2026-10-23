// Office upkeep of the Orders table (admin actions "delete_duplicate_order" and "fill_blank_week").
// Both refuse anything but the narrow case they are for, and return counts only: their output lands in a
// public Actions log, so no names.

const DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri"];
const norm = (v) => String(v == null ? "" : v).trim().toLowerCase();

async function readTable(call, t) {
  const r = await call("GET", `/tables/${t}/range?$select=text`);
  if (!r.ok) return { error: `table ${r.status}` };
  const [head, ...rows] = (await r.json()).text || [];
  return { rows, col: Object.fromEntries((head || []).map((h, i) => [String(h).trim(), i])) };
}

// Removes every row of one lunch order, but only if each row repeats a row of another order exactly: same
// Week of, Student, Grade and days (a parent who ordered the same week twice). Bottom row first.
export async function deleteDuplicateOrder(call, table, id) {
  const t = encodeURIComponent(table);
  const { rows, col, error } = await readTable(call, t);
  if (error) return { deleted: 0, error };
  for (const need of ["Order #", "Week of", "Student", "Grade", ...DAYS]) {
    if (!(need in col)) return { deleted: 0, error: `no "${need}" column` };
  }
  const num = (row) => String(row[col["Order #"]] || "").trim().toUpperCase();
  const key = (row) => ["Week of", "Student", "Grade", ...DAYS].map((c) => norm(row[col[c]])).join("|");
  const hits = rows.map((row, i) => [row, i]).filter(([row]) => num(row) === id);
  if (!hits.length) return { deleted: 0, error: "no order with that number" };
  const twin = ([row]) => rows.some((o) => num(o) !== id && key(o) === key(row));
  if (!hits.every(twin)) return { deleted: 0, error: "not a duplicate: a row has no identical row in another order" };
  let deleted = 0;
  for (const [, i] of hits.reverse()) {
    let d = await call("DELETE", `/tables/${t}/rows/itemAt(index=${i})`);
    if (!d.ok) d = await call("DELETE", `/tables/${t}/rows/${i}`);
    if (!d.ok) return { deleted, error: `delete ${d.status}` };
    deleted++;
  }
  return { deleted };
}

// A row the office typed in for a phone order but left without "Week of" is missing from the Kitchen,
// Homerooms and Gradelink tabs. Fills only blank cells of such rows: Week of with the given Monday (written the
// way the website writes its own rows) and Price per lunch with the given price.
export async function fillBlankWeek(call, table, weekIso, price) {
  const t = encodeURIComponent(table);
  const { rows, col, error } = await readTable(call, t);
  if (error) return { filled: 0, error };
  for (const need of ["Week of", "Student", "Price per lunch"]) {
    if (!(need in col)) return { filled: 0, error: `no "${need}" column` };
  }
  const [y, m, d] = weekIso.split("-").map(Number);
  const cell = (i, c) => `/tables/${t}/rows/itemAt(index=${i})/range/cell(row=0,column=${col[c]})`;
  let filled = 0;
  for (const [i, row] of rows.entries()) {
    if (norm(row[col["Week of"]]) || !norm(row[col.Student])) continue;
    const w = await call("PATCH", cell(i, "Week of"), { values: [[`${m}/${d}/${y}`]] });
    if (!w.ok) return { filled, error: `week ${w.status}` };
    if (!norm(row[col["Price per lunch"]])) {
      const p = await call("PATCH", cell(i, "Price per lunch"), { values: [[price]] });
      if (!p.ok) return { filled, error: `price ${p.status}` };
    }
    filled++;
  }
  return { filled };
}
