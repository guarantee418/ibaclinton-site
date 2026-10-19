// Upkeep of the office workbook's layout (admin actions "inspect" and "organize").
//
// inspect: what the workbook looks like, for checking before a change: each tab's name, place and size, its
// tables (headings and row count) and the notes beside them. It never returns rows, so no family details.
// Its output lands in the public Actions log, so it also leaves out the Teachers tab's names and emails.

export async function inspect(call) {
  const r = await call("GET", "/worksheets?$select=name,position,visibility");
  if (!r.ok) throw new Error(`worksheets ${r.status}`);
  const out = [];
  for (const w of (await r.json()).value || []) {
    const s = `/worksheets/${encodeURIComponent(w.name)}`;
    const used = await call("GET", `${s}/usedRange(valuesOnly=true)?$select=address`);
    const sheet = { name: w.name, position: w.position, visibility: w.visibility,
      used: used.ok ? String((await used.json()).address || "").split("!").pop() : `error ${used.status}` };
    const t = await call("GET", `${s}/tables?$select=name`);
    sheet.tables = [];
    let lastCol = 0;
    for (const tb of t.ok ? (await t.json()).value || [] : []) {
      const head = await call("GET", `/tables/${encodeURIComponent(tb.name)}/headerRowRange?$select=values`);
      const rng = await call("GET", `/tables/${encodeURIComponent(tb.name)}/range?$select=address,rowCount,columnCount`);
      const h = head.ok ? ((await head.json()).values || [[]])[0] : [];
      const g = rng.ok ? await rng.json() : {};
      const first = colNumber(String(g.address || "A1").split("!").pop().split(":")[0]);
      lastCol = Math.max(lastCol, first + (g.columnCount || h.length) - 1);
      sheet.tables.push({ name: tb.name, address: String(g.address || "").split("!").pop(), headings: h,
        rows: Math.max((g.rowCount || 1) - 1, 0) });
    }
    // Text that is safe to show: the whole tab for the guide and the kitchen counts, else the notes to the
    // right of the tables (or the top rows of a tab without a table, minus the Homerooms student list).
    let addr = "";
    if (w.name === "How to use" || w.name === "Kitchen") addr = "A1:H60";
    else if (lastCol) addr = `${colLetter(lastCol + 1)}1:${colLetter(lastCol + 4)}15`;
    else if (w.name === "Homerooms") addr = "A1:H5";
    else if (w.name === "Teachers") addr = "D1:H10"; // notes only: teacher emails stay out of the (public) logs
    else addr = "A1:H8";
    const txt = await call("GET", `${s}/range(address='${addr}')?$select=text,formulas`);
    if (txt.ok) {
      const j = await txt.json();
      sheet.cells = cellsOf(addr, j.text || [], j.formulas || []);
    }
    if (w.name === "Homerooms") {
      const k = await call("GET", `${s}/range(address='J1:L6')?$select=text`);
      if (k.ok) sheet.helper = cellsOf("J1:L6", (await k.json()).text || [], []);
    }
    out.push(sheet);
  }
  return out.sort((a, b) => a.position - b.position);
}

// Non-empty cells as {"B3": "text"} (formulas shown as "=...").
function cellsOf(addr, text, formulas) {
  const [start] = addr.split(":");
  const c0 = colNumber(start), r0 = Number(start.replace(/^[A-Z]+/, ""));
  const out = {};
  text.forEach((row, i) => row.forEach((v, j) => {
    const f = formulas[i] && formulas[i][j];
    const shown = typeof f === "string" && f.startsWith("=") ? `${v}  [${f}]` : String(v);
    if (String(v).trim() || (typeof f === "string" && f.startsWith("="))) out[`${colLetter(c0 + j)}${r0 + i}`] = shown;
  }));
  return out;
}

export function colLetter(n) {
  let s = "";
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

export function colNumber(a) {
  return String(a).replace(/[^A-Z]/g, "").split("").reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
}
