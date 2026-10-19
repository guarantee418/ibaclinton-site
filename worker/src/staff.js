// Staff families: children of IBA staff pay the staff price for hot lunch (Garrett, 2026-10-19).
//
// The list lives on the "Staff Families" tab of the office workbook (one family per row: last name and, if
// known, the parent's email), so the office keeps it up to date; START only fills a new tab once. An
// order is a staff family's when the parent's email is listed, or the parent's or a child's last name is.
// The website never shows the staff price (no total on the form or thank-you); the confirmation email has it.

export const STAFF_SHEET = "Staff Families";
const HEAD = ["Family (last name)", "Parent email (optional)", "Notes"];
const START = [["Kirby", "", ""], ["Buck", "", ""], ["Mora", "", "Andy Mora"]]; // Garrett's list, 2026-10-19

export async function ensureStaffTab(call) {
  const r = await call("GET", "/worksheets?$select=name");
  if (!r.ok) throw new Error(`worksheets ${r.status}`);
  if (((await r.json()).value || []).some((w) => w.name === STAFF_SHEET)) return false;
  await must(call("POST", "/worksheets/add", { name: STAFF_SHEET }));
  const s = `/worksheets/${encodeURIComponent(STAFF_SHEET)}`;
  await must(call("PATCH", `${s}/range(address='A1:C${1 + START.length}')`, { values: [HEAD, ...START] }));
  await call("PATCH", `${s}/range(address='A1:C1')/format/font`, { bold: true });
  for (const [c, w] of [["A", 150], ["B", 220], ["C", 240]]) {
    await call("PATCH", `${s}/range(address='${c}:${c}')/format`, { columnWidth: w });
  }
  return true;
}

export async function readStaffFamilies(call) {
  const r = await call("GET", `/worksheets/${encodeURIComponent(STAFF_SHEET)}/usedRange?$select=values`);
  if (r.status === 404) return [];
  if (!r.ok) throw new Error(`staff families ${r.status}`);
  const [, ...rows] = (await r.json()).values || [];
  return rows.map((x) => ({ name: norm(x[0]), email: String(x[1] || "").trim().toLowerCase() }))
    .filter((f) => f.name || f.email);
}

// Is this order (parent + children as typed on the form) from a staff family on the list?
export function isStaffFamily(families, parent, childNames) {
  if (!families || !families.length) return false;
  const email = String(parent.email || "").trim().toLowerCase();
  const lasts = [parent.last, ...childNames.map((n) => String(n || "").trim().split(/\s+/).pop())].map(norm).filter(Boolean);
  return families.some((f) => (f.email && f.email === email) || (f.name && lasts.includes(f.name)));
}

// "Kirby", "the Kirbys", "Kirby family" all mean kirby.
function norm(v) {
  return String(v || "").toLowerCase().replace(/^the\s+/, "").replace(/\s+family$/, "").replace(/[^a-z'\- ]/g, "")
    .trim().replace(/s$/, (m, i, s) => (s.length > 3 && !/ss$/.test(s) ? "" : m));
}

async function must(p) {
  const r = await p;
  if (!r.ok) throw new Error(`workbook ${r.status} ${(await r.text()).slice(0, 200)}`);
  return r;
}
