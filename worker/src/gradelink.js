// Gradelink billing: a "Gradelink" tab in the office workbook that lists, for one lunch week, what to charge
// each student in Gradelink (Garrett, 2026-10-20). Formulas only, read from the Orders table: one row per
// student (two orders for the same student that week are added together). Since 2026-11-02 it follows Garrett's
// own billing sheet: First Name, Last Name, Grade, Charge, youngest grade first (staff last), then by last name.
// No IFERROR around it: a broken formula shows as an error (checkGradelink reports it) instead of looking like an
// empty week. "Week of" holds text like 10/26/2026 (as the service writes it) or a date the office typed: both count.
// The tab is made only when missing (admin "setup", run by every lunch deploy); the office's edits are kept.

export const GRADELINK_SHEET = "Gradelink";
const NAVY = "#1B2A4A";
const HEAD = ["First Name", "Last Name", "Grade", "Charge"];
// Grade labels as the order form writes them, youngest first; "Staff" (a staff member's own lunch) goes last.
const GRADE_ORDER = ["Preschool (K3)", "Preschool (K4)", "Kindergarten (K5)", "1st Grade", "2nd Grade", "3rd Grade",
  "4th Grade", "5th Grade", "6th Grade", "7th Grade", "8th Grade", "9th Grade", "10th Grade", "11th Grade", "12th Grade", "Staff"];

// The yellow "Week of" boxes on the Kitchen, Homerooms and Gradelink tabs show the earliest week with orders,
// starting from a given Monday (Garrett, 2026-10-20: all three showed 10/26 while the week of 10/19 had none),
// or that Monday when no later week has orders. "Week of" may hold a date or text like 10/26/2026.
export function weekBox(from) {
  return `=LET(m,${from},d,Orders[Week of],w,INT(IF(ISNUMBER(d),d,IFERROR(DATEVALUE(d),0))),IFERROR(MIN(FILTER(w,w>=m)),m))`;
}
// The office picks another week in the yellow "Pick a week" box above it (B2; any day counts as its Monday) and
// deletes it to go back, so "Week of" (B3) is never typed over (Garrett, 2026-10-20).
const picked = (auto) => `=IF(ISNUMBER($B$2),$B$2-WEEKDAY($B$2,3),${auto.slice(1)})`;
// Kitchen and Homerooms start from this week (Saturday and Sunday: the coming week).
export const SHOW_WEEK = picked(weekBox("TODAY()-WEEKDAY(TODAY(),3)+IF(WEEKDAY(TODAY(),2)>5,7,0)"));
// Gradelink starts from this week Monday to Thursday and from the coming week on Friday (orders close Friday
// 9:00 AM before the week), so Friday's billing shows the week that just closed.
export const BILL_WEEK = picked(weekBox("TODAY()+3-WEEKDAY(TODAY()+3,3)"));
export const PICK_LABEL = "Pick a week:";
export const PICK_HINT = "← type a Monday; delete it to go back";

const W = "Orders[Week of],$B$3,Orders[Student],s,Orders[Grade],g";
export const BILLING_FORMULA =
  "=LET(d,Orders[Week of],w,IF(ISNUMBER(d),d,IFERROR(DATEVALUE(d),0)),b,IF(ISNUMBER($B$3),$B$3,IFERROR(DATEVALUE($B$3),-1))," +
  "f,(INT(w)=INT(b))*(Orders[Student]<>\"\"),m,\"No lunch orders for this week yet.\",IF(SUM(f)=0,m,LET(" +
  "k,UNIQUE(FILTER(HSTACK(Orders[Student],Orders[Grade]),f)),s,CHOOSECOLS(k,1),g,CHOOSECOLS(k,2)," +
  `n,SUMIFS(Orders[Lunches],${W}),a,SUMIFS(Orders[Total],${W}),` +
  "nm,TRIM(s),fst,IFERROR(TEXTBEFORE(nm,\" \",-1),nm),lst,IFERROR(TEXTAFTER(nm,\" \",-1),\"\")," +
  `o,IFERROR(XMATCH(g,{${GRADE_ORDER.map((x) => `"${x}"`).join(",")}}),99),` +
  "r,FILTER(HSTACK(fst,lst,g,a,o),n>0,m),IF(ROWS(r)*COLUMNS(r)=1,r," +
  "CHOOSECOLS(SORTBY(r,CHOOSECOLS(r,5),1,CHOOSECOLS(r,2),1,CHOOSECOLS(r,1),1),1,2,3,4)))))";

// The first two lists (2026-10-20: Family, Student, Grade, Parent, Lunches, Price, Charge, memo, Notes) are replaced.
const OLD_STARTS = ["=LET(f,(Orders[Week of]=$B$3)*(Orders[Student]<>\"\"),",
  "=LET(d,Orders[Week of],w,IF(ISNUMBER(d),d,IFERROR(DATEVALUE(d),0)),b,IF(ISNUMBER($B$3),$B$3,IFERROR(DATEVALUE($B$3),-1)),f,(INT(w)=INT(b))*(Orders[Student]<>\"\"),m,\"No lunch orders for this week yet.\",IF(SUM(f)=0,m,LET(k,UNIQUE(FILTER(HSTACK(Orders[Student],Orders[Grade]),f)),s,CHOOSECOLS(k,1),g,CHOOSECOLS(k,2),n,SUMIFS(Orders[Lunches],Orders[Week of],$B$3,Orders[Student],s,Orders[Grade],g),a,SUMIFS(Orders[Total],Orders[Week of],$B$3,Orders[Student],s,Orders[Grade],g),p,"];

// Headings, the list and its formats (also turns an older list into this one).
async function writeList(call) {
  const s = `/worksheets/${GRADELINK_SHEET}`;
  await must(call("PATCH", `${s}/range(address='A6:I6')`, { values: [[...HEAD, "", "", "", "", ""]] }));
  await must(call("PATCH", `${s}/range(address='A7')`, { formulas: [[BILLING_FORMULA]] }));
  await call("PATCH", `${s}/range(address='D7:D400')`, { numberFormat: Array(394).fill(['"$"#,##0.00']) });
  await call("PATCH", `${s}/range(address='E7:I400')`, { numberFormat: Array(394).fill(Array(5).fill("General")) });
  await call("POST", `${s}/range(address='E6:I6')/format/fill/clear`, {});
  await call("PATCH", `${s}/range(address='A6:D6')/format/font`, { bold: true, color: "#FFFFFF" });
  await call("PATCH", `${s}/range(address='A6:D6')/format/fill`, { color: NAVY });
  for (const [c, wd] of [["A", 120], ["B", 120], ["C", 130], ["D", 80]]) {
    await call("PATCH", `${s}/range(address='${c}:${c}')/format`, { columnWidth: wd });
  }
}

export async function ensureGradelinkTab(call) {
  const r = await call("GET", "/worksheets?$select=name");
  if (!r.ok) throw new Error(`worksheets ${r.status}`);
  if (((await r.json()).value || []).some((w) => w.name === GRADELINK_SHEET)) {
    // An older list (the service's own formula, not something the office typed) becomes this one.
    const a7 = await call("GET", `/worksheets/${GRADELINK_SHEET}/range(address='A7')?$select=formulas`);
    const f = a7.ok ? String(((await a7.json()).formulas || [[""]])[0][0]) : "";
    if (OLD_STARTS.some((o) => f.startsWith(o))) await writeList(call);
    return false;
  }
  await must(call("POST", "/worksheets/add", { name: GRADELINK_SHEET }));
  const s = `/worksheets/${GRADELINK_SHEET}`;
  await must(call("PATCH", `${s}/range(address='A1:C4')`, {
    formulas: [["Gradelink billing", "", ""], ["", "", ""],
      ["Week of (Monday)", BILL_WEEK, ""],
      ["Total to bill", "=SUMIFS(Orders[Total],Orders[Week of],$B$3)", ""]],
  }));
  await must(call("PATCH", `${s}/range(address='B3')`, { numberFormat: [["m/d/yyyy"]] }));
  await must(call("PATCH", `${s}/range(address='B4')`, { numberFormat: [['"$"#,##0.00']] }));
  await writeList(call);
  await call("PATCH", `${s}/range(address='A1')/format/font`, { bold: true, size: 14, color: NAVY });
  await call("PATCH", `${s}/range(address='A3:A4')/format/font`, { bold: true });
  await call("PATCH", `${s}/range(address='B4')/format/font`, { bold: true });
  await call("PATCH", `${s}/range(address='B3')/format/fill`, { color: "#FFF7E0" });
  await call("PATCH", `${s}/range(address='C3')/format/font`, { italic: true, color: "#666666" });
  return true;
}

// What the billing list shows now, without any names: "list (n rows)", the empty-week text or an error.
export async function checkGradelink(call) {
  const r = await call("GET", `/worksheets/${GRADELINK_SHEET}/range(address='A7:A400')?$select=values`);
  if (!r.ok) return `unreadable (${r.status})`;
  const col = ((await r.json()).values || []).map((row) => row[0]).filter((v) => v !== "" && v != null);
  if (!col.length) return "empty";
  if (String(col[0]).startsWith("#")) return `error ${col[0]}`;
  if (String(col[0]).startsWith("No lunch orders")) return "no orders this week";
  return `list (${col.length} rows)`;
}

// Admin "gradelink" check: what the list shows for one week, without names. With "week" (YYYY-MM-DD) the yellow
// box shows that Monday for a few seconds and then gets its formula (or whatever the office typed) back.
export async function tryGradelink(call, week, orderRows) {
  const s = `/worksheets/${GRADELINK_SHEET}`;
  const b = await call("GET", `${s}/range(address='B3')?$select=formulas`);
  if (!b.ok) return { error: `unreadable (${b.status})` };
  const before = ((await b.json()).formulas || [[""]])[0][0];
  const [y, m, d] = String(week || "").split("-").map(Number); // typed like the office does: 10/26/2026
  if (week) await must(call("PATCH", `${s}/range(address='B3')`, { values: [[`${m}/${d}/${y}`]] }));
  try {
    const r = await call("GET", `${s}/range(address='A7:D400')?$select=values`);
    const rows = r.ok ? ((await r.json()).values || []).filter((x) => x[0] !== "" && x[0] != null) : [];
    const first = rows.length ? String(rows[0][0]) : "";
    const t = await call("GET", `${s}/range(address='B3:B4')?$select=text`);
    const [[wk], [total]] = t.ok ? (await t.json()).text : [[""], [""]];
    return { week: wk, total_to_bill: total, order_rows_that_week: orderRows(wk),
      shows: !rows.length ? "empty" : first.startsWith("#") ? `error ${first}` : first.startsWith("No lunch orders") ? "no orders this week"
        : `list (${rows.length} students, charges $${rows.reduce((n, x) => n + (Number(x[3]) || 0), 0).toFixed(2)})` };
  } finally {
    if (week) await must(call("PATCH", `${s}/range(address='B3')`, { formulas: [[before]] }));
  }
}

// Admin "week_boxes": on the Kitchen, Homerooms and Gradelink tabs, row 2 gets the yellow "Pick a week" box and
// "Week of" (B3) its formula back (over a date someone typed), no longer yellow. A week already picked in B2 is
// kept. Returns the week each tab shows now.
export async function resetWeekBoxes(call) {
  const out = {};
  for (const [tab, f] of [["Kitchen", SHOW_WEEK], ["Homerooms", SHOW_WEEK], [GRADELINK_SHEET, BILL_WEEK]]) {
    const s = `/worksheets/${tab}`;
    const w = await call("PATCH", `${s}/range(address='B3')`, { formulas: [[f]] });
    if (!w.ok) { out[tab] = `not changed (${w.status})`; continue; }
    await call("POST", `${s}/range(address='B3')/format/fill/clear`, {});
    await call("PATCH", `${s}/range(address='C3')`, { values: [[""]] }); // the old "type a Monday here" hint
    await call("PATCH", `${s}/range(address='A2')`, { values: [[PICK_LABEL]] });
    await call("PATCH", `${s}/range(address='C2')`, { values: [[PICK_HINT]] });
    await call("PATCH", `${s}/range(address='B2')`, { numberFormat: [["m/d/yyyy"]] });
    await call("PATCH", `${s}/range(address='B2')/format/fill`, { color: "#FFF7E0" });
    await call("PATCH", `${s}/range(address='A2')/format/font`, { bold: true });
    await call("PATCH", `${s}/range(address='C2')/format/font`, { italic: true, color: "#666666" });
    const t = await call("GET", `${s}/range(address='B3')?$select=text`);
    out[tab] = t.ok ? ((await t.json()).text || [[""]])[0][0] : `unreadable (${t.status})`;
  }
  return out;
}

async function must(p) {
  const r = await p;
  if (!r.ok) throw new Error(`workbook ${r.status} ${(await r.text()).slice(0, 200)}`);
  return r;
}
