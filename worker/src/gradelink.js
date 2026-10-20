// Gradelink billing: a "Gradelink" tab in the office workbook that lists, for one lunch week, what to charge
// each student in Gradelink (Garrett, 2026-10-20). Formulas only, read from the Orders table: one row per
// student (two orders for the same student that week are added together), siblings together by family name,
// with the number of lunches, the price, the charge and a memo to paste into Gradelink. No IFERROR around it:
// a broken formula shows as an error (checkGradelink reports it) instead of looking like an empty week.
// The tab is made only when missing (admin "setup", run by every lunch deploy); the office's edits are kept.

export const GRADELINK_SHEET = "Gradelink";
const NAVY = "#1B2A4A";
const HEAD = ["Family", "Student", "Grade", "Parent", "Lunches", "Price per lunch", "Charge", "Gradelink memo", "Notes"];

// The lunch week to bill: this week Monday to Thursday, the coming week from Friday on (orders close Friday
// 9:00 AM before the week), so Friday's billing shows the week that just closed.
export const BILL_WEEK = "=TODAY()+3-WEEKDAY(TODAY()+3,3)";

const W = "Orders[Week of],$B$3,Orders[Student],s,Orders[Grade],g";
export const BILLING_FORMULA =
  "=LET(f,(Orders[Week of]=$B$3)*(Orders[Student]<>\"\"),m,\"No lunch orders for this week yet.\",IF(SUM(f)=0,m,LET(" +
  "k,UNIQUE(FILTER(HSTACK(Orders[Student],Orders[Grade]),f)),s,CHOOSECOLS(k,1),g,CHOOSECOLS(k,2)," +
  `n,SUMIFS(Orders[Lunches],${W}),a,SUMIFS(Orders[Total],${W}),` +
  "p,MAP(s,g,LAMBDA(x,y,XLOOKUP(1,f*(Orders[Student]=x)*(Orders[Grade]=y),Orders[Parent],\"\")))," +
  "nt,MAP(s,g,LAMBDA(x,y,TEXTJOIN(\"; \",TRUE,UNIQUE(FILTER(Orders[Notes],f*(Orders[Student]=x)*(Orders[Grade]=y)*(Orders[Notes]<>\"\"),\"\")))))," +
  "fam,IFERROR(TEXTAFTER(TRIM(p),\" \",-1),p)," +
  "t,HSTACK(fam,s,g,p,n,IF(n>0,a/n,0),a,\"Hot Lunch week of \"&TEXT($B$3,\"m/d/yyyy\")&\" (\"&n&IF(n=1,\" lunch)\",\" lunches)\"),nt)," +
  "r,FILTER(t,n>0,m),IF(ROWS(r)*COLUMNS(r)=1,r,SORTBY(r,CHOOSECOLS(r,1),1,CHOOSECOLS(r,2),1)))))";

export async function ensureGradelinkTab(call) {
  const r = await call("GET", "/worksheets?$select=name");
  if (!r.ok) throw new Error(`worksheets ${r.status}`);
  if (((await r.json()).value || []).some((w) => w.name === GRADELINK_SHEET)) return false;
  await must(call("POST", "/worksheets/add", { name: GRADELINK_SHEET }));
  const s = `/worksheets/${GRADELINK_SHEET}`;
  await must(call("PATCH", `${s}/range(address='A1:C4')`, {
    formulas: [["Gradelink billing", "", ""], ["", "", ""],
      ["Week of (Monday)", BILL_WEEK, "Shows this week (Monday–Thursday) or the coming week (from Friday on). Type another Monday's date to see that week."],
      ["Total to bill", "=SUMIFS(Orders[Total],Orders[Week of],$B$3)", ""]],
  }));
  await must(call("PATCH", `${s}/range(address='B3')`, { numberFormat: [["m/d/yyyy"]] }));
  await must(call("PATCH", `${s}/range(address='B4')`, { numberFormat: [['"$"#,##0.00']] }));
  await must(call("PATCH", `${s}/range(address='A6:I6')`, { values: [HEAD] }));
  await must(call("PATCH", `${s}/range(address='A7')`, { formulas: [[BILLING_FORMULA]] }));
  await call("PATCH", `${s}/range(address='F7:G400')`, { numberFormat: Array(394).fill(['"$"#,##0.00', '"$"#,##0.00']) });
  await call("PATCH", `${s}/range(address='A1')/format/font`, { bold: true, size: 14, color: NAVY });
  await call("PATCH", `${s}/range(address='A3:A4')/format/font`, { bold: true });
  await call("PATCH", `${s}/range(address='B4')/format/font`, { bold: true });
  await call("PATCH", `${s}/range(address='B3')/format/fill`, { color: "#FFF7E0" });
  await call("PATCH", `${s}/range(address='C3')/format/font`, { italic: true, color: "#666666" });
  await call("PATCH", `${s}/range(address='A6:I6')/format/font`, { bold: true, color: "#FFFFFF" });
  await call("PATCH", `${s}/range(address='A6:I6')/format/fill`, { color: NAVY });
  for (const [c, w] of [["A", 110], ["B", 160], ["C", 110], ["D", 150], ["E", 60], ["F", 90], ["G", 70], ["H", 250], ["I", 200]]) {
    await call("PATCH", `${s}/range(address='${c}:${c}')/format`, { columnWidth: w });
  }
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
    const r = await call("GET", `${s}/range(address='A7:I400')?$select=values`);
    const rows = r.ok ? ((await r.json()).values || []).filter((x) => x[0] !== "" && x[0] != null) : [];
    const first = rows.length ? String(rows[0][0]) : "";
    const t = await call("GET", `${s}/range(address='B3:B4')?$select=text`);
    const [[wk], [total]] = t.ok ? (await t.json()).text : [[""], [""]];
    return { week: wk, total_to_bill: total, order_rows_that_week: orderRows(wk),
      shows: !rows.length ? "empty" : first.startsWith("#") ? `error ${first}` : first.startsWith("No lunch orders") ? "no orders this week"
        : `list (${rows.length} students, ${rows.reduce((n, x) => n + (Number(x[4]) || 0), 0)} lunches, charges $${rows.reduce((n, x) => n + (Number(x[6]) || 0), 0).toFixed(2)})` };
  } finally {
    if (week) await must(call("PATCH", `${s}/range(address='B3')`, { formulas: [[before]] }));
  }
}

async function must(p) {
  const r = await p;
  if (!r.ok) throw new Error(`workbook ${r.status} ${(await r.text()).slice(0, 200)}`);
  return r;
}
