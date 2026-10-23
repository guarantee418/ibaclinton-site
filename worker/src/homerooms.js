// Homeroom lunch lists: every Monday at 7:00 AM Eastern each homeroom teacher gets one email
// listing their students with hot lunch that week (from the Orders table), Cc the office.
//
// Who teaches which grade lives in the "Teachers" tab of the office workbook (grade, teacher,
// email), so staff email addresses stay out of this public repo and the office can update it
// each year. The service also adds a "Homerooms" tab (the same list, live, for the office).
// Both tabs are created only when missing; the office's edits are never overwritten.

import { SHOW_WEEK } from "./gradelink.js";

const DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri"];
export const TEACHERS_SHEET = "Teachers";
export const HOMEROOMS_SHEET = "Homerooms";

// Grade labels exactly as the order form writes them (lunch.py GRADES), youngest first,
// with the homeroom teacher as of 2026-2027. Only used to fill a new Teachers tab.
export const GRADES = [
  ["Preschool (K3)", "Miss Dove"], ["Preschool (K4)", "Mrs. Febo"], ["Kindergarten (K5)", "Ms. Schmidley"],
  ["1st Grade", "Miss Ryans"], ["2nd Grade", "Ms. Raymond"], ["3rd Grade", "Mrs. Lundberg"],
  ["4th Grade", "Mrs. Kirby"], ["5th Grade", "Mrs. Baumgartner"], ["6th Grade", "Mrs. Baumgartner"],
  ["7th Grade", "Mr. Napierala"], ["8th Grade", "Mr. Napierala"], ["9th Grade", "Mr. Damm"],
  ["10th Grade", "Mr. Damm"], ["11th Grade", "Mr. Greene"], ["12th Grade", "Mr. Greene"],
];
const GRADE_ORDER = new Map(GRADES.map(([g], i) => [g, i]));

// Day cells read Yes / No (Garrett, 2026-10-23; the Orders tab keeps its X), or the lunch choice when a day has
// choices. FILTER turns an empty Orders cell into 0, so anything that isn't text is a No.
export const HOMEROOMS_FORMULA =
  "=LET(f,Orders[Week of]=$B$3,t,CHOOSECOLS(FILTER(Orders[[Student]:[Fri]],f),2,1,5,6,7,8,9)," +
  "d,CHOOSECOLS(t,3,4,5,6,7),y,IF(ISTEXT(d)*(d<>\"\"),IF(UPPER(d)=\"X\",\"Yes\",d),\"No\"),u,HSTACK(CHOOSECOLS(t,1,2),y)," +
  "IFERROR(SORTBY(u,IFERROR(XMATCH(CHOOSECOLS(u,1),$K$6:$K$20),99),1,CHOOSECOLS(u,2),1)," +
  "\"No lunch orders for this week yet.\"))";

// Row 6, under Mon-Fri: that week's menu (Garrett, 2026-10-23). The meals come from a small list the service keeps
// on the same tab (MENU_TOP), one row per lunch week, filled in from the website's menu (homeroomMenus).
// Below row 15: the organize notes step clears M1:M15 (the tab's note sits at M1).
export const MENU_TOP = 18, MENU_FIRST = 20, MENU_LAST = 70; // M18 title, M19 headings, M20:R70 weeks
const MENU_CLEAR = 8; // rows 8-17 held the first version of the list; written blank
export const MENU_ROW = [["Menu", "", ...["N", "O", "P", "Q", "R"].map((c) =>
  `=IF($B$3="","",XLOOKUP($B$3,$M$${MENU_FIRST}:$M$${MENU_LAST},${c}$${MENU_FIRST}:${c}$${MENU_LAST},"")&"")`)]];

// What one menu day shows: the meal, or its choices ("Cheese pizza / Pepperoni pizza").
export function mealText(day) {
  if (!day) return "";
  const opts = Array.isArray(day.options) ? day.options.filter(Boolean) : [];
  return String(day.meal || opts.join(" / ") || "").trim();
}

// The rows for the menu list: every week in menu.json (or its single week), plus weeks already in the list that
// the website no longer shows (past weeks stay, so a picked past week still has its menu). Newest last, at most
// MENU_LAST - MENU_FIRST + 1 weeks.
export function menuRows(menu, existing = []) {
  const weeks = menu && Array.isArray(menu.weeks) && menu.weeks.length ? menu.weeks : menu && menu.week_start ? [menu] : [];
  const byWeek = new Map();
  for (const r of existing) {
    const iso = excelDate(r[0]);
    if (iso) byWeek.set(iso, r.slice(1, 6).map((v) => (v == null ? "" : String(v))));
  }
  for (const w of weeks) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(w.week_start || ""))) continue;
    const days = new Map((w.days || []).map((d) => [d.key, mealText(d)]));
    const off = new Set((w.no_school || []).map((d) => String(d.day || "").slice(0, 3).toLowerCase()));
    byWeek.set(w.week_start, DAYS.map((d) => days.get(d.toLowerCase()) || (off.has(d.toLowerCase()) ? "No school" : "")));
  }
  const max = MENU_LAST - MENU_FIRST + 1;
  return [...byWeek.entries()].sort(([a], [b]) => a.localeCompare(b)).slice(-max)
    .map(([iso, meals]) => { const [y, m, d] = iso.split("-").map(Number); return [`=DATE(${y},${m},${d})`, ...meals]; });
}

// Writes the menu list on the Homerooms tab (M18:R70). Returns how many weeks it holds.
export async function homeroomMenus(call, menu) {
  const s = `/worksheets/${HOMEROOMS_SHEET}`;
  const r = await call("GET", `${s}/range(address='M${MENU_FIRST}:R${MENU_LAST}')?$select=values`);
  const existing = r.ok ? ((await r.json()).values || []).filter((row) => row[0] !== "" && row[0] != null) : [];
  const rows = menuRows(menu, existing);
  const blank = ["", "", "", "", "", ""];
  const all = [...Array(MENU_TOP - MENU_CLEAR).fill(blank), ["Menus by week (the service fills this in from the website; please leave it)", "", "", "", "", ""],
    ["Week of", ...DAYS], ...rows, ...Array(MENU_LAST - MENU_FIRST + 1 - rows.length).fill(blank)];
  await must(call("PATCH", `${s}/range(address='M${MENU_CLEAR}:R${MENU_LAST}')`, { formulas: all }));
  await call("PATCH", `${s}/range(address='M${MENU_FIRST}:M${MENU_LAST}')`, { numberFormat: Array(MENU_LAST - MENU_FIRST + 1).fill(["m/d/yyyy"]) });
  await call("PATCH", `${s}/range(address='M${MENU_TOP + 1}:R${MENU_TOP + 1}')/format/font`, { bold: true });
  return rows.length;
}

// Puts the menu row under the day headings (row 6) and moves the student list down to row 7. Safe to repeat.
export async function homeroomMenuRow(call) {
  const s = `/worksheets/${HOMEROOMS_SHEET}`;
  await must(call("PATCH", `${s}/range(address='A6:G6')`, { formulas: MENU_ROW }));
  await must(call("PATCH", `${s}/range(address='A7')`, { formulas: [[HOMEROOMS_FORMULA]] }));
  await call("PATCH", `${s}/range(address='A6:G6')/format`, { wrapText: true, verticalAlignment: "Top" });
  await call("PATCH", `${s}/range(address='A6:G6')/format/font`, { italic: true, color: "#541F78" });
  await call("PATCH", `${s}/range(address='C:G')/format`, { columnWidth: 110, horizontalAlignment: "Center" });
}

// ------------------------------------------------------------------ workbook tabs
export async function ensureTabs(call) {
  const r = await call("GET", "/worksheets?$select=name");
  if (!r.ok) throw new Error(`worksheets ${r.status}`);
  const have = new Set(((await r.json()).value || []).map((w) => w.name));
  const made = [];
  if (!have.has(TEACHERS_SHEET)) {
    await must(call("POST", "/worksheets/add", { name: TEACHERS_SHEET }));
    const s = `/worksheets/${TEACHERS_SHEET}`;
    await must(call("PATCH", `${s}/range(address='A1:C16')`, {
      values: [["Grade", "Teacher", "Email"], ...GRADES.map(([g, t]) => [g, t, ""])],
    }));
    await must(call("PATCH", `${s}/range(address='E1:E3')`, {
      values: [["Homeroom teacher for each grade. Every Monday at 7:00 AM each teacher gets an email listing"],
        ["their students with hot lunch that week. Type each teacher's email in column C; a grade with no"],
        ["email gets no list. Keep the Grade column as it is. Update the teachers each school year."]],
    }));
    await call("PATCH", `${s}/range(address='A1:C1')/format/font`, { bold: true });
    await call("PATCH", `${s}/range(address='A:A')/format`, { columnWidth: 120 });
    await call("PATCH", `${s}/range(address='B:B')/format`, { columnWidth: 130 });
    await call("PATCH", `${s}/range(address='C:C')/format`, { columnWidth: 220 });
    made.push(TEACHERS_SHEET);
  }
  if (!have.has(HOMEROOMS_SHEET)) {
    await must(call("POST", "/worksheets/add", { name: HOMEROOMS_SHEET }));
    const s = `/worksheets/${HOMEROOMS_SHEET}`;
    await must(call("PATCH", `${s}/range(address='A1:C3')`, {
      formulas: [["Lunches by homeroom", "", ""], ["", "", ""],
        ["Week of (Monday)", SHOW_WEEK,
          ""]],
    }));
    await must(call("PATCH", `${s}/range(address='B3')`, { numberFormat: [["m/d/yyyy"]] }));
    await must(call("PATCH", `${s}/range(address='A5:G5')`, { values: [["Grade", "Student", ...DAYS]] }));
    await must(call("PATCH", `${s}/range(address='K5:K20')`, { values: [["Grade order"], ...GRADES.map(([g]) => [g])] }));
    await homeroomMenuRow(call);
    await call("PATCH", `${s}/range(address='A1')/format/font`, { bold: true, size: 14 });
    await call("PATCH", `${s}/range(address='A5:G5')/format/font`, { bold: true });
    await call("PATCH", `${s}/range(address='A:B')/format`, { columnWidth: 130 });
    made.push(HOMEROOMS_SHEET);
  }
  return made;
}

// What the Homerooms list shows now, without any names: "list (n rows)", the empty-week text or an error.
export async function checkHomerooms(call) {
  const r = await call("GET", `/worksheets/${HOMEROOMS_SHEET}/range(address='A7:A300')?$select=values`);
  if (!r.ok) return `unreadable (${r.status})`;
  const col = ((await r.json()).values || []).map((row) => row[0]).filter((v) => v !== "" && v != null);
  if (!col.length) return "empty";
  if (String(col[0]).startsWith("#")) return `error ${col[0]}`;
  if (String(col[0]).startsWith("No lunch orders")) return "no orders this week";
  return `list (${col.length} rows)`;
}

// ------------------------------------------------------------------ the weekly emails
export async function readWeek(call, weekIso) {
  const t = await call("GET", "/tables/Orders/range?$select=values");
  if (!t.ok) throw new Error(`orders ${t.status}`);
  const [head, ...rows] = (await t.json()).values || [];
  const col = Object.fromEntries((head || []).map((h, i) => [String(h).trim(), i]));
  for (const need of ["Week of", "Student", "Grade", ...DAYS]) {
    if (!(need in col)) throw new Error(`Orders table has no "${need}" column`);
  }
  const kids = new Map();
  for (const r of rows) {
    if (excelDate(r[col["Week of"]]) !== weekIso) continue;
    const name = String(r[col.Student] || "").trim(), grade = String(r[col.Grade] || "").trim();
    if (!name) continue;
    const key = `${grade}|${name.toLowerCase()}`;
    const k = kids.get(key) || { name, grade, days: new Set(), choices: {} };
    // A day's cell is "X", or the lunch choice the parent picked (e.g. "Pepperoni pizza").
    DAYS.forEach((d) => {
      const v = String(r[col[d]] || "").trim();
      if (!v) return;
      k.days.add(d);
      if (v.toUpperCase() !== "X") k.choices[d] = v;
    });
    kids.set(key, k);
  }
  return [...kids.values()].filter((k) => k.days.size);
}

export async function weekCounts(call) {
  const t = await call("GET", "/tables/Orders/range?$select=values");
  if (!t.ok) throw new Error(`orders ${t.status}`);
  const [head, ...rows] = (await t.json()).values || [];
  const i = (head || []).map((h) => String(h).trim()).indexOf("Week of");
  const out = {};
  for (const r of rows) {
    const raw = r[i];
    const key = excelDate(raw) || (raw === "" || raw == null ? "(blank)" : `(unreadable: ${typeof raw})`);
    out[key] = (out[key] || 0) + 1;
  }
  return out;
}

// Fill in blank emails on the Teachers tab from "K3=a@b.org,K4=...,1=...,12=..." (grade code or label).
// Emails already typed in the tab are never changed. Returns how many rows were filled.
export async function fillTeacherEmails(call, list) {
  const want = {};
  for (const part of String(list || "").split(/[,;\n]+/)) {
    const [k, v] = part.split("=").map((x) => (x || "").trim());
    if (k && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v)) want[gradeLabel(k)] = v;
  }
  const r = await call("GET", `/worksheets/${TEACHERS_SHEET}/range(address='A1:C40')?$select=values`);
  if (!r.ok) throw new Error(`teachers ${r.status}`);
  const rows = (await r.json()).values || [];
  let filled = 0;
  const col = rows.map((row, i) => {
    const have = row[2] == null ? "" : String(row[2]).trim();
    const email = want[String(row[0] || "").trim()];
    if (i > 0 && !have && email) { filled++; return [email]; }
    return [row[2] == null ? "" : row[2]];
  });
  if (filled) {
    const w = await call("PATCH", `/worksheets/${TEACHERS_SHEET}/range(address='C1:C${col.length}')`, { values: col });
    if (!w.ok) throw new Error(`teachers write ${w.status}`);
  }
  return filled;
}

function gradeLabel(k) {
  const u = k.toUpperCase();
  if (u === "K3" || u === "K4") return `Preschool (${u})`;
  if (u === "K5" || u === "K") return "Kindergarten (K5)";
  const n = Number(k);
  if (Number.isInteger(n) && n >= 1 && n <= 12) return `${n}${n === 1 ? "st" : n === 2 ? "nd" : n === 3 ? "rd" : "th"} Grade`;
  return k;
}

export async function readTeachers(call) {
  const r = await call("GET", `/worksheets/${TEACHERS_SHEET}/usedRange?$select=values`);
  if (!r.ok) throw new Error(`teachers ${r.status}`);
  const [, ...rows] = (await r.json()).values || [];
  return rows.map((r) => ({ grade: String(r[0] || "").trim(), teacher: String(r[1] || "").trim(), email: String(r[2] || "").trim() }))
    .filter((t) => t.grade);
}

// One email per teacher (a teacher with two grades gets one list). Grades with no teacher
// email, and teachers with no lunches that week, get nothing.
export function homeroomMails(students, teachers, weekIso, env, testTo) {
  const byEmail = new Map();
  for (const t of teachers) {
    if (!testTo && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(t.email)) continue;
    const key = (t.email || `(no email) ${t.teacher}`).toLowerCase();
    const g = byEmail.get(key) || { teacher: t.teacher, email: t.email, grades: [] };
    g.grades.push(t.grade);
    byEmail.set(key, g);
  }
  const label = weekLabel(weekIso);
  const mails = [];
  for (const g of byEmail.values()) {
    const kids = students.filter((s) => g.grades.includes(s.grade)).sort((a, b) =>
      (rank(a.grade) - rank(b.grade)) || a.name.localeCompare(b.name));
    if (!kids.length) continue;
    const e = escapeHtml;
    const two = g.grades.length > 1;
    const th = (x) => `<th style="text-align:left;padding:4px 10px;border-bottom:2px solid #1B2A4A">${x}</th>`;
    const td = (x, c) => `<td style="padding:4px 10px;border-bottom:1px solid #ddd${c ? ";text-align:center" : ""}">${x}</td>`;
    const rows = kids.map((k) => `<tr>${td(e(k.name))}${two ? td(e(k.grade)) : ""}${DAYS.map((d) => td(k.days.has(d) ? (k.choices && k.choices[d] ? e(k.choices[d]) : "&#10003;") : "", true)).join("")}</tr>`).join("");
    const totals = `<tr><td style="padding:4px 10px"><b>Lunches</b></td>${two ? "<td></td>" : ""}` +
      DAYS.map((d) => `<td style="padding:4px 10px;text-align:center"><b>${kids.filter((k) => k.days.has(d)).length}</b></td>`).join("") + "</tr>";
    const content =
      (testTo ? `<p style="color:#a00"><b>TEST.</b> This would go to ${e(g.teacher)} &lt;${e(g.email || "no email yet")}&gt;` +
        `${env.HOMEROOM_CC ? `, Cc ${e(env.HOMEROOM_CC)}` : ""}.</p>` : "") +
      `<p>Good morning${g.teacher ? `, ${e(g.teacher)}` : ""}!</p>` +
      `<p>These students in your homeroom (${e(g.grades.join(", "))}) have hot lunch this week, the week of ${e(label)}:</p>` +
      `<table style="border-collapse:collapse;font-size:15px"><thead><tr>${th("Student")}${two ? th("Grade") : ""}` +
      `${DAYS.map(th).join("")}</tr></thead><tbody>${rows}${totals}</tbody></table>` +
      `<p>Orders placed online by the Friday deadline are included. If something changes, the office will let you know.</p>` +
      `<p>Thank you,<br>${e(env.SIGNATURE || "Independent Baptist Academy").replace(/\n/g, "<br>")}</p>`;
    const message = {
      subject: `${testTo ? "[TEST] " : ""}Hot lunch this week: ${g.grades.length > 1 ? g.grades.map(short).join(" & ") : g.grades[0]} (week of ${label})`,
      body: { contentType: "HTML", content },
      toRecipients: [{ emailAddress: { address: testTo || g.email, name: testTo ? undefined : g.teacher } }],
    };
    if (!testTo && env.HOMEROOM_CC) message.ccRecipients = env.HOMEROOM_CC.split(",").map((a) => ({ emailAddress: { address: a.trim() } }));
    if (env.REPLY_TO) message.replyTo = [{ emailAddress: { address: env.REPLY_TO } }];
    mails.push({ teacher: g.teacher, mail: { message, saveToSentItems: true } });
  }
  return mails;
}

// The Monday of the school week to report on, if `now` is Monday 7 AM Eastern; else "".
export function mondayToSend(now) {
  const p = Object.fromEntries(new Intl.DateTimeFormat("en-US", {
    timeZone: "America/New_York", year: "numeric", month: "2-digit", day: "2-digit", weekday: "short",
    hour: "numeric", hourCycle: "h23",
  }).formatToParts(now).map((x) => [x.type, x.value]));
  if (p.weekday !== "Mon" || Number(p.hour) !== 7) return "";
  return `${p.year}-${p.month}-${p.day}`;
}

// ------------------------------------------------------------------ helpers
async function must(p) {
  const r = await p;
  if (!r.ok) throw new Error(`workbook ${r.status} ${(await r.text()).slice(0, 200)}`);
  return r;
}

// Excel stores dates as day numbers since 1899-12-30; a row may also hold the typed text "10/26/2026".
export function excelDate(v) {
  if (typeof v === "number" && v > 0) return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000).toISOString().slice(0, 10);
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(v || "").trim());
  if (m) return `${m[3]}-${m[1].padStart(2, "0")}-${m[2].padStart(2, "0")}`;
  const iso = /^(\d{4}-\d{2}-\d{2})/.exec(String(v || "").trim());
  return iso ? iso[1] : "";
}

function weekLabel(iso) {
  return new Date(`${iso}T12:00:00Z`).toLocaleDateString("en-US", { timeZone: "UTC", month: "long", day: "numeric" });
}

function rank(grade) { return GRADE_ORDER.has(grade) ? GRADE_ORDER.get(grade) : 99; }
function short(grade) { return grade.replace(/ Grade$/, "").replace(/^Preschool \((K\d)\)$/, "$1").replace(/^Kindergarten \(K5\)$/, "K5"); }

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
