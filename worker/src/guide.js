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
    else if (w.name === "Gradelink") addr = "A1:I6"; // the billing list below has names
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

// ------------------------------------------------------------------ organize
// organize: tidies the workbook without touching what the website needs (file name, tab names, tables, column
// headings, rows with anything in them): tabs in a fixed order, the "How to use" guide, a short reminder beside
// each tab's table, the Kitchen week box, and empty rows left at the bottom of the tables
// after deletions (a table keeps at least one row). Safe to run again; it writes the same thing each time.

const ORDER = ["How to use", "Orders", "Kitchen", "Homerooms", "Gradelink", "Teachers", "Staff Families", "Conferences", "Tours", "Transcripts", "Inquiries", "Feedback", "Absences"];
const NAVY = "#1B2A4A";
const RED = "#B00020";
const MONDAY = "TODAY()-WEEKDAY(TODAY(),3)+IF(WEEKDAY(TODAY(),2)>5,7,0)";

// Tables whose empty rows can go: the column that is always filled on a real row, and the columns that are
// formulas (Excel fills those itself, so they don't count as something typed).
const TABLES = [
  { sheet: "Orders", table: "Orders", formulas: ["Lunches", "Total"] },
  { sheet: "Conferences", table: "Bookings", formulas: [] },
  { sheet: "Tours", table: "Signups", formulas: [] },
  { sheet: "Transcripts", table: "Requests", formulas: [] },
  { sheet: "Inquiries", table: "InfoRequests", formulas: [] },
  { sheet: "Feedback", table: "Responses", formulas: [] },
  { sheet: "Absences", table: "Reports", formulas: [] },
];

// The reminder beside each tab's table: [cell where it starts, lines].
export const NOTES = {
  Orders: ["T1", [
    "How to use this tab (full steps: How to use tab)",
    "• Each hot lunch order from the website adds one row per student, with an X (or the choice, e.g. Cheese pizza) under each day ordered.",
    "• Billing: the Gradelink tab lists what to charge each student for a week.",
    "• Phone change: type an X (or the choice) under a day, or delete it. Cancel a student's whole week: delete the row.",
    "• Use Notes for anything worth remembering. Order # is the number in the parent's email.",
    "⚠ NEVER rename the headings in row 1 (e.g. Price per lunch, Lunches, Total) or add columns inside the table: orders stop going through. Your own columns go right of Notes, with a blank column between.",
  ]],
  Kitchen: ["G1", [
    "How to use this tab (full steps: How to use tab)",
    "• Shows lunches per day for one week: K5–12th (staff included) and Preschool.",
    "• Days with choices (e.g. cheese or pepperoni pizza) are counted per choice in the Choices table below.",
    "• It shows the newest week with orders. To see another week, type its Monday (like 10/26/2026) in the yellow box.",
    "• Everything else is a formula: nothing to type.",
  ]],
  Homerooms: ["M1", [
    "How to use this tab (full steps: How to use tab)",
    "• This week's lunches by grade, youngest first. The week changes by itself (weekends show the coming week).",
    "• Teachers get the same list by email every Monday at 7:00 AM.",
    "• Nothing to type. Leave column K (the grade order) and the rows below the headings as they are.",
  ]],
  Gradelink: ["K1", [
    "How to use this tab (full steps: How to use tab)",
    "• What to charge in Gradelink for one lunch week: one row per student, brothers and sisters together by family name.",
    "• Monday–Thursday it shows this week; from Friday on (orders close Friday 9:00 AM) the coming week. Type another Monday in the yellow box to see that week.",
    "• Charge each student the Charge amount; the Gradelink memo can be copied into the charge's memo.",
    "• Grade 'Staff' is a staff member's own lunch, not a student's. Notes come from the Orders tab (for example 'Staff family price').",
    "• Phone changes made on the Orders tab show up here by themselves. Nothing to type except the yellow box.",
  ]],
  Teachers: ["E1", [
    "How to use this tab (full steps: How to use tab)",
    "• One row per grade: the homeroom teacher and email.",
    "• Every Monday at 7:00 AM each teacher gets an email listing their students with hot lunch that week.",
    "• Conference booking notices also go to the teachers listed here (matched by name).",
    "• When a teacher changes, type the new name and email on that grade's row. A grade with no email gets no list.",
    "• Keep the Grade column exactly as it is. Check this tab at the start of each school year.",
  ]],
  "Staff Families": ["E1", [
    "How to use this tab (full steps: How to use tab)",
    "• Children of IBA staff pay the staff price for hot lunch. List each staff family here, one per row.",
    "• An order counts as a staff family's when the parent's or a child's last name matches, or the parent's email does.",
    "• Those orders show no prices on the website; their rows on the Orders tab say 'Staff family price' in Notes.",
  ]],
  Conferences: ["L1", [
    "How to use this tab (full steps: How to use tab)",
    "• Each conference booking from the website adds one row.",
    "• One teacher's day: filter the Teacher column, then sort by Time.",
    "• Cancel: delete the row; that time opens again on the website.",
    "• Parent who calls: add a row at the bottom: Date like 11/16/2026, Time like 9:30 AM, Teacher as on the website.",
    "• The day, hours and teachers are changed in the website editor, not here.",
  ]],
  Tours: ["K1", [
    "How to use this tab (full steps: How to use tab)",
    "• Each tour sign-up from the website adds one row per family.",
    "• Each tour takes the number of families set in the website editor (5 unless changed); then it shows as Full.",
    "• Cancel: delete the row; that spot opens again.",
    "• Family who calls: add a row at the bottom: Tour date like 1/28/2027, Time like 1:00 PM.",
    "• New tours are added in the website editor (School Tours), not here.",
  ]],
  Transcripts: ["X1", [
    "How to use this tab (full steps: How to use tab)",
    "• Each transcript request from the website adds one row. The fee is $10.",
    "• Status 'Fee to post in Gradelink' = current student: post the fee on the family's account.",
    "• Status 'Waiting for payment' = alumni: call or email them to collect the fee.",
    "• When paid, change Status to Paid. When sent, change it to Sent and put the date in Notes.",
    "• Release a transcript only after the fee is paid. Official copies go straight to the 'Send to' recipient.",
  ]],
  Inquiries: ["P1", [
    "How to use this tab (full steps: How to use tab)",
    "• Each Request Info form from the website adds one row. The office also gets an email; Reply answers the family.",
    "• Follow up, then type where things stand in Status (for example: Called 10/20, touring 1/28).",
    "• Don't add columns to this table: the website fills it. Write longer notes to the right, outside the table.",
  ]],
  Feedback: ["L1", [
    "How to use this tab (full steps: How to use tab)",
    "• Each 'Feedback for Administration' form on The Kings Corner adds one row. No email is sent; check this tab.",
    "• Name, email and phone are optional: '(no name)' means the parent didn't sign.",
    "• Change Status as you go (for example Read, Answered) and use Notes for what was done.",
  ]],
  Absences: ["O1", [
    "How to use this tab (full steps: How to use tab)",
    "• Each absence reported on the website adds one row. The office gets an email; ibs@ and the homeroom teacher are Cc'd.",
    "• 'Teacher told' names the teacher who was Cc'd; 'No email on Teachers tab' means tell them yourself.",
    "• Sort by First day to see who is out on a given day. Use Notes for anything worth remembering.",
  ]],
};

// The "How to use" tab: [kind, column A, column B]. Kinds: title, text, head (a section), tab (a row of the
// tabs-at-a-glance list), step (a numbered step), warn (a step in bold red), gap.
export const GUIDE = [
  ["title", "IBA Portal Forms: how to use this workbook"],
  ["text", "Everything families fill out on ibaclinton.com lands in this workbook by itself: hot lunch orders, parent-teacher conferences, school tours, transcript requests, Request Info inquiries, parents' feedback for administration and absence reports."],
  ["text", "The website adds each new entry as a row at the bottom of the table on its tab. You can read, sort, filter, edit and delete rows. Each tab also has a short reminder to the right of its table. Please follow the rules at the bottom of this page so the website keeps working."],
  ["gap"],
  ["head", "The tabs at a glance"],
  ["tab", "Orders", "Hot lunch orders: one row per student per week. Filled in by the website. Used for billing in Gradelink."],
  ["tab", "Kitchen", "Lunch counts per day for one week (K5–12th and Preschool) and the week's charges. Formulas only."],
  ["tab", "Homerooms", "This week's lunches by grade. Formulas only. Teachers get the same list by email every Monday."],
  ["tab", "Gradelink", "What to charge each student in Gradelink for one lunch week. Formulas only."],
  ["tab", "Teachers", "Each grade's homeroom teacher and email. Kept up to date by the office."],
  ["tab", "Staff Families", "Staff families whose children get the staff hot lunch price. Kept up to date by the office."],
  ["tab", "Conferences", "Parent-teacher conference bookings, one row per booking."],
  ["tab", "Tours", "School tour sign-ups, one row per family."],
  ["tab", "Transcripts", "Transcript requests from current students and alumni, with the $10 fee status."],
  ["tab", "Inquiries", "Request Info forms from families interested in enrolling."],
  ["tab", "Feedback", "Feedback for Administration from The Kings Corner, one row per response."],
  ["tab", "Absences", "Absences parents report on the website, one row per report."],
  ["text", "Forms the office adds to The Kings Corner in the website editor each get their own tab too, named after the form (see 'Forms added in the editor' below)."],
  ["gap"],
  ["head", "Hot lunch: Orders tab"],
  ["step", "1", "Parents order on The Kings Corner page of the website. Each order adds one row per student: Week of (that Monday), the parent's contact, grade, price per lunch and an X under each day ordered; on a day with choices the cell says which one (for example Pepperoni pizza). Lunches and Total add themselves up."],
  ["step", "2", "Prices: Preschool (K3 and K4) $5 a lunch, K5–12th $8, Staff $4 (staff rows say Grade 'Staff'). Orders close at 9:00 AM on the Friday before the lunch week unless the menu says otherwise."],
  ["step", "3", "Billing, each week after orders close: use the Gradelink tab (below). It adds up each student's lunches for the week."],
  ["step", "4", "Changes by phone: type an X under a day to add a lunch (on a day with choices, type the choice exactly as on the menu), or clear the cell to remove one; the totals update. To cancel a student's whole week, right-click the row number and choose Delete."],
  ["step", "5", "Use the Notes column for anything worth remembering (paid in cash, called to change, and so on)."],
  ["step", "6", "To find an order, press Ctrl+F (Cmd+F on a Mac) and type the Order # from the parent's email, or the student's name."],
  ["warn", "⚠", "NEVER rename, add, remove or move the column headings on the Orders tab. Row 1 must read exactly: Submitted, Week of, Order #, Parent, Email, Phone, Student, Grade, Preschool, Price per lunch, Mon, Tue, Wed, Thu, Fri, Lunches, Total, Notes. If one is changed, no parent can order (the website says \"We couldn't save your order\") until it is put back. Your own columns go to the right of Notes, with one blank column in between."],
  ["gap"],
  ["head", "Hot lunch: Kitchen tab"],
  ["step", "1", "Shows how many lunches to make each day of one week: K5–12th (staff are counted here), Preschool and Total, plus the week total and the week's charges."],
  ["step", "2", "It shows the newest week that has orders. To see a different week, type that week's Monday (like 10/26/2026) in the yellow box next to 'Week of (Monday)'."],
  ["step", "3", "When a day has choices (for example cheese or pepperoni pizza), the Choices table below the day counts shows how many of each to make, K5–12th and Preschool."],
  ["step", "4", "Everything else on this tab is a formula: please don't type over it."],
  ["gap"],
  ["head", "Hot lunch: Homerooms tab"],
  ["step", "1", "Shows this week's lunches grouped by grade, youngest first, taken from the Orders tab. The week changes by itself every Monday; on Saturday and Sunday it shows the coming week."],
  ["step", "2", "Every Monday at 7:00 AM each homeroom teacher gets the list for their grades by email (Cc ibs@ibcministries.org). A teacher with no lunches that week gets no email."],
  ["step", "3", "Nothing to type here. Leave column K (the grade order the list uses) and everything below the headings as it is."],
  ["gap"],
  ["head", "Hot lunch billing: Gradelink tab"],
  ["step", "1", "Shows what to charge in Gradelink for one lunch week: one row per student (two orders for the same student are added together), brothers and sisters together by family name, with Lunches, Price per lunch, Charge and a memo to copy into Gradelink. 'Total to bill' at the top is the whole week."],
  ["step", "2", "Monday to Thursday it shows this week; from Friday on (orders close Friday at 9:00 AM) it shows the coming week. To bill a different week, type that week's Monday (like 10/26/2026) in the yellow box."],
  ["step", "3", "For each row, open the student in Gradelink and post the Charge (paste the memo if you like). Grade 'Staff' is a staff member's own lunch, not a student's. Notes come from the Orders tab (for example 'Staff family price')."],
  ["step", "4", "Phone changes made on the Orders tab show up here by themselves. Everything except the yellow box is a formula: please don't type over it."],
  ["gap"],
  ["head", "Teachers tab"],
  ["step", "1", "One row per grade: Grade, Teacher and Email. The Monday lunch lists and the conference booking notices use it."],
  ["step", "2", "When a teacher changes, type the new name and email on that grade's row. A grade with no email gets no Monday list."],
  ["step", "3", "Keep the Grade column exactly as it is. Check this tab at the start of each school year."],
  ["gap"],
  ["head", "Staff Families tab"],
  ["step", "1", "Children of IBA staff pay the staff price for hot lunch (the same as a staff lunch). List each staff family here: the family's last name, and the parent's email if their children have a different last name."],
  ["step", "2", "When a parent orders, the website checks this list: if the parent's last name, a child's last name or the parent's email matches, every child on that order gets the staff price. The website shows them no prices, and the order's rows say 'Staff family price' in Notes."],
  ["step", "3", "Add or remove families as staff change; check it at the start of each school year."],
  ["gap"],
  ["head", "Parent-teacher conferences: Conferences tab"],
  ["step", "1", "Parents sign up on the website (ibaclinton.com/conferences). Each booking adds a row: date, time, teacher, student, parent, contact and booking number. The parent and the teacher both get an email."],
  ["step", "2", "To see one teacher's day: click the arrow on the Teacher heading, choose that teacher, then sort Time from smallest to largest."],
  ["step", "3", "To cancel: delete the row. That time opens again on the website within a minute."],
  ["step", "4", "For a parent who calls: add a row at the bottom of the table with Date like 11/16/2026, Time like 9:30 AM and Teacher exactly as it appears on the website. That time then shows as taken."],
  ["step", "5", "The conference day, hours, length and teachers are changed in the website editor ('Parent-Teacher Conferences'), not here."],
  ["gap"],
  ["head", "School tours: Tours tab"],
  ["step", "1", "Families sign up on the Book a Tour page. Each sign-up adds a row: tour date and time, parent, contact, children and grades, notes. The family gets a confirmation and iba@ gets the details."],
  ["step", "2", "Each tour takes the number of families set in the website editor (5 unless changed). When it is full, the website shows Full and takes no more."],
  ["step", "3", "To cancel: delete the row; that spot opens again. For a family who calls: add a row with Tour date like 1/28/2027 and Time like 1:00 PM; it takes a spot."],
  ["step", "4", "New tours are added in the website editor ('School Tours'). Past tours drop off the website by themselves."],
  ["gap"],
  ["head", "Transcript requests: Transcripts tab"],
  ["step", "1", "Current students and alumni request transcripts on the website (ibaclinton.com/transcripts). Each request adds a row; the requester gets a confirmation and iba@ gets the details. The fee is $10."],
  ["step", "2", "Status 'Fee to post in Gradelink' (current student): post the $10 Transcript Fee on the family's Gradelink account. When it is paid, change Status to Paid."],
  ["step", "3", "Status 'Waiting for payment' (alumni): call or email them to collect the $10 until online payment is set up. When it is paid, change Status to Paid."],
  ["step", "4", "Prepare the transcript from Gradelink. Official copies go straight to the 'Send to' recipient by mail or email; unofficial copies go to the requester; 'Pick up' waits at the office."],
  ["step", "5", "When it is sent, change Status to Sent and type the date in Notes. Release a transcript only after the fee is paid. The 'Signed' column is the requester's typed signature on the release."],
  ["gap"],
  ["head", "Request Info: Inquiries tab"],
  ["step", "1", "Families ask for information on the Request Info page. Each form adds a row with Status 'New'. The family gets a confirmation; iba@ and Mr. Greene get an email, and pressing Reply on it answers the family."],
  ["step", "2", "Follow up within two school days, then type where things stand in Status (for example: Called 10/20, touring 1/28, Enrolled)."],
  ["gap"],
  ["head", "Feedback for Administration: Feedback tab"],
  ["step", "1", "Parents send feedback from The Kings Corner page (the 'Feedback for Administration' section). Each response adds a row with Status 'New', a topic and the message. No email is sent, so check this tab regularly."],
  ["step", "2", "Name, email and phone are optional. A row that says '(no name)' was sent without a name; it can't be answered directly."],
  ["step", "3", "Pass anything that needs Mr. Greene's attention to him. Change Status as you go (for example Read, Answered) and type what was done in Notes."],
  ["gap"],
  ["head", "Absence reports: Absences tab"],
  ["step", "1", "Parents report an absence in The Kings Corner on the website (Report an Absence). Each report adds a row: first and last day, student, grade, reason, parent and contact. The parent gets a confirmation; iba@ gets the details, with ibs@ and the homeroom teacher Cc'd."],
  ["step", "2", "The homeroom teacher comes from the Teachers tab (matched by grade). 'Teacher told' shows who was Cc'd; 'No email on Teachers tab' means the teacher wasn't told, so please pass it on."],
  ["step", "3", "To see who is out on a day, sort or filter the First day column. Use Notes for anything worth remembering (called in, doctor's note received, and so on)."],
  ["gap"],
  ["head", "Forms added in the editor"],
  ["step", "1", "In the website editor, open The Kings Corner, open a section and fill in its Form: a name (also the tab's name here), up to 15 questions with an answer type each, and Required where needed. Every form also asks the person's name and email. Save, and the form shows in that section's pop-up a minute or two later."],
  ["step", "2", "The first response makes the form's tab in this workbook, with one column per question. Each response adds a row with Status 'New'; the person gets a confirmation email with their answers, and iba@ and ibs@ get the details (pressing Reply answers the person)."],
  ["step", "3", "Change Status as you go (for example Read, Done) and use Notes. To stop a form, turn on 'Stop taking responses' in the editor, or delete the form from the section."],
  ["warn", "⚠", "Adding a question adds a column at the end of the form's tab. Renaming the form, or renaming, removing or reordering its questions, starts a NEW tab (for example 'Field Trip 2'), so answers never land under the wrong heading. NEVER rename, move or delete the columns on a form's tab: the next response would start a new tab."],
  ["gap"],
  ["head", "Rules that keep the website working"],
  ["step", "•", "Don't rename this file or move it out of the main folder of the OneDrive."],
  ["warn", "⚠", "NEVER rename the tabs, the tables or their column headings, and never add, remove or move columns inside a table: the website stops saving that form until they are put back. Your own notes can go to the right of a table, with one blank column in between."],
  ["step", "•", "Add rows at the bottom of a table (press Tab in the last cell of the table), not below it or in between."],
  ["step", "•", "Type dates like 11/16/2026 and times like 9:30 AM, so the website can read them."],
  ["step", "•", "Rows whose names start with TEST are tests; Mr. Greene removes them."],
  ["step", "•", "Questions or something not working: Mr. Greene (ggreene@ibcministries.org)."],
];

// One step per call: each stays under Cloudflare's limit on outside calls per request (50 on the free plan).
export const STEPS = ["order", "guide", "headings", "notes", "notes2", "kitchen", "rows"];

export async function organize(call, step) {
  const list = await call("GET", "/worksheets?$select=name,position");
  if (!list.ok) throw new Error(`worksheets ${list.status}`);
  const have = new Set(((await list.json()).value || []).map((w) => w.name));
  const sheet = (name) => `/worksheets/${encodeURIComponent(name)}`;
  const done = [];

  if (step === "order") { // tabs in a fixed order
    let pos = 0;
    for (const name of ORDER) {
      if (!have.has(name)) continue;
      const r = await call("PATCH", sheet(name), { position: pos++ });
      done.push(r.ok ? name : `${name}: ${r.status}`);
    }
  } else if (step === "guide") { // the How to use tab: its text
    if (have.has("How to use")) {
      await writeGuide(call, sheet("How to use"));
      done.push(`${GUIDE.length} lines`);
    }
  } else if (step === "headings") { // ... and its section headings and warnings (a step of their own: 50 calls)
    if (have.has("How to use")) {
      await formatGuide(call, sheet("How to use"));
      done.push("headings");
    }
  } else if (step === "notes" || step === "notes2") { // the reminder beside each tab's table, half the tabs per step
    const all = Object.entries(NOTES), half = Math.ceil(all.length / 2);
    for (const [name, [start, lines]] of step === "notes" ? all.slice(0, half) : all.slice(half)) {
      if (!have.has(name)) continue;
      const col = start.replace(/\d+/g, "");
      await call("POST", `${sheet(name)}/range(address='${col}1:${col}15')/clear`, { applyTo: "All" });
      await must(call("PATCH", `${sheet(name)}/range(address='${col}1:${col}${lines.length}')`, { values: lines.map((l) => [l]) }));
      await call("PATCH", `${sheet(name)}/range(address='${col}2:${col}${lines.length}')/format/font`, { color: "#444444" });
      await call("PATCH", `${sheet(name)}/range(address='${col}1')/format/font`, { bold: true, color: NAVY });
      for (const [i, line] of lines.entries()) { // a "⚠" line is a warning: bold red
        if (line.startsWith("⚠")) await call("PATCH", `${sheet(name)}/range(address='${col}${i + 1}')/format/font`, { bold: true, color: RED });
      }
      done.push(name);
    }
  } else if (step === "kitchen") { // with no orders yet the week box showed 1/0/1900; then show this week
    if (have.has("Kitchen")) {
      const b3 = await call("GET", `${sheet("Kitchen")}/range(address='B3')?$select=formulas`);
      const f = b3.ok ? String(((await b3.json()).formulas || [[""]])[0][0]) : "";
      if (/^=MAX\(Orders\[Week of\]\)$/i.test(f)) {
        await must(call("PATCH", `${sheet("Kitchen")}/range(address='B3')`, {
          formulas: [[`=IF(COUNT(Orders[Week of])=0,${MONDAY},MAX(Orders[Week of]))`]] }));
        done.push("week box fixed");
      } else {
        done.push("week box left as it is");
      }
      done.push(...await kitchenChoices(call, sheet("Kitchen")));
    }
  } else if (step === "rows") { // empty rows left in the tables after deletions
    for (const t of TABLES) {
      if (!have.has(t.sheet)) continue;
      done.push(`${t.table}: ${await dropEmptyRows(call, t)} removed`);
    }
  } else {
    throw new Error(`unknown step ${step}`);
  }
  return done;
}

// Days can have choices (the Orders cell then holds e.g. "Pepperoni pizza" instead of "X"): the day counts
// count any filled cell, and a Choices table below them counts each choice per day for the Kitchen's week.
const KDAYS = [["Monday", "Mon"], ["Tuesday", "Tue"], ["Wednesday", "Wed"], ["Thursday", "Thu"], ["Friday", "Fri"]];
const CHOICE_ROWS = 5; // room per day; lunch.py MAX_OPTIONS
export const CHOICES_TOP = 14;

export function choiceFormula(col) {
  const c = `Orders[${col}]`, w = "Orders[Week of],$B$3";
  return `=IFERROR(LET(u,SORT(UNIQUE(FILTER(${c},(Orders[Week of]=$B$3)*(${c}<>"")*(${c}<>"X")))),` +
    `HSTACK(u,COUNTIFS(${w},${c},u,Orders[Preschool],"No"),COUNTIFS(${w},${c},u,Orders[Preschool],"Yes"),` +
    `COUNTIFS(${w},${c},u))),"")`;
}

async function kitchenChoices(call, s) {
  const done = [];
  const f = await call("GET", `${s}/range(address='C6:D10')?$select=formulas`);
  const old = f.ok ? (await f.json()).formulas || [] : [];
  if (old.length === 5 && old.every((r) => r.every((x) => /,"X",Orders\[Preschool\]/.test(String(x))))) {
    await must(call("PATCH", `${s}/range(address='C6:D10')`, {
      formulas: old.map((r) => r.map((x) => String(x).replace(',"X",Orders[Preschool]', ',"<>",Orders[Preschool]'))) }));
    done.push("day counts include choices");
  }
  const top = CHOICES_TOP, last = top + 1 + KDAYS.length * CHOICE_ROWS;
  await call("POST", `${s}/range(address='A${top}:E${last}')/clear`, { applyTo: "All" });
  const rows = [["Choices (days with more than one option)", "", "", "", ""], ["Day", "Choice", "K5–12th", "Preschool", "Total"]];
  KDAYS.forEach(([day, col]) => {
    rows.push([day, choiceFormula(col), "", "", ""]);
    for (let i = 1; i < CHOICE_ROWS; i++) rows.push(["", "", "", "", ""]);
  });
  await must(call("PATCH", `${s}/range(address='A${top}:E${last}')`, { formulas: rows }));
  await call("PATCH", `${s}/range(address='A${top}')/format/font`, { bold: true, color: NAVY });
  await call("PATCH", `${s}/range(address='A${top + 1}:E${top + 1}')/format/font`, { bold: true, color: "#FFFFFF" });
  await call("PATCH", `${s}/range(address='A${top + 1}:E${top + 1}')/format/fill`, { color: NAVY });
  await call("PATCH", `${s}/range(address='B:B')/format`, { columnWidth: 150 });
  done.push("choices table");
  return done;
}

async function writeGuide(call, s) {
  await call("POST", `${s}/range(address='A1:F300')/clear`, { applyTo: "All" });
  const n = GUIDE.length;
  await must(call("PATCH", `${s}/range(address='A1:B${n}')`, {
    values: GUIDE.map(([kind, a = "", b = ""]) => (kind === "tab" || kind === "step" || kind === "warn" ? [a, b] : [a, ""])) }));
  await call("PATCH", `${s}/range(address='A:A')/format`, { columnWidth: 110, horizontalAlignment: "Right" });
  await call("PATCH", `${s}/range(address='A:A')/format/font`, { bold: true, color: NAVY });
  await call("PATCH", `${s}/range(address='B:B')/format`, { columnWidth: 640, wrapText: true });
  await call("PATCH", `${s}/range(address='A1:B${n}')/format`, { verticalAlignment: "Top" });
  await call("POST", `${s}/range(address='A1:B${n}')/format/autofitRows`, {});
  for (let i = 0; i < n; i++) {
    const [kind, a = ""] = GUIDE[i];
    const r = `${s}/range(address='A${i + 1}:B${i + 1}')`;
    if (kind === "title") {
      await call("PATCH", `${s}/range(address='A${i + 1}')/format`, { horizontalAlignment: "Left" });
      await call("PATCH", `${s}/range(address='A${i + 1}')/format/font`, { size: 16 });
    } else if (kind === "text") { // a paragraph runs across both columns
      await call("POST", `${r}/merge`, { across: false });
      await call("PATCH", `${r}/format`, { horizontalAlignment: "Left", wrapText: true, rowHeight: 16 * Math.ceil(a.length / 115) + 4 });
      await call("PATCH", `${r}/format/font`, { bold: false, color: "#000000" });
    }
  }
  // The tabs-at-a-glance rows sit together: tab names left-aligned in one go.
  const tabs = GUIDE.map(([kind], i) => (kind === "tab" ? i + 1 : 0)).filter(Boolean);
  if (tabs.length) {
    await call("PATCH", `${s}/range(address='A${tabs[0]}:A${tabs[tabs.length - 1]}')/format`, { horizontalAlignment: "Left" });
  }
}

async function formatGuide(call, s) {
  for (let i = 0; i < GUIDE.length; i++) {
    const [kind] = GUIDE[i];
    const r = `${s}/range(address='A${i + 1}:B${i + 1}')`;
    if (kind === "head") {
      await call("PATCH", `${r}/format/fill`, { color: NAVY });
      await call("PATCH", `${r}/format/font`, { size: 12, color: "#FFFFFF" });
    } else if (kind === "warn") {
      await call("PATCH", `${r}/format/font`, { bold: true, color: RED });
    }
  }
}

// Deletes rows with nothing typed in them (formula columns aside), bottom up; a table keeps at least one row.
async function dropEmptyRows(call, t) {
  const r = await call("GET", `/tables/${t.table}/range?$select=values`);
  if (!r.ok) return 0;
  const [head, ...rows] = (await r.json()).values || [];
  const skip = new Set(t.formulas.map((h) => (head || []).indexOf(h)));
  const empty = [];
  rows.forEach((row, i) => { if (row.every((v, j) => skip.has(j) || String(v ?? "").trim() === "")) empty.push(i); });
  if (empty.length === rows.length) empty.shift(); // nothing real in the table: its first row stays
  let removed = 0;
  for (const i of empty.reverse()) {
    let d = await call("DELETE", `/tables/${t.table}/rows/itemAt(index=${i})`);
    if (!d.ok) d = await call("DELETE", `/tables/${t.table}/rows/${i}`);
    if (d.ok) removed++;
  }
  return removed;
}

async function must(p) {
  const r = await p;
  if (!r.ok) throw new Error(`workbook ${r.status} ${(await r.text()).slice(0, 200)}`);
  return r;
}
