// Request Info inquiries (ibaclinton.com/request-info/), on the same service and workbook as hot lunch,
// conferences, tours and transcripts. They used to go through Formspree to Garrett's inbox only.
//
// Each inquiry is one row of the "InfoRequests" table on the Inquiries tab of the office workbook (the service
// adds that tab when it is missing); the office changes Status as it follows up (for example Called, Toured,
// Enrolled). The family gets a confirmation and INQUIRY_NOTIFY gets the details.

export const SHEET = "Inquiries";
export const TABLE = "InfoRequests";
export const COLUMNS = ["Received", "Inquiry #", "Status", "Parent", "Second parent", "Email", "Phone", "Student",
  "Age", "Grade", "Birthday", "Other children", "Current school", "Heard about IBA"];
const WIDTHS = [130, 80, 110, 160, 160, 200, 110, 160, 50, 100, 90, 240, 220, 220];

export const GRADES = ["Preschool", "Kindergarten", "1st Grade", "2nd Grade", "3rd Grade", "4th Grade", "5th Grade",
  "6th Grade", "7th Grade", "8th Grade", "9th Grade", "10th Grade", "11th Grade", "12th Grade"];

export async function ensureInquiryTab(call) {
  const r = await call("GET", "/worksheets?$select=name");
  if (!r.ok) throw new Error(`worksheets ${r.status}`);
  if (((await r.json()).value || []).some((w) => w.name === SHEET)) return false;
  await must(call("POST", "/worksheets/add", { name: SHEET }));
  const s = `/worksheets/${SHEET}`;
  const last = String.fromCharCode(64 + COLUMNS.length);
  // Text columns, so Excel leaves the birthday as typed.
  await must(call("PATCH", `${s}/range(address='K1:K2')`, { numberFormat: [["@"], ["@"]] }));
  await must(call("PATCH", `${s}/range(address='A1:${last}1')`, { values: [COLUMNS] }));
  const t = await must(call("POST", `${s}/tables/add`, { address: `A1:${last}1`, hasHeaders: true }));
  await must(call("PATCH", `/tables/${encodeURIComponent((await t.json()).id)}`, { name: TABLE }));
  for (let i = 0; i < COLUMNS.length; i++) {
    const c = String.fromCharCode(65 + i);
    await call("PATCH", `${s}/range(address='${c}:${c}')/format`, { columnWidth: WIDTHS[i] });
  }
  const note = String.fromCharCode(66 + COLUMNS.length);
  await call("PATCH", `${s}/range(address='${note}1:${note}3')`, {
    values: [["Requests from the website's Request Info page land here by themselves, one row per family."],
      ["Change Status as you follow up (for example Called, Toured, Enrolled)."],
      ["Keep the column names as they are."]],
  });
  return true;
}

export function validateInquiry(body, now) {
  const str = (v, max) => String(v == null ? "" : v).replace(/[\u0000-\u001f]/g, " ").trim().slice(0, max);
  const q = {
    first: str(body.parent_first, 60), last: str(body.parent_last, 60), parent2: str(body.parent2_name, 120),
    email: str(body.email, 120).toLowerCase(), phone: str(body.phone, 30),
    student_first: str(body.student_first, 60), student_last: str(body.student_last, 60),
    age: Number.parseInt(str(body.student_age, 2), 10), grade: str(body.student_grade, 20),
    birthday: str(body.student_birthday, 10), others: str(body.other_students, 500),
    school: str(body.current_school, 300), heard: str(body.heard_about, 300),
  };
  if (!q.first || !q.last) return { error: "Please enter your first and last name" };
  if (!/^[^@\s<>"]+@[^@\s<>"]+\.[^@\s<>"]+$/.test(q.email)) return { error: "Please enter a valid email address" };
  if (!q.phone) return { error: "Please enter a phone number" };
  if (!q.student_first || !q.student_last) return { error: "Please enter the student's first and last name" };
  if (!(q.age >= 2 && q.age <= 20)) return { error: "Please enter the student's age" };
  if (!GRADES.includes(q.grade)) return { error: "Please choose a grade" };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(q.birthday) || !(Date.parse(q.birthday) < now) || q.birthday < "2000-01-01") {
    return { error: "Please enter the student's birthday" };
  }
  if (!q.school) return { error: "Please enter the school the student attends now" };
  if (!q.heard) return { error: "Please tell us how you heard about IBA" };
  return { inquiry: q };
}

export function inquiryRow(q, id, received) {
  return [received, id, "New", text(`${q.first} ${q.last}`), text(q.parent2), q.email, text(q.phone),
    text(`${q.student_first} ${q.student_last}`), q.age, q.grade, usDate(q.birthday), text(q.others),
    text(q.school), text(q.heard)];
}

export function familyMail(q, id, env) {
  const e = escapeHtml;
  const content =
    `<p>Dear ${e(q.first)} ${e(q.last)},</p>` +
    `<p>Thank you for your interest in Independent Baptist Academy! We received your request for information ` +
    `about ${e(q.student_first)} (${e(q.grade)}), and someone from our school office will be in touch soon.</p>` +
    `<p>In the meantime, you're welcome to <a href="https://ibaclinton.com/tour/">sign up for a school tour</a> ` +
    `or call the office at (301) 856-1616. (Request number ${e(id)}.)</p>` +
    `<p>Gratefully,<br>${e(env.SIGNATURE || "Independent Baptist Academy").replace(/\n/g, "<br>")}</p>`;
  const message = {
    subject: "Thank you for your interest in Independent Baptist Academy",
    body: { contentType: "HTML", content },
    toRecipients: [{ emailAddress: { address: q.email, name: `${q.first} ${q.last}` } }],
  };
  if (env.REPLY_TO) message.replyTo = [{ emailAddress: { address: env.REPLY_TO } }];
  return { message, saveToSentItems: true };
}

export function officeMail(q, id, env) {
  const e = escapeHtml;
  const lines = [
    `Parent: ${q.first} ${q.last}` + (q.parent2 ? `; second parent: ${q.parent2}` : ""),
    `Email: ${q.email}`, `Phone: ${q.phone}`,
    `Student: ${q.student_first} ${q.student_last}, age ${q.age}, ${q.grade}, born ${usDate(q.birthday)}`,
  ];
  if (q.others) lines.push(`Other children: ${q.others}`);
  lines.push(`Current school: ${q.school}`, `Heard about IBA: ${q.heard}`);
  const content =
    `<p>A family asked for information on the website:</p>` +
    `<p style="margin:12px 0">${lines.map(e).join("<br>")}</p>` +
    `<p>Reply to this email to answer them. Every request is on the Inquiries tab of the office workbook ` +
    `(request ${e(id)}).</p>`;
  return {
    message: {
      subject: `Information request: ${q.first} ${q.last} (${q.grade})`,
      body: { contentType: "HTML", content },
      toRecipients: String(env.INQUIRY_NOTIFY).split(",").map((a) => ({ emailAddress: { address: a.trim() } })),
      replyTo: [{ emailAddress: { address: q.email, name: `${q.first} ${q.last}` } }],
    },
    saveToSentItems: true,
  };
}

// Removes one TEST inquiry by its number (admin upkeep); only a row whose Student starts with "TEST".
export async function deleteTestInquiry(call, id) {
  const r = await call("GET", `/tables/${TABLE}/range?$select=text`);
  if (!r.ok) return { deleted: false, error: `table ${r.status}` };
  const [head, ...rows] = (await r.json()).text || [];
  const col = Object.fromEntries((head || []).map((h, i) => [String(h).trim(), i]));
  const i = rows.findIndex((row) => String(row[col["Inquiry #"]] || "").trim() === id);
  if (i < 0) return { deleted: false, error: "no request with that number" };
  if (!/^TEST/i.test(String(rows[i][col.Student] || "").trim())) return { deleted: false, error: "not a TEST request" };
  let d = await call("DELETE", `/tables/${TABLE}/rows/itemAt(index=${i})`);
  if (!d.ok) d = await call("DELETE", `/tables/${TABLE}/rows/${i}`);
  return d.ok ? { deleted: true } : { deleted: false, error: `delete ${d.status}` };
}

// ------------------------------------------------------------------ helpers
function text(s) { return /^[=+\-@]/.test(s) ? `'${s}` : s; }

function usDate(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return `${m}/${d}/${y}`;
}

async function must(p) {
  const r = await p;
  if (!r.ok) throw new Error(`workbook ${r.status} ${(await r.text()).slice(0, 200)}`);
  return r;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
