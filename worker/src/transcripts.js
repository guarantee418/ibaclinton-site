// Transcript requests, on the same service and workbook as hot lunch, conferences and tours.
//
// Current students and alumni fill out one form (ibaclinton.com/transcripts/). Each request is one row of the
// "Requests" table on the Transcripts tab of the office workbook (the service adds that tab when it is missing);
// the office works from that tab and changes Status as it goes (for example Paid, Sent). Nothing is released
// by the service: the office sends the transcript once the fee is paid.
//
// Fees and the payment link are vars in wrangler.toml: TRANSCRIPT_FEE_ALUMNI, TRANSCRIPT_FEE_STUDENT and
// TRANSCRIPT_PAY_URL (a PaySimple payment page; empty = the office collects the fee). A current student's
// fee goes on the family's Gradelink account; an alumnus pays at the link, or the office calls.

export const SHEET = "Transcripts";
export const TABLE = "Requests";
export const COLUMNS = ["Requested", "Request #", "Status", "Fee", "Who", "Student", "Name at IBA", "Birth date",
  "Grade or class of", "Requested by", "Email", "Phone", "Transcript", "Copies", "Reason", "Delivery", "Send to",
  "Attention", "Mailing address", "Recipient email", "Notes", "Signed"];
const WIDTHS = [130, 80, 150, 60, 80, 160, 140, 90, 90, 160, 200, 110, 90, 60, 130, 90, 200, 140, 260, 200, 240, 160];

export const GRADES = ["K5", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"];
export const REASONS = ["College/University", "Employment", "Military", "Personal", "Background check", "Other"];
export const DELIVERY = { mail: "Mail", email: "Email (PDF)", pickup: "Pick up" };

export async function ensureTranscriptTab(call) {
  const r = await call("GET", "/worksheets?$select=name");
  if (!r.ok) throw new Error(`worksheets ${r.status}`);
  if (((await r.json()).value || []).some((w) => w.name === SHEET)) return false;
  await must(call("POST", "/worksheets/add", { name: SHEET }));
  const s = `/worksheets/${SHEET}`;
  const last = String.fromCharCode(64 + COLUMNS.length);
  // Text columns, so Excel doesn't turn a grade like "10" or a birth date into something else.
  await must(call("PATCH", `${s}/range(address='H1:I2')`, { numberFormat: [["@", "@"], ["@", "@"]] }));
  await must(call("PATCH", `${s}/range(address='A1:${last}1')`, { values: [COLUMNS] }));
  const t = await must(call("POST", `${s}/tables/add`, { address: `A1:${last}1`, hasHeaders: true }));
  await must(call("PATCH", `/tables/${encodeURIComponent((await t.json()).id)}`, { name: TABLE }));
  for (let i = 0; i < COLUMNS.length; i++) {
    const c = String.fromCharCode(65 + i);
    await call("PATCH", `${s}/range(address='${c}:${c}')/format`, { columnWidth: WIDTHS[i] });
  }
  const note = String.fromCharCode(66 + COLUMNS.length);
  await call("PATCH", `${s}/range(address='${note}1:${note}4')`, {
    values: [["Transcript requests from the website land here by themselves, one row per request."],
      ["Change Status as you go (for example Paid, then Sent). Release a transcript only once its fee is paid."],
      ["Current students: post the fee on the family's Gradelink account. Alumni pay online or by phone with the office."],
      ["Keep the column names as they are."]],
  });
  return true;
}

// What the page needs to show: the fees and whether alumni can pay online.
export function settings(env) {
  return { fee_alumni: fee(env.TRANSCRIPT_FEE_ALUMNI), fee_student: fee(env.TRANSCRIPT_FEE_STUDENT),
    pay_url: payUrl(env) };
}

export function validateRequest(body, now) {
  const str = (v, max) => String(v == null ? "" : v).replace(/[\u0000-\u001f]/g, " ").trim().slice(0, max);
  const who = str(body.who, 10);
  if (who !== "student" && who !== "alumni") return { error: "Please choose current student or alumni" };
  const q = {
    who, first: str(body.student_first, 60), last: str(body.student_last, 60), former: str(body.former_name, 120),
    birth: str(body.birth_date, 10), grade: str(body.grade, 3), year: str(body.class_year, 4),
    relation: str(body.relation, 10), requester: str(body.requester_name, 120),
    email: str(body.email, 120).toLowerCase(), phone: str(body.phone, 30),
    official: str(body.kind, 12) === "unofficial" ? "Unofficial" : str(body.kind, 12) === "official" ? "Official" : "",
    copies: Number.parseInt(str(body.copies, 2), 10), reason: str(body.reason, 30), reason_other: str(body.reason_other, 120),
    delivery: str(body.delivery, 10), send_to: str(body.send_to, 150), attention: str(body.attention, 120),
    street: str(body.street, 120), street2: str(body.street2, 120), city: str(body.city, 60), state: str(body.state, 2).toUpperCase(),
    zip: str(body.zip, 10), recipient_email: str(body.recipient_email, 120).toLowerCase(), notes: str(body.notes, 500),
    consent: body.consent === true, signature: str(body.signature, 120),
  };
  if (!q.first || !q.last) return { error: "Please enter the student's first and last name" };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(q.birth) || !(Date.parse(q.birth) < now) || q.birth < "1940-01-01") {
    return { error: "Please enter the student's date of birth" };
  }
  if (who === "student" && !GRADES.includes(q.grade)) return { error: "Please choose the student's grade" };
  const thisYear = new Date(now).getUTCFullYear();
  if (who === "alumni" && !(/^\d{4}$/.test(q.year) && +q.year >= 1973 && +q.year <= thisYear + 1)) {
    return { error: "Please enter the year you graduated or last attended" };
  }
  // A former student asks for their own record; a parent can ask only while the student is enrolled.
  if (who === "alumni") q.relation = "student";
  if (q.relation !== "student" && q.relation !== "parent") return { error: "Please say who is making this request" };
  if (q.relation === "parent" && !q.requester) return { error: "Please enter your name (parent or guardian)" };
  if (!/^[^@\s<>"]+@[^@\s<>"]+\.[^@\s<>"]+$/.test(q.email)) return { error: "Please enter a valid email address" };
  if (!q.phone) return { error: "Please enter a phone number" };
  if (!q.official) return { error: "Please choose official or unofficial" };
  if (!(q.copies >= 1 && q.copies <= 5)) return { error: "Please choose 1 to 5 copies" };
  if (!REASONS.includes(q.reason)) return { error: "Please choose a reason for the request" };
  if (q.reason === "Other" && !q.reason_other) return { error: "Please tell us the reason for the request" };
  if (!DELIVERY[q.delivery]) return { error: "Please choose how the transcript should be delivered" };
  if (q.delivery !== "pickup" && !q.send_to) return { error: "Please enter who the transcript goes to" };
  if (q.delivery === "mail" && !(q.street && q.city && /^[A-Z]{2}$/.test(q.state) && /^\d{5}(-\d{4})?$/.test(q.zip))) {
    return { error: "Please enter the full mailing address (street, city, state and ZIP code)" };
  }
  if (q.delivery === "email" && !/^[^@\s<>"]+@[^@\s<>"]+\.[^@\s<>"]+$/.test(q.recipient_email)) {
    return { error: "Please enter the email address the transcript goes to" };
  }
  if (!q.consent || !q.signature) return { error: "Please check the release box and type your full name to sign" };
  return { request: q };
}

// The fee for this request, and the Status its row starts with.
export function feeFor(q, env) {
  return q.who === "alumni" ? fee(env.TRANSCRIPT_FEE_ALUMNI) : fee(env.TRANSCRIPT_FEE_STUDENT);
}

export function startStatus(q, env) {
  if (!feeFor(q, env)) return "New";
  return q.who === "student" ? "Fee to post in Gradelink" : "Waiting for payment";
}

export function requestRow(q, id, submitted, env) {
  const amount = feeFor(q, env);
  return [submitted, id, startStatus(q, env), amount ? money(amount) : "None", q.who === "alumni" ? "Alumni" : "Current",
    text(`${q.first} ${q.last}`), text(q.former), usDate(q.birth), q.who === "alumni" ? q.year : q.grade,
    text(q.relation === "parent" ? `${q.requester} (parent/guardian)` : "Student"), q.email, text(q.phone),
    q.official, q.copies, text(reasonText(q)), DELIVERY[q.delivery], text(q.send_to), text(q.attention),
    text(address(q)), q.recipient_email, text(q.notes), text(q.signature)];
}

// What the requester is told about paying, as plain sentences (the email escapes them).
export function payment(q, env) {
  const amount = feeFor(q, env);
  if (!amount) return [];
  if (q.who === "student") {
    return [`The ${money(amount)} transcript fee will be added to your Gradelink account. ` +
      "We'll send the transcript once it is paid."];
  }
  return payUrl(env)
    ? [`Please pay the ${money(amount)} transcript fee online. Enter your request number in the payment notes. ` +
      "We'll send the transcript once the payment comes through."]
    : [`There is a ${money(amount)} transcript fee. The school office will contact you to collect it, ` +
      "and we'll send the transcript once it is paid."];
}

export function requesterMail(q, id, env) {
  const e = escapeHtml;
  const name = q.relation === "parent" ? q.requester : `${q.first} ${q.last}`;
  const link = q.who === "alumni" && feeFor(q, env) && payUrl(env)
    ? `<p><a href="${e(payUrl(env))}">Pay the transcript fee</a> (request number ${e(id)})</p>` : "";
  const content =
    `<p>Dear ${e(name)},</p>` +
    `<p>We received your transcript request for <b>${e(q.first)} ${e(q.last)}</b>. Your request number is <b>${e(id)}</b>.</p>` +
    `<p style="margin:12px 0">${e(q.official)} transcript, ${q.copies} ${q.copies === 1 ? "copy" : "copies"}<br>` +
    `${e(deliveryText(q))}</p>` +
    payment(q, env).map((p) => `<p>${e(p)}</p>`).join("") + link +
    `<p>Please allow at least 3 business days for processing. If you have questions, reply to this email or call ` +
    `the school office at (301) 856-1616.</p>` +
    `<p>Gratefully,<br>${e(env.SIGNATURE || "Independent Baptist Academy").replace(/\n/g, "<br>")}</p>`;
  const message = {
    subject: `Transcript request ${id} received`,
    body: { contentType: "HTML", content },
    toRecipients: [{ emailAddress: { address: q.email, name } }],
  };
  if (env.REPLY_TO) message.replyTo = [{ emailAddress: { address: env.REPLY_TO } }];
  return { message, saveToSentItems: true };
}

export function officeMail(q, id, env) {
  const e = escapeHtml;
  const amount = feeFor(q, env);
  const lines = [
    `Student: ${q.first} ${q.last}` + (q.former ? ` (at IBA as ${q.former})` : ""),
    `Birth date: ${usDate(q.birth)}`,
    q.who === "alumni" ? `Alumni, class of ${q.year}` : `Current student, grade ${q.grade}`,
    `Requested by: ${q.relation === "parent" ? `${q.requester} (parent/guardian)` : "the student"}`,
    `Email: ${q.email}`, `Phone: ${q.phone}`,
    `${q.official} transcript, ${q.copies} ${q.copies === 1 ? "copy" : "copies"}, reason: ${reasonText(q)}`,
    deliveryText(q),
  ];
  if (q.notes) lines.push(`Notes: ${q.notes}`);
  lines.push(`Signed: ${q.signature}`);
  const todo = !amount ? "No fee is set for this request."
    : q.who === "student" ? `Post the ${money(amount)} transcript fee on the family's Gradelink account.`
    : payUrl(env) ? `The ${money(amount)} fee is paid online (PaySimple); check for request number ${id} before releasing.`
    : `Please contact them to collect the ${money(amount)} fee.`;
  const content =
    `<p>A transcript request came in on the website:</p>` +
    `<p style="margin:12px 0">${lines.map(e).join("<br>")}</p>` +
    `<p><b>${e(todo)}</b> Release the transcript once the fee is paid.</p>` +
    `<p>Every request is on the Transcripts tab of the office workbook (request ${e(id)}).</p>`;
  return {
    message: {
      subject: `Transcript request: ${q.first} ${q.last} (${q.who === "alumni" ? "alumni" : "current student"})`,
      body: { contentType: "HTML", content },
      toRecipients: String(env.TRANSCRIPT_NOTIFY).split(",").map((a) => ({ emailAddress: { address: a.trim() } })),
      replyTo: [{ emailAddress: { address: q.email } }],
    },
    saveToSentItems: true,
  };
}

// Removes one TEST request by its number (admin upkeep); only a row whose Student starts with "TEST".
export async function deleteTestRequest(call, id) {
  const r = await call("GET", `/tables/${TABLE}/range?$select=text`);
  if (!r.ok) return { deleted: false, error: `table ${r.status}` };
  const [head, ...rows] = (await r.json()).text || [];
  const col = Object.fromEntries((head || []).map((h, i) => [String(h).trim(), i]));
  const i = rows.findIndex((row) => String(row[col["Request #"]] || "").trim() === id);
  if (i < 0) return { deleted: false, error: "no request with that number" };
  if (!/^TEST/i.test(String(rows[i][col.Student] || "").trim())) return { deleted: false, error: "not a TEST request" };
  let d = await call("DELETE", `/tables/${TABLE}/rows/itemAt(index=${i})`);
  if (!d.ok) d = await call("DELETE", `/tables/${TABLE}/rows/${i}`);
  return d.ok ? { deleted: true } : { deleted: false, error: `delete ${d.status}` };
}

// ------------------------------------------------------------------ helpers
function fee(v) {
  const n = Number(v);
  return Number.isFinite(n) && n > 0 ? n : 0;
}

function payUrl(env) {
  const u = String(env.TRANSCRIPT_PAY_URL || "").trim();
  return u.startsWith("https://") ? u : "";
}

function reasonText(q) { return q.reason === "Other" ? `Other: ${q.reason_other}` : q.reason; }

function address(q) {
  if (q.delivery !== "mail") return "";
  return [q.street, q.street2, `${q.city}, ${q.state} ${q.zip}`].filter(Boolean).join(", ");
}

function deliveryText(q) {
  const to = [q.send_to, q.attention && `Attn: ${q.attention}`].filter(Boolean).join(", ");
  if (q.delivery === "pickup") return "Pick up at the school office";
  if (q.delivery === "email") return `By email to ${to} (${q.recipient_email})`;
  return `By mail to ${to}, ${address(q)}`;
}

function text(s) { return /^[=+\-@]/.test(s) ? `'${s}` : s; }

function money(n) { return "$" + Number(n).toFixed(2); }

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
