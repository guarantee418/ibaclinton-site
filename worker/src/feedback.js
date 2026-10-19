// Feedback for Administration (the Parent Portal section that used to link to the town hall Microsoft Form),
// on the same service and workbook as hot lunch, conferences, tours, transcripts and Request Info.
//
// Each response is one row of the "Responses" table on the Feedback tab of the office workbook (the service adds
// that tab when it is missing). Name, email and phone are optional, so a parent can write without signing.
// No email goes out: the office and Mr. Greene read the tab and update Status.

export const SHEET = "Feedback";
export const TABLE = "Responses";
export const COLUMNS = ["Received", "Feedback #", "Status", "Name", "Email", "Phone", "Grade(s)", "Topic", "Feedback", "Notes"];
const WIDTHS = [130, 85, 110, 160, 200, 110, 110, 130, 480, 240];

export const TOPICS = ["Academics", "Communication", "Discipline", "Activities & athletics", "Hot lunch", "Facilities & safety",
  "Suggestion", "Praise", "Other"];

export async function ensureFeedbackTab(call) {
  const r = await call("GET", "/worksheets?$select=name");
  if (!r.ok) throw new Error(`worksheets ${r.status}`);
  if (((await r.json()).value || []).some((w) => w.name === SHEET)) return false;
  await must(call("POST", "/worksheets/add", { name: SHEET }));
  const s = `/worksheets/${SHEET}`;
  const last = String.fromCharCode(64 + COLUMNS.length);
  await must(call("PATCH", `${s}/range(address='A1:${last}1')`, { values: [COLUMNS] }));
  const t = await must(call("POST", `${s}/tables/add`, { address: `A1:${last}1`, hasHeaders: true }));
  await must(call("PATCH", `/tables/${encodeURIComponent((await t.json()).id)}`, { name: TABLE }));
  for (let i = 0; i < COLUMNS.length; i++) {
    const c = String.fromCharCode(65 + i);
    await call("PATCH", `${s}/range(address='${c}:${c}')/format`, { columnWidth: WIDTHS[i] });
  }
  await call("PATCH", `${s}/range(address='I:I')/format`, { wrapText: true });
  const note = String.fromCharCode(66 + COLUMNS.length);
  await call("PATCH", `${s}/range(address='${note}1:${note}3')`, {
    values: [["Feedback from the Parent Portal's 'Feedback for Administration' form lands here, one row per response."],
      ["Change Status as you go (for example Read, Answered)."],
      ["Keep the column names as they are."]],
  });
  return true;
}

export function validateFeedback(body) {
  const str = (v, max) => String(v == null ? "" : v).replace(/[\u0000-\u0008\u000b-\u001f]/g, " ").trim().slice(0, max);
  const q = {
    name: str(body.name, 120).replace(/\s+/g, " "), email: str(body.email, 120).toLowerCase(), phone: str(body.phone, 30),
    grades: str(body.grades, 60), topic: str(body.topic, 40), message: str(body.message, 4000),
  };
  if (q.email && !/^[^@\s<>"]+@[^@\s<>"]+\.[^@\s<>"]+$/.test(q.email)) return { error: "Please check your email address, or leave it blank" };
  if (!TOPICS.includes(q.topic)) return { error: "Please choose a topic" };
  if (q.message.length < 5) return { error: "Please write your feedback" };
  return { feedback: q };
}

export function feedbackRow(q, id, received) {
  return [received, id, "New", text(q.name || "(no name)"), q.email, text(q.phone), text(q.grades), q.topic, text(q.message), ""];
}

// Removes one TEST response by its number (admin upkeep); only a row whose Name starts with "TEST".
export async function deleteTestFeedback(call, id) {
  const r = await call("GET", `/tables/${TABLE}/range?$select=text`);
  if (!r.ok) return { deleted: false, error: `table ${r.status}` };
  const [head, ...rows] = (await r.json()).text || [];
  const col = Object.fromEntries((head || []).map((h, i) => [String(h).trim(), i]));
  const i = rows.findIndex((row) => String(row[col["Feedback #"]] || "").trim() === id);
  if (i < 0) return { deleted: false, error: "no response with that number" };
  if (!/^TEST/i.test(String(rows[i][col.Name] || "").trim())) return { deleted: false, error: "not a TEST response" };
  let d = await call("DELETE", `/tables/${TABLE}/rows/itemAt(index=${i})`);
  if (!d.ok) d = await call("DELETE", `/tables/${TABLE}/rows/${i}`);
  return d.ok ? { deleted: true } : { deleted: false, error: `delete ${d.status}` };
}

// ------------------------------------------------------------------ helpers
function text(s) { return /^[=+\-@]/.test(s) ? `'${s}` : s; }

async function must(p) {
  const r = await p;
  if (!r.ok) throw new Error(`workbook ${r.status} ${(await r.text()).slice(0, 200)}`);
  return r;
}
