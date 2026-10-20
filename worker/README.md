# Hot lunch order service (Cloudflare Worker)

The Worker behind the site's forms. It receives each submission, checks it against the published
schedule/menu/form, writes a row to the office's Excel workbook, and emails a confirmation.

Handles: hot lunch orders (`src/index.js`), homeroom lunch lists (`src/homerooms.js`), the Gradelink
billing tab (`src/gradelink.js`), parent-teacher conferences (`src/conferences.js`), school tours
(`src/tours.js`), transcript requests (`src/transcripts.js`), Request Info inquiries (`src/inquiries.js`),
absence reports (`src/absences.js`) and office-built forms (`src/forms.js`). Optional Monday.com copies in
`src/monday.js`.

- Config: `wrangler.toml` (non-secret vars only).
- Deploy: Actions → "Deploy hot lunch order service".
- Tests: `node --test worker/test/worker.test.mjs` (no network; Microsoft is faked).

**Setup, accounts, secrets, mailbox details and upkeep are documented privately**, not in this public
repo — see the school's private `ibaclinton` operations repo (`docs/ibaclinton-site/hot-lunch-worker.md`).
Never add credentials, account IDs, mailbox details, exports or private notes here.
