# Hot lunch order service

Families order hot lunch on The Kings Corner. This Cloudflare Worker receives each order,
checks it against the published menu (`site/lunch/menu.json`, built from what the office types
in the editor), adds one row per student to the **Orders** table of the office workbook, and
emails the parent a confirmation. Lunch charges are still entered in Gradelink by hand from
the workbook (the office handles call-in changes and exceptions first).

- **Mailbox:** `ibs@ibaclinton.com` (Mrs. Napierala's office account). Confirmations
  are sent from it, replies go to `ibs@ibcministries.org`, and a copy stays in its Sent Items.
- **Workbook:** `IBA Portal Forms.xlsx` in the root of that account's OneDrive (named
  "Hot Lunch Orders.xlsx" until 2026-10; `WORKBOOK_PATH` in `wrangler.toml` must match the file name). Made by
  `make_workbook.py`; the Orders table's columns must stay in order.
- **Access:** delegated Microsoft Graph permissions (User.Read, Mail.Send, Files.ReadWrite,
  offline_access), so the Worker can act only as that one account. No password is stored;
  the account signs in once at `/connect` and the Worker keeps Microsoft's refresh token in KV.
- **Approved by Garrett (2026-10-16):** automatic confirmation emails to parents for lunch orders.
- **Homeroom lunch lists** (`src/homerooms.js`, approved by Garrett 2026-10-16): Mondays 7:00 AM
  Eastern (cron), each homeroom teacher gets one email listing their students with hot lunch that
  week, Cc `HOMEROOM_CC` (ibs@ibcministries.org). Teachers and their emails are in the workbook's
  **Teachers** tab (kept out of this public repo; the office updates it each year). Grades with no
  email, and teachers with no lunches that week, get nothing. Sent once per week (KV `homerooms:<date>`).
  The Worker also adds a live **Homerooms** tab. Both tabs are added only if missing (on each deploy and
  each Monday run). Blank teacher emails are filled from the `TEACHER_EMAILS` repo secret on each
  deploy (`K3=a@x.org,K4=...,1=...,12=...`); emails typed in the tab are never changed. Test: Actions -> "Send a test homeroom lunch email" (one week, one address, marked TEST).
- **Parent-teacher conferences** (`src/conferences.js`, built 2026-10-16, not live until Garrett approves):
  parents sign up on ibaclinton.com/conferences/ (one panel per teacher). The day, hours, session length,
  break and teachers come from `content/conferences.json` (office editor), published as
  `site/conferences/schedule.json` (`CONF_URL`, read from GitHub like the menu). `GET /conferences` returns the taken times (no names);
  `POST /conference` checks the booking, adds a row to the **Bookings** table on the workbook's
  **Conferences** tab (added if missing), emails the parent a confirmation and, when the Teachers tab has
  each teacher in the panel's email (matched by name), those teachers. High School (7th-12th) is one shared
  panel: Mr. Napierala, Mr. Damm and Mr. Greene meet each family together (`members` in the content file).
  The table decides what is taken: deleting a row
  frees the time; the office can type in call-in bookings (Time like "9:30 AM", Teacher as on the website).
  One booking per parent email per teacher. A KV lock (60 s) stops two parents taking one time at once.
- **School tours** (`src/tours.js`, built 2026-10-17): families sign up on ibaclinton.com/tour/ (every "Book a
  Tour" button goes there; it replaced the Calendly embed). Tours (date, time, families per tour, deadline)
  come from `content/tours.json` (office editor), published as `site/tours/schedule.json` (`TOUR_URL`, read
  from GitHub). `GET /tours` returns how many families each tour has (no names); `POST /tour` checks the
  sign-up, refuses it once the tour has as many sign-ups as spots, adds a row to the **Signups** table on the
  workbook's **Tours** tab (added if missing), emails the family a confirmation and the office (`TOUR_NOTIFY`)
  the details. One sign-up per email per tour; a KV lock makes sign-ups for one tour go one at a time.
- **Transcript requests** (`src/transcripts.js`, built 2026-10-18): current students and alumni ask for a
  transcript on ibaclinton.com/transcripts/ (footer link). `GET /transcripts` returns the fees and payment link;
  `POST /transcript` checks the request (student, birth date, grade or class year, delivery, signed release; a
  parent may ask only for an enrolled student), adds a row to the **Requests** table on the workbook's
  **Transcripts** tab (added if missing) with Status "Fee to post in Gradelink" (current student), "Waiting for
  payment" (alumni) or "New" (no fee), and emails the requester and the office (`TRANSCRIPT_NOTIFY`). Fees and the
  PaySimple payment page are `TRANSCRIPT_FEE_STUDENT`, `TRANSCRIPT_FEE_ALUMNI` and `TRANSCRIPT_PAY_URL` in
  `wrangler.toml`; with no payment page the office collects alumni fees. The office releases the transcript.

## One-time setup (Garrett, as Microsoft 365 admin)

1. **Workbook.** Sign in to OneDrive as `ibs@ibaclinton.com` and upload
   `IBA Portal Forms.xlsx` to the main folder (not a subfolder). Share it with whoever needs it
   (Garrett, Mrs. Mora, kitchen). The account needs a license with Exchange Online and OneDrive.
2. **Cloudflare (free).** Create an account at dash.cloudflare.com and pick a workers.dev
   subdomain (for example `ibaclinton`). The service will live at
   `https://ibaclinton-lunch.<subdomain>.workers.dev`.
   - Account ID: Workers & Pages → right-hand side.
   - API token: My Profile → API Tokens → Create Token → "Edit Cloudflare Workers" template →
     Account resources: your account → Create. Copy it once.
3. **Microsoft Entra app.** entra.microsoft.com → Applications → App registrations → New registration.
   - Name `IBA Hot Lunch Orders`, "Accounts in this organizational directory only".
   - Redirect URI: platform **Web**, `https://ibaclinton-lunch.<subdomain>.workers.dev/callback`.
   - API permissions → Add → Microsoft Graph → **Delegated**: `User.Read`, `Mail.Send`,
     `Files.ReadWrite`, `offline_access` → **Grant admin consent**.
   - Certificates & secrets → New client secret (24 months). Copy the **Value** once.
   - Overview: copy **Application (client) ID** and **Directory (tenant) ID**.
4. **GitHub secrets** (this repo → Settings → Secrets and variables → Actions):
   `CLOUDFLARE_API_TOKEN`, `CLOUDFLARE_ACCOUNT_ID`, `MS_TENANT_ID`, `MS_CLIENT_ID`,
   `MS_CLIENT_SECRET`, and `LUNCH_SETUP_KEY` (any long random phrase; it protects the sign-in link).
5. **Deploy:** Actions → "Deploy hot lunch order service" → Run workflow (on `main`). The run
   summary shows the service address; `/health` should answer `"connected":false`.
6. **Connect the mailbox:** whoever signs in as `ibs@ibaclinton.com` opens
   `https://ibaclinton-lunch.<subdomain>.workers.dev/connect?key=<LUNCH_SETUP_KEY>` and signs in.
   The page says "Connected" (it refuses any other account). `/health` then says `"connected":true`.
7. **Turn it on:** set `LUNCH_API` in `build.py` to the service address, rebuild, and publish.
   Test with a test week and an order to a staff email before families use it.

## Upkeep

- **Client secret expires** after 24 months: make a new one in Entra, update `MS_CLIENT_SECRET`,
  re-run the deploy workflow.
- If the account's password is reset or its sessions are revoked, orders fail with "Online ordering
  isn't connected yet": repeat step 6.
- When the site moves to ibaclinton.com, change `MENU_URL` in `wrangler.toml` and redeploy.
- Tests: `node --test worker/test/worker.test.mjs` (no network; Microsoft is faked).

## Monday.com copies (tours, info requests, transcripts)

Each tour sign-up, Request Info inquiry and transcript request is also added as an item on a Monday.com
board (`src/monday.js`), after the workbook row and the emails. The workbook stays the record; Monday is a
copy (changes on Monday don't flow back), and a Monday problem never stops a form (it is only logged).

1. In Monday: profile picture -> Developers -> API token -> Copy. Items show as made by that person.
2. GitHub: repo Settings -> Secrets and variables -> Actions -> New repository secret, name `MONDAY_TOKEN`.
3. Run "Deploy hot lunch order service". Its last step makes the boards if missing ("Website: Info Requests",
   "Website: Tour Sign-ups", "Website: Transcript Requests"; `MONDAY_BOARD_KIND` in `wrangler.toml`, private
   by default: share them in Monday) and any missing column. Columns are matched by title, so renaming the
   board or adding columns is fine; renaming a column means it is no longer filled.
4. The TEST cleanup workflow also deletes the matching Monday item for those three kinds.
