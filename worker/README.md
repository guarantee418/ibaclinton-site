# Hot lunch order service

Families order hot lunch on the Parent Portal. This Cloudflare Worker receives each order,
checks it against the published menu (`site/lunch/menu.json`, built from what the office types
in the editor), adds one row per student to the **Orders** table of the office workbook, and
emails the parent a confirmation. Lunch charges are still entered in Gradelink by hand from
the workbook (the office handles call-in changes and exceptions first).

- **Mailbox:** `ibs@ibaclinton.com` (Mrs. Napierala's office account). Confirmations
  are sent from it, replies go to `ibs@ibcministries.org`, and a copy stays in its Sent Items.
- **Workbook:** `Hot Lunch Orders.xlsx` in the root of that account's OneDrive. Made by
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

## One-time setup (Garrett, as Microsoft 365 admin)

1. **Workbook.** Sign in to OneDrive as `ibs@ibaclinton.com` and upload
   `Hot Lunch Orders.xlsx` to the main folder (not a subfolder). Share it with whoever needs it
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
