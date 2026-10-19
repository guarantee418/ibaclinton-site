# ibaclinton-site

Website files for Independent Baptist Academy, Clinton, MD (ibaclinton.com).

- `src/` is where pages are edited: `src/layout.html` (header, menu, footer shared by every page),
  `src/cta.html` (the "Book a Tour / Request Info" band) and one file per page in `src/pages/`.
- `python3 build.py` combines them into the finished site in `site/`. Commit both `src/` and `site/`.
- `site/img/` holds photos (resized, location data stripped); `site/files/` holds PDFs.
- `site/styles.css` is the one stylesheet.
- Publishing is manual: Actions tab → "Publish to GitHub Pages" → Run workflow **on `main`**. The
  github-pages environment only deploys from `main`, so merge the working branch into `main` first.
- Staff photos: 480x600 JPEG, all metadata removed, named `last-first.jpg` in `site/img/staff/`.
- Request Info form: since 2026-10-18 it goes to the same service as hot lunch (`POST /inquiry`,
  `worker/src/inquiries.js`): a row on the Inquiries tab of the office workbook, a confirmation to the family
  and the details to `INQUIRY_NOTIFY`. Formspree (form `xaeqabaz`) is no longer used.
- `build.py` also writes `sitemap.xml` and `robots.txt`. While `DOMAIN` is empty the pages carry a
  `noindex` tag so the github.io preview stays out of search results.

## Online editor for the school office (Pages CMS)

- Mrs. Mora and Mrs. Napierala edit the **Hot Lunch** menu (`content/lunch.json`) and the whole
  **Kings Corner** parents page (`content/parents.json`: the purple banner's text and button, the tiles (title,
  description, which section each opens in a pop-up (`site/kings.js`), optional tag), every section's heading, text, links and PDFs, tile and
  section order; PDFs uploaded to `site/files/`) at app.pagescms.org
  (`ibaclinton.com/admin/` links there). They are invited by email as collaborators, no GitHub account.
- `.pages.yml` defines what they can edit. Collaborators can't change it; changing it needs Garrett's OK.
- Their saves commit to `main`, and the Pages workflow publishes changes to `content/` or `site/files/`
  automatically (approved by Garrett 2026-10-16). Everything else is still published by hand.
- The workflow runs `build.py`, so the committed `site/` copy of those pages may lag behind; that's fine.
- Everything typed in the editor is HTML-escaped by `build.py`, and only `https://` links are used.
  Section text is rich text: `build.py` keeps only simple formatting (bold, italics, lists, links).

## Hot lunch ordering

- The Hot Lunch section of The Kings Corner comes from `content/lunch.json` via `lunch.py`: menu
  table, prices, deadline (9:00 AM the Friday before the week, unless the office types another) and (when `LUNCH_API` in `build.py` is set) the online order form
  (`site/lunch.js`). `site/lunch/menu.json` is the published menu the order service checks against.
- The office lists **several weeks** (`weeks` in `content/lunch.json`): each shows as its own card (banners alternate
  purple/gold and are tagged This week / Next week / Coming up; weeks that are over drop off), with an optional menu
  graphic (uploaded to `site/files/`) and/or a typed menu. Only a typed menu can be ordered online; parents pick the
  week on the form. `menu.json` lists every week under `weeks` (its top level is the first open week), and the
  service checks each order against the week picked.
- Orders go to the Cloudflare Worker in `worker/` (setup and upkeep in `worker/README.md`), which
  writes them to the office's Excel workbook and emails the parent a confirmation.
- **Staff lunches** ($4, `price_staff`): staff order on the same form by choosing "Staff" under Grade (shown to
  everyone since 2026-10-19; the old `?staff` link still works). **Staff families** pay it for every child: the
  Worker checks the workbook's Staff Families tab (last name or parent email; `worker/src/staff.js`) and their rows
  say "Staff family price" in Notes. Staff prices are never shown on the site (no total on the form or thank-you); the confirmation email has them. The row's
  Grade says Staff (counted with K5-12th on the Kitchen tab).
- With `LUNCH_API` empty, the page shows the menu and the backup Microsoft Forms link, if any.

## Go-live checklist (only with Garrett Greene's approval)

Done 2026-10-16 (steps 1-8). Still open: step 9 once GitHub has issued the certificate, and step 10.
The ibs account's sign-in name is now ibs@ibaclinton.com too.

DNS moves from the old Cloudflare account (no login) to GoDaddy, and Microsoft 365 gets the
domain at the same time, so lunch confirmations come from ibs@ibaclinton.com. The old WordPress
site stops being reachable at ibaclinton.com once the nameservers switch.

1. Merge the go-live commit (`DOMAIN = "ibaclinton.com"` in `build.py`, `MENU_URL` in
   `worker/wrangler.toml`), publish, and re-run "Deploy hot lunch order service".
2. GitHub repo Settings → Pages → Custom domain: `ibaclinton.com` (it warns until DNS is ready).
3. GoDaddy → ibaclinton.com → DNS → Nameservers → Change → **GoDaddy default**.
4. GoDaddy DNS records: delete any "Parked" A record for `@`, then add
   A `@` → 185.199.108.153, 185.199.109.153, 185.199.110.153, 185.199.111.153;
   CNAME `www` → `guarantee418.github.io` (edit the existing `www` record if there is one).
5. Microsoft 365 admin center → Settings → Domains → Add domain → `ibaclinton.com`. Let it
   sign in to GoDaddy and add the records automatically (verification TXT, MX, SPF TXT,
   autodiscover CNAME). Exchange records only; skip anything for Teams/Intune if asked.
6. Exchange admin center → Recipients → Mailboxes → ibs → Email addresses → add
   `ibs@ibaclinton.com` and **Set as primary**. (Sign-in name stays ibs@ibaministries.onmicrosoft.com.)
7. Microsoft Defender portal → Email & collaboration → Policies → Email authentication settings →
   DKIM → ibaclinton.com: add the two CNAME records it shows in GoDaddy, then **Enable**.
   Optional: TXT `_dmarc` → `v=DMARC1; p=none; rua=mailto:ibs@ibaclinton.com`.
8. Test: a test week and test order; the confirmation should arrive from ibs@ibaclinton.com.
9. When GitHub Pages shows the domain check passed, tick **Enforce HTTPS**.
10. Google Search Console: add ibaclinton.com and submit `https://ibaclinton.com/sitemap.xml`.
11. Keep the Flywheel/WordPress account until nothing is missed (it is no longer reachable at the
    domain, but its content stays there).

- This repo is public. Only website content belongs here: no credentials, exports, database dumps,
  student names or private notes.
