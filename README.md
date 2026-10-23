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

## Forms the office adds (The Kings Corner)

- Any Kings Corner section can have a **Form** in the editor: a name, optional text above it, button text, "Stop
  taking responses", and up to 15 questions (Short answer, Long answer, Email, Phone, Date, Number, Dropdown,
  Checkboxes, Yes/No box; required; help text). Every form also asks the person's name and email. Approved by
  Garrett 2026-10-20: forms go live when the office saves them.
- `forms.py` draws the form (sent by `site/forms.js`) and writes `site/forms/forms.json`, which the Pages workflow
  commits; the service (`worker/src/forms.js`, `POST /form`) reads it from GitHub (`FORMS_URL`) and checks every
  answer against it. A page built with other questions is refused ("This form changed; please reload").
- The first response makes the form's own tab in IBA Portal Forms.xlsx (named after the form, table `Form_<Name>`).
  A question added at the end adds a column; renaming the form, or renaming, removing or reordering questions,
  starts a new tab ("Name 2"). The person gets a confirmation with their answers; `FORM_NOTIFY` (iba@ and ibs@)
  gets the details, Reply-To the person.
- TEST responses: Actions -> "Delete a TEST booking, lunch order or tour sign-up", kind "form response" (only a Name
  starting with TEST; a form tab named "TEST..." is removed once empty).

## Teacher Portal (ibaclinton.com/teachers/), shown to teachers as "The King's Court"

- A hidden, password-protected page for IBA staff that looks and works like The Kings Corner (banner and announcement
  slides, tiles that open pop-ups, links and files, meeting notes, forms). Not linked from any page, not in the
  sitemap, `noindex` (`hidden: yes` page key in `src/pages/teachers.html`).
- This public repo holds only the code and the sign-in box. The content (meeting notes, files, forms, the password)
  lives in a **private** repo, edited by the office in the same online editor (Pages CMS). `teacher-portal/` here is
  only the starter template for that repo.
- The private repo's workflow runs `teachers.py` (from this repo's `main`) to build `built/portal.json`. The order
  service (`worker/src/teachers.js`) reads it with a read-only token (`TEACHER_REPO_TOKEN` secret), checks the
  password and only then sends the page, its files (links that last 6 hours) and takes its forms
  (`POST /teachers/form`). `site/teachers.js` signs in and shows the portal.
- Teacher Portal forms use the same engine as The Kings Corner forms (a tab each in the office workbook), plus their
  own routing set in the editor: who gets each response (school addresses only), approvers who are emailed to reply
  with their approval, and an optional copy to the form's own monday.com board (`worker/src/monday.js`).
- TEST responses: Actions -> "Delete a TEST booking, lunch order or tour sign-up", kind "form response".

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
- **No school days** (`no_school` on a week, "No school on" in the editor): those days drop off the menu and the
  form (so they can't be ordered), the card says "No school Friday, Oct 30.", and the week's dates shorten
  (October 26 – 29). A Friday marked off moves the next week's usual deadline to Thursday 9:00 AM.
- With `LUNCH_API` empty, the page shows the menu and the backup Microsoft Forms link, if any.

## This repo is public

Only website content belongs here: **no credentials, exports, database dumps, student names or private
notes.** Operational runbooks — the domain/DNS/Microsoft 365 go-live steps, the Cloudflare Worker's
account setup and secrets, and the accounts inventory — are kept in the school's **private** operations
repo (`ibaclinton`, under `docs/ibaclinton-site/`), not here.
