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
- Request Info form: Formspree form `xaeqabaz` (account and inbox: ggreene@ibcministries.org; free plan,
  50 submissions a month). The form ID is public by nature, so it lives in `build.py`; no secret is needed.
- `build.py` also writes `sitemap.xml` and `robots.txt`. While `DOMAIN` is empty the pages carry a
  `noindex` tag so the github.io preview stays out of search results.

## Online editor for the school office (Pages CMS)

- Mrs. Mora and Mrs. Napierala edit the **Hot Lunch** menu (`content/lunch.json`) and the Parent Portal
  document lists (`content/documents.json`, PDFs uploaded to `site/files/`) at app.pagescms.org
  (`ibaclinton.com/admin/` links there). They are invited by email as collaborators, no GitHub account.
- `.pages.yml` defines what they can edit. Collaborators can't change it; changing it needs Garrett's OK.
- Their saves commit to `main`, and the Pages workflow publishes changes to `content/` or `site/files/`
  automatically (approved by Garrett 2026-10-16). Everything else is still published by hand.
- The workflow runs `build.py`, so the committed `site/` copy of those pages may lag behind; that's fine.
- Everything typed in the editor is HTML-escaped by `build.py`, and only `https://` links are used.

## Go-live checklist (only with Garrett's approval)

1. Request Info form tested (one test submission reached ggreene@).
2. In `build.py` set `DOMAIN = "ibaclinton.com"`, run `python3 build.py`, commit and push. This adds
   `site/CNAME`, the Google Analytics tag (`GA_ID`, the same tag WordPress uses), switches the
   privacy policy's cookie paragraph and removes `noindex`.
3. Run "Publish to GitHub Pages".
4. Repo Settings → Pages → Custom domain: `ibaclinton.com`; tick "Enforce HTTPS" once offered.
5. DNS (Cloudflare): point the `ibaclinton.com` A records at GitHub Pages
   (185.199.108.153, 185.199.109.153, 185.199.110.153, 185.199.111.153) and change `www` (now A records) to a CNAME to
   `guarantee418.github.io`. **Leave the Zoho MX records and other TXT records alone**, or
   @ibaclinton.com email stops. Set the records to "DNS only" (grey cloud) until GitHub issues
   the HTTPS certificate.
6. In Google Search Console, submit `https://ibaclinton.com/sitemap.xml`.
7. Keep the WordPress site on Flywheel for a few weeks as a fallback.
- This repo is public. Only website content belongs here: no credentials, exports, database dumps,
  student names or private notes.
