# ibaclinton-site

Website files for Independent Baptist Academy, Clinton, MD (ibaclinton.com).

- `src/` is where pages are edited: `src/layout.html` (header, menu, footer shared by every page),
  `src/cta.html` (the "Book a Tour / Request Info" band) and one file per page in `src/pages/`.
- `python3 build.py` combines them into the finished site in `site/`. Commit both `src/` and `site/`.
- `site/img/` holds photos (resized, location data stripped); `site/files/` holds PDFs.
- `site/styles.css` is the one stylesheet.
- Publishing is manual: Actions tab → "Publish to GitHub Pages" → Run workflow.
- Request Info form: set `FORMSPREE_ID` in `build.py` once the Formspree form exists, then rebuild.
- When ibaclinton.com is pointed here, change `BASE_PATH` in `build.py` to `/` and rebuild.
- This repo is public. Only website content belongs here: no credentials, exports, database dumps,
  student names or private notes.
