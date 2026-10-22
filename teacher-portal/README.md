# Teacher Portal content (private repo template)

This folder is the starting point for the **private** repo `guarantee418/ibaclinton-teachers`. It is kept here only
as a template: nothing in it is private (no password, no meeting notes). Once that repo exists, the real content
lives there and this folder is not used.

How the portal works:

- ibaclinton.com/teachers/ (hidden: not linked anywhere, not in the sitemap, search engines asked not to list it)
  shows only a password box (`src/pages/teachers.html`, `site/teachers.js` in the site repo).
- The office edits the portal in the online editor (Pages CMS, the same as The Kings Corner) by choosing this
  private repo: settings (the password), banner, announcements, tiles, sections, links and files, meeting notes
  and forms. Files go in `files/`.
- Each save runs "Build the Teacher Portal" (`.github/workflows/build.yml`), which runs `teachers.py` from the
  site repo and commits `built/portal.json`.
- The order service (`worker/src/teachers.js` in the site repo) reads `built/portal.json` and `files/` with the
  read-only `TEACHER_REPO_TOKEN`, checks the password, and only then sends the portal and its files to the browser.
  Form responses go to their own tabs in IBA Portal Forms.xlsx (like The Kings Corner forms), with a confirmation to
  the teacher and the details to `TEACHER_FORM_NOTIFY`.

## Setting it up (Garrett)

1. Create the private repo `guarantee418/ibaclinton-teachers` and copy this folder's contents into it
   (`.pages.yml`, `.github/`, `content/`, `files/`).
2. Make a fine-grained GitHub token: only that repo, permission **Contents: Read-only**. Add it to the site repo
   as the Actions secret `TEACHER_REPO_TOKEN`.
3. Run Actions -> "Deploy hot lunch order service" on `main` (after the Teacher Portal change is merged).
4. In Pages CMS, open the new repo, type a password under Teacher Portal -> Settings and save.
5. Invite Mrs. Mora and Mrs. Napierala to the private repo as collaborators (like the site repo).
