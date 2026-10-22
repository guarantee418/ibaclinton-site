#!/usr/bin/env python3
"""Build the Teacher Portal: a hidden, password-protected page for IBA staff that looks and works like
The Kings Corner (tiles that open sections in pop-ups, announcement slides, links and files, forms).

Its content is NOT in this public repo. It lives in the private repo guarantee418/ibaclinton-teachers
(content/teachers.json and files/, edited by the office in the same online editor as The Kings Corner). That
repo's workflow checks out this repo and runs

    python3 teachers.py content/teachers.json built/portal.json

and commits built/portal.json. The order service (worker/src/teachers.js) reads it from the private repo with a
read-only token and hands the page to a browser only after the portal password is given; files are passed
through the service the same way. ibaclinton.com/teachers/ is only the sign-in box (src/pages/teachers.html,
site/teachers.js), so nothing private is ever published on GitHub Pages.

Same as build.py: everything typed in the editor is escaped, rich text keeps only simple formatting, links are
https only.
"""
import hashlib
import html
import json
import pathlib
import sys
import urllib.parse

import build
import forms

ROOT = "../"  # the page is ibaclinton.com/teachers/
NAME = "Teacher Portal"
FORM_ENDPOINT = "/teachers/form"
# Uploaded files are served by the service; site/teachers.js swaps this for the signed address.
FILES = "__TP_FILES__/"
FILE_TYPES = (".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".jpg", ".jpeg", ".png", ".gif", ".webp")
# Same list as worker/src/teachers.js FILE_TYPES.


def file_url(name):
    name = pathlib.PurePosixPath(str(name or "")).name
    if not name.lower().endswith(FILE_TYPES):
        return ""
    return FILES + html.escape(urllib.parse.quote(name), quote=True)


def _date_label(value):
    try:
        d = build.lunch.dt.date.fromisoformat(str(value or "")[:10])
    except ValueError:
        return ""
    return f"{d:%B} {d.day}, {d.year}"


def meetings_html(items):
    """Meeting notes, newest first: each a fold-out with the date, title, notes and an optional file."""
    rows = []
    for m in items or []:
        if not isinstance(m, dict):
            continue
        title, when = build.esc(m.get("title")), _date_label(m.get("date"))
        notes, attach = build.clean_html(m.get("notes")), file_url(m.get("file"))
        if not (title or when) or not (notes or attach):
            continue
        rows.append((str(m.get("date") or "")[:10], title, when, notes, attach, build.esc(m.get("file_title"))))
    if not rows:
        return ""
    rows.sort(key=lambda r: r[0], reverse=True)
    out = []
    for i, (_, title, when, notes, attach, file_title) in enumerate(rows):
        link = (f'<ul class="docs"><li><a href="{attach}">{file_title or "Open the notes"}</a></li></ul>'
                if attach else "")
        head = f'<span class="tp-date">{when}</span> {title}' if when else title
        out.append(f'<details class="tp-meeting"{" open" if i == 0 else ""}><summary>{head}</summary>'
                   f'<div class="tp-notes">{notes}{link}</div></details>')
    return '<div class="tp-meetings">' + "".join(out) + "</div>"


def extras(section, root):
    return meetings_html(section.get("meetings"))


def password_hash(value):
    """The service compares a hash, so the built file never holds the password itself."""
    value = str(value or "").strip()
    return hashlib.sha256(("ibaclinton-teachers|" + value).encode()).hexdigest() if value else ""


def portal(data):
    for s in data.get("sections") or []:
        for item in s.get("links") or []:
            name = pathlib.PurePosixPath(str(item.get("file") or "")).name
            if name and not name.lower().endswith(FILE_TYPES):
                print(f"  {name!r} left out: only PDF, Word, Excel, PowerPoint and picture files")
    page = build.portal_block(data, ROOT, NAME, extras, file_url=file_url, form_endpoint=FORM_ENDPOINT)
    settings = data.get("settings") or {}
    return {
        "html": page,
        "forms": forms.specs(data, "teacher-", "the Teacher Portal"),
        "password": password_hash(settings.get("password")),
    }


if __name__ == "__main__":
    src, dest = pathlib.Path(sys.argv[1]), pathlib.Path(sys.argv[2])
    out = portal(json.loads(src.read_text()))
    if not out["password"]:
        print("  no portal password set: nobody can sign in until one is typed under Settings")
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(json.dumps(out, indent=1) + "\n")
    print("wrote", dest)
