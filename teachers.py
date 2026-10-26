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
NAME = "The King's Court"  # what teachers see; "Teacher Portal" in code and notes
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


# Who a form's emails may go to: school addresses only, so a typo or a pasted outside address can't send staff
# answers outside the school.
SCHOOL_DOMAINS = ("ibcministries.org", "ibaministries.org", "ibaclinton.com")
APPROVAL_REPLY = "iba@ibcministries.org"
MAX_NOTIFY, MAX_APPROVALS = 10, 5


def school_email(value, where):
    v = str(value or "").strip().lower()
    if not v:
        return ""
    if not build.re.fullmatch(r"[^@\s<>\"]+@[^@\s<>\"]+", v) or v.rsplit("@", 1)[1] not in SCHOOL_DOMAINS:
        print(f"  {where}: {v!r} left out: only {', '.join(SCHOOL_DOMAINS)} addresses")
        return ""
    return v


def routing(section, spec):
    """A Teacher Portal form's own emails and copies, from the editor (each form can differ):
    notify     who gets each response (else TEACHER_FORM_NOTIFY on the service)
    approvals  people asked to approve each response by replying to approval_reply (iba@), each optionally only
               when a question has a certain answer (e.g. a vehicle only when "School vehicle needed?" is Yes);
               when that answer isn't given they get the plain notice instead
    monday     also add each response to the form's own monday.com board
    one_email  no separate approval emails: one email to the list and the needed approvers, who answer by Reply All"""
    f, name = section.get("form") or {}, spec["name"]
    notify = []
    for v in f.get("notify") or []:
        v = school_email(v, f"form {name!r}")
        if v and v not in notify:
            notify.append(v)
    by_text = {q["text"].lower(): q for q in spec["questions"]} | {q["label"].lower(): q for q in spec["questions"]}
    approvals = []
    for a in f.get("approvals") or []:
        if not isinstance(a, dict) or len(approvals) == MAX_APPROVALS:
            continue
        email, title = school_email(a.get("email"), f"form {name!r} approval"), forms._clean(a.get("title"), 40)
        if not email:
            continue
        rule = {}
        question = forms._clean(a.get("only_if_question"), 80)
        if question:
            q = by_text.get(question.lower())
            if not q:
                print(f"  form {name!r}: approval {title!r} asks about {question!r}, which isn't a question; asked every time")
            else:
                rule = {"question": q["key"], "answer": forms._clean(a.get("only_if_answer"), 80) or "Yes"}
        approvals.append({"title": title or "Approval", "name": forms._clean(a.get("name"), 60), "email": email, **rule})
    return {"notify": notify[:MAX_NOTIFY], "approvals": approvals,
            "approval_reply": school_email(f.get("approval_reply"), f"form {name!r}") or APPROVAL_REPLY,
            "monday": bool(f.get("monday")), "one_email": bool(f.get("one_email"))}


def form_specs(data):
    """The forms with their routing; same order and ids as the page (build.portal_block)."""
    used, out = set(), []
    for sec in data.get("sections") or []:
        if not str(sec.get("title") or "").strip():
            continue
        s = forms.spec(sec, used, "teacher-", "The King's Court")
        if s:
            out.append({**s, **routing(sec, s)})
    return out


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
        "forms": form_specs(data),
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
