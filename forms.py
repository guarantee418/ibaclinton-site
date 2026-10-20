"""Forms the office adds to The Kings Corner sections in the online editor (.pages.yml, "Form" on a section).

A section's form has a name, optional text above it, a button text, a "stop taking responses" switch and up to
MAX_QUESTIONS questions (each a label, a type, choices for the pick types, required, help text). Every form also
asks the person's name and email, so the confirmation email can go out.

build.py calls form_html() for the section's pop-up and forms_file() for site/forms/forms.json, which the form
service (worker/src/forms.js, the same service as hot lunch) reads to check every response: the questions always
come from that file, never from the visitor's browser. The first response makes the form's own tab in the office
workbook (named after the form); a question added later adds a column at the end, and renaming, removing or
reordering questions starts a new tab ("Name 2"), so answers never land under the wrong heading.
"""
import hashlib
import html
import json
import re

MAX_QUESTIONS = 15
MAX_CHOICES = 20
# The editor's question types (.pages.yml) and what they become.
TYPES = {
    "Short answer": "short", "Long answer": "long", "Email": "email", "Phone": "phone", "Date": "date",
    "Number": "number", "Dropdown (pick one)": "choice", "Checkboxes (pick any)": "multi", "Yes/No box": "check",
}
# The first columns of every form's table (worker/src/forms.js FIXED); a question can't take one of these names.
FIXED = ["Received", "Response #", "Status", "Notes", "Name", "Email"]
# The workbook's own tabs (worker/src/guide.js ORDER): a form named like one gets " form" added to its tab name.
TABS = ["How to use", "Orders", "Kitchen", "Homerooms", "Gradelink", "Teachers", "Staff Families", "Conferences", "Tours",
        "Transcripts", "Inquiries", "Feedback", "Absences"]

def esc(value):
    return html.escape(str(value or "").strip())

def _slug(value):
    return re.sub(r"[^a-z0-9]+", "-", str(value or "").lower()).strip("-")

def _clean(value, max_len):
    return re.sub(r"\s+", " ", str(value or "")).strip()[:max_len]

def _choices(value):
    value = value if isinstance(value, list) else str(value or "").split("\n")
    out = []
    for v in value:
        v = _clean(v, 80)
        if v and v.lower() not in {c.lower() for c in out}:
            out.append(v)
    return out[:MAX_CHOICES]

def sheet_name(name):
    """The form's tab name: Excel allows 31 characters and no []:*?/\\ (28 here, room for " 2")."""
    s = re.sub(r"[\[\]:*?/\\]", " ", name).strip(" '")
    s = re.sub(r"\s+", " ", s)[:28].strip()
    return f"{s[:23].strip()} form" if s.lower() in {t.lower() for t in TABS} else s or "Form"

def table_name(form_id):
    """The table on that tab: letters and digits only (Excel's rule), e.g. Form_FieldTripPermission."""
    return "Form_" + ("".join(p.capitalize() for p in form_id.split("-"))[:40] or "Responses")

def spec(section, used=None):
    """One section's form as plain data, or None (no form, no name or no questions)."""
    f = section.get("form") if isinstance(section, dict) else None
    if not isinstance(f, dict):
        return None
    name = _clean(f.get("name"), 60)
    questions, labels = [], {c.lower() for c in FIXED}
    for q in f.get("questions") or []:
        if not isinstance(q, dict) or len(questions) == MAX_QUESTIONS:
            continue
        label, kind = _clean(q.get("label"), 80), TYPES.get(str(q.get("type") or "").strip(), "short")
        if not label:
            continue
        choices = _choices(q.get("choices")) if kind in ("choice", "multi") else []
        if kind in ("choice", "multi") and not choices:
            print(f"  form {name!r}: {label!r} has no choices, so it is a short answer")
            kind = "short"
        # Each question is a column heading, so no two may match (Excel would rename one).
        heading, n = label, 2
        while heading.lower() in labels:
            heading, n = f"{label} ({n})", n + 1
        labels.add(heading.lower())
        questions.append({"key": f"q{len(questions) + 1}", "label": heading, "text": label, "type": kind,
                          "required": bool(q.get("required")), "choices": choices,
                          "help": _clean(q.get("help"), 200)})
    if not name or not questions:
        if name or questions:
            print(f"  form {name!r} left out: it needs a name and at least one question")
        return None
    form_id, n = _slug(name) or "form", 2
    while used is not None and form_id in used:
        form_id, n = f"{_slug(name) or 'form'}-{n}", n + 1
    if used is not None:
        used.add(form_id)
    # The page and forms.json are built together; the service refuses a response from a page built with
    # other questions (an old copy still open in a browser), so answers can't shift to the wrong question.
    shape = json.dumps([[q["label"], q["type"], q["required"], q["choices"]] for q in questions])
    return {"id": form_id, "name": name, "sheet": sheet_name(name), "table": table_name(form_id),
            "version": hashlib.sha256(shape.encode()).hexdigest()[:10], "closed": bool(f.get("closed")),
            "intro": _clean(f.get("intro"), 600), "button": _clean(f.get("button"), 40) or "Send",
            "questions": questions}

def specs(data):
    used = set()
    return [s for s in (spec(sec, used) for sec in data.get("sections") or [] if str(sec.get("title") or "").strip()) if s]

def forms_file(data):
    """site/forms/forms.json: every form on The Kings Corner, for the form service."""
    forms = [{k: v for k, v in s.items() if k not in ("intro", "button")} for s in specs(data)]
    return json.dumps({"forms": forms}, indent=2) + "\n"

def _question(q, form_id):
    key, label = q["key"], esc(q.get("text") or q["label"])  # the page shows the question as typed
    named = f'name="{key}" data-label="{label}"'
    opt = "" if q["required"] else ' <span class="opt">(optional)</span>'
    req = " required" if q["required"] else ""
    help_ = f'<span class="cf-help">{esc(q["help"])}</span>' if q["help"] else ""
    if q["type"] == "multi":
        boxes = "".join(f'<label class="cf-check"><input type="checkbox" {named} value="{esc(c)}"> {esc(c)}</label>'
                        for c in q["choices"])
        return (f'<div class="full cf-group" role="group" aria-labelledby="{form_id}-{key}"{" data-required" if req else ""}>'
                f'<span class="cf-label" id="{form_id}-{key}">{label}{opt}</span>{help_}{boxes}</div>')
    if q["type"] == "check":
        return (f'<div class="full cf-group"><label class="cf-check"><input type="checkbox" {named} value="Yes"{req}> '
                f'{label}{opt}</label>{help_}</div>')
    if q["type"] == "choice":
        options = "".join(f"<option>{esc(c)}</option>" for c in q["choices"])
        field = f'<select {named}{req}><option value="">Choose…</option>{options}</select>'
    elif q["type"] == "long":
        field = f'<textarea {named} maxlength="4000"{req}></textarea>'
    else:
        kind = {"email": "email", "phone": "tel", "date": "date", "number": "number"}.get(q["type"], "text")
        extra = ' step="any"' if kind == "number" else ' maxlength="200"' if kind in ("text", "email", "tel") else ""
        field = f'<input type="{kind}" {named}{extra}{req}>'
    wide = q["type"] in ("long", "choice") or len(q["label"]) > 40 or q["help"]
    return f'<label{" class=\"full\"" if wide else ""}>{label}{opt}{help_}{field}</label>'

def form_html(section, root, api, used=None):
    """The form in a section's pop-up (site/forms.js sends it), or "" when the section has no form."""
    s = spec(section, used)
    if not s:
        return ""
    intro = f'<p class="cf-intro">{esc(s["intro"])}</p>' if s["intro"] else ""
    if s["closed"] or not api:
        return intro + '<p class="notice">This form isn’t taking responses right now. Questions? Call the office at (301) 856-1616.</p>'
    fid = esc(s["id"])
    return (
        f'<div class="cf-status" id="cf-status-{fid}" role="status" aria-live="polite"></div>'
        f'<form class="info cf-form" id="cf-{fid}" data-api="{esc(api)}" data-form="{fid}" data-version="{s["version"]}" novalidate hidden>'
        f'{intro}<fieldset><legend>{esc(s["name"])}</legend>'
        '<label>Your name<input name="name" autocomplete="name" maxlength="120" required></label>'
        '<label>Your email<input type="email" name="email" autocomplete="email" maxlength="120" required></label>'
        + "".join(_question(q, fid) for q in s["questions"]) +
        '</fieldset>'
        '<label class="hp" aria-hidden="true">Leave this field blank<input name="website" tabindex="-1" autocomplete="off"></label>'
        f'<button class="btn" type="submit">{esc(s["button"])}</button>'
        '<p class="cf-fine">You’ll get a confirmation email with a copy of your answers.</p>'
        '</form>'
        '<noscript><p class="notice">To fill out this form, please turn on JavaScript, or email '
        '<a href="mailto:iba@ibcministries.org">iba@ibcministries.org</a> or call (301) 856-1616.</p></noscript>'
        f'<script src="{root}forms.js" defer></script>')
