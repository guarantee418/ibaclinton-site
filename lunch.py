"""Hot lunch: the menu table, the online order form and the public menu file.

The office edits content/lunch.json in the online editor (.pages.yml). build.py calls
lunch_block() for The King's Court and menu_file() for site/lunch/menu.json, which the
order service (worker/) reads to check every order: prices, open days and the deadline
always come from that file, never from the parent's browser.
"""
import datetime as dt
import html
import json
import pathlib
from zoneinfo import ZoneInfo

EASTERN = ZoneInfo("America/New_York")
USUAL_DEADLINE = dt.time(9, 0)  # orders close Friday at 9:00 AM the week before
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
# Grades a parent can pick. Preschool grades pay the preschool price.
GRADES = [("K3", "Preschool (K3)", True), ("K4", "Preschool (K4)", True), ("K5", "Kindergarten (K5)", False)] + [
    (str(n), f"{n}{'st' if n == 1 else 'nd' if n == 2 else 'rd' if n == 3 else 'th'} Grade", False)
    for n in range(1, 13)]
# Staff order on the same form, at the staff price: "Staff" is the last choice under Grade (Garrett, 2026-10-19).
STAFF = {"value": "STAFF", "label": "Staff", "preschool": False, "staff": True}

def esc(value):
    return html.escape(str(value or "").strip())

def money(value):
    return f"${float(value):,.2f}"

def _date(value):
    try:
        return dt.date.fromisoformat(str(value or "").strip()[:10])
    except ValueError:
        return None

def _deadline(value, start=None):
    """The office's deadline if they typed one, else the usual: 9:00 AM on the Friday before the week."""
    try:
        return dt.datetime.fromisoformat(str(value or "").strip()).replace(tzinfo=EASTERN)
    except ValueError:
        if not start:
            return None
        return dt.datetime.combine(start - dt.timedelta(days=(start.weekday() - 4) % 7 or 7),
                                   USUAL_DEADLINE, tzinfo=EASTERN)

def _range_label(start, end):
    if start.month == end.month:
        return f"{start:%B} {start.day} – {end.day}"
    return f"{start:%B} {start.day} – {end:%B} {end.day}"

def _monday(day):
    """The Monday of the school week a date belongs to: Saturday and Sunday count as the coming week, so a
    week start typed as Sunday 10/18 means Monday 10/19."""
    if not day:
        return None
    return day + dt.timedelta(days=7 - day.weekday()) if day.weekday() >= 5 else day - dt.timedelta(days=day.weekday())

IMAGE_TYPES = (".jpg", ".jpeg", ".png", ".gif", ".webp")
MAX_OPTIONS = 5  # the Kitchen tab has room for this many choices per day

def _options(value):
    """A day's choices (e.g. Cheese pizza, Pepperoni pizza): each listed once, at most MAX_OPTIONS.
    One choice alone is no choice, so it counts as none."""
    seen, out = set(), []
    for v in value if isinstance(value, list) else []:
        v = " ".join(str(v or "").split())[:60]
        if v and v.upper() != "X" and v.lower() not in seen:
            seen.add(v.lower())
            out.append(v)
    return out[:MAX_OPTIONS] if len(out) > 1 else []

def _no_school(w):
    """The weekdays the office marked "No school" for a week (a list from the editor, or one day as text)."""
    value = w.get("no_school")
    value = value if isinstance(value, list) else [value]
    return [d for d in DAYS if d in {str(v or "").strip() for v in value}]

def _days_off(listed):
    """Every no-school date across the listed weeks, so a week's usual deadline can skip a Friday off."""
    out = set()
    for w in listed:
        start = _monday(_date(w.get("week_start"))) if isinstance(w, dict) else None
        if start:
            out.update(start + dt.timedelta(days=DAYS.index(d)) for d in _no_school(w))
    return out

def _day_off(day, start):
    date = start + dt.timedelta(days=DAYS.index(day))
    return {"day": day, "date": date.isoformat(), "date_label": f"{date:%b} {date.day}"}

def _week(w, days_off=frozenset()):
    """One lunch week as plain data: its Monday, label, order deadline, note, menu graphic, days and days off.
    A day marked No school is left off the menu (so it can't be ordered) and out of the week's dates."""
    start = _monday(_date(w.get("week_start")))
    off = _no_school(w)
    days = []
    for item in w.get("menu") or []:
        day, meal = str(item.get("day") or "").strip(), str(item.get("meal") or "").strip()
        if day in off:
            continue
        options = _options(item.get("options"))
        if not meal and options:
            meal = " or ".join(options)
        if not meal or day not in DAYS:
            continue
        date = start + dt.timedelta(days=DAYS.index(day)) if start else None
        days.append({"key": day[:3].lower(), "day": day, "meal": meal, "options": options,
                     "date": date.isoformat() if date else "",
                     "date_label": f"{date:%b} {date.day}" if date else ""})
    deadline = _deadline(w.get("order_deadline"), start) if days else None
    if deadline and not str(w.get("order_deadline") or "").strip():
        # The usual Friday deadline moves to the school day before when that Friday is a day off.
        while deadline.date() in days_off and deadline.weekday() > 0:
            deadline -= dt.timedelta(days=1)
    week_label = str(w.get("week_of") or "").strip()
    school = [i for i in range(5) if DAYS[i] not in off] or [0, 4]
    if start and not week_label:
        week_label = _range_label(start + dt.timedelta(days=school[0]), start + dt.timedelta(days=school[-1]))
    graphic = pathlib.PurePosixPath(str(w.get("graphic") or "").strip()).name
    return {
        "week_start": start.isoformat() if start else "",
        "week_end": (start + dt.timedelta(days=4)).isoformat() if start else "",
        "week_label": week_label,
        "deadline": deadline.astimezone(dt.timezone.utc).isoformat() if deadline else "",
        "deadline_label": (f"{deadline:%A, %B} {deadline.day} at {deadline:%I:%M %p}".replace(" 0", " ")
                           if deadline else ""),
        "note": str(w.get("note") or "").strip(),
        "graphic": f"files/{graphic}" if graphic.lower().endswith(IMAGE_TYPES) else "",
        "days": days,
        "no_school": [_day_off(d, start) for d in off] if start else [],
    }

def _weeks(data, today=None):
    """The weeks the office listed (content/lunch.json "weeks"; an older file with one week at the top level
    still works), earliest first, without weeks that are over or have nothing to show."""
    today = today or dt.datetime.now(EASTERN).date()
    listed = data.get("weeks") if isinstance(data.get("weeks"), list) else [data]
    off = _days_off(listed)
    weeks = [_week(w, off) for w in listed if isinstance(w, dict)]
    weeks = [w for w in weeks if w["week_start"] and (w["days"] or w["graphic"]) and w["week_end"] >= today.isoformat()]
    return sorted(weeks, key=lambda w: w["week_start"])

def menu_data(data, today=None):
    """The menu as plain data, shared by the page, menu.json and the order service. The top level is the
    first week parents can still order (so an older reader of menu.json keeps working); "weeks" has them all."""
    weeks = _weeks(data, today)
    now = dt.datetime.now(dt.timezone.utc).isoformat()
    first = next((w for w in weeks if w["days"] and w["deadline"] > now), None) or next(
        (w for w in weeks if w["days"]), None) or {"week_start": "", "week_end": "", "week_label": "", "deadline": "",
                                                  "deadline_label": "", "note": "", "graphic": "", "days": [],
                                                  "no_school": []}
    return {
        **first,
        "prices": {"regular": float(data.get("price_regular") or 8), "preschool": float(data.get("price_preschool") or 5),
                   "staff": float(data.get("price_staff") or 4)},
        "grades": [{"value": v, "label": l, "preschool": p} for v, l, p in GRADES] + [STAFF],
        "weeks": weeks,
    }

def menu_file(data):
    return json.dumps(menu_data(data), indent=2) + "\n"

def can_order_online(week):
    """Online ordering needs a week start date, a deadline and at least one meal."""
    return bool(week["week_start"] and week["deadline"] and week["days"])

def _no_school_line(week):
    """"No school Friday, Oct 30." for the days the office marked off."""
    if not week["no_school"]:
        return ""
    names = [f'{d["day"]}, {d["date_label"]}' for d in week["no_school"]]
    return f'<p class="lunch-off">No school {esc(" and ".join(names) if len(names) < 3 else "; ".join(names))}.</p>'

def _menu_table(week):
    return '<table class="lunch"><tbody>' + "".join(
        f'<tr><th scope="row">{esc(d["day"])}'
        + (f' <span class="lunch-date">{esc(d["date_label"])}</span>' if d["date_label"] else "")
        + f'</th><td>{esc(d["meal"])}'
        + (f'<span class="lunch-choices">Choice of: {esc(", ".join(d["options"]))}</span>'
           if d["options"] and d["meal"] != " or ".join(d["options"]) else "")
        + '</td></tr>' for d in week["days"]) + "</tbody></table>"

def _week_card(week, i, root):
    """One week, with a banner that says which week it is (lunch.js marks This week / Next week)."""
    out = [f'<section class="lunch-wk lunch-wk-{"a" if i % 2 == 0 else "b"}" data-start="{week["week_start"]}" '
           f'data-end="{week["week_end"]}" aria-label="Week of {esc(week["week_label"])}">'
           f'<h3 class="lunch-wk-head"><span class="lunch-wk-tag" hidden></span>'
           f'<span class="lunch-wk-dates">Week of {esc(week["week_label"])}</span></h3><div class="lunch-wk-body">']
    if week["graphic"]:
        src = f'{root}{esc(week["graphic"])}'
        out.append(f'<a class="lunch-graphic" href="{src}" target="_blank" rel="noopener">'
                   f'<img src="{src}" alt="Hot lunch menu for the week of {esc(week["week_label"])}" loading="lazy"></a>')
    if week["days"]:
        out.append(_menu_table(week))
    out.append(_no_school_line(week))
    if week["note"]:
        out.append(f'<p>{esc(week["note"])}</p>')
    out.append("</div></section>")
    return "".join(out)

def lunch_block(data, root, api, today=None):
    menu = menu_data(data, today)
    weeks = menu["weeks"]
    price_line = (f"<p>Lunches are {money(menu['prices']['regular'])} "
                  f"({money(menu['prices']['preschool'])} for preschool).</p>")
    contact = ('<p>If your student is enrolled and you are not receiving the lunch form, please contact the '
               'school office at <a href="tel:3018561616">(301) 856-1616</a> or '
               '<a href="mailto:iba@ibcministries.org">iba@ibcministries.org</a>.</p>')
    if not weeks:
        return ('<p>A lunch menu form is sent out via email/text one week in advance for students to select '
                'the days they would like to eat hot lunch. Orders are due by 9:00 AM on the Friday before.</p>' + price_line + contact +
                f'<img src="{root}img/hot-lunch-menu.jpg" alt="Sample weekly hot lunch menu with a different meal '
                'for each day, such as chicken parmesan, soft tacos, hot dogs, chili and burgers" loading="lazy" '
                'width="1000" height="562" style="border-radius:4px;margin-top:.8rem">')
    out = [price_line, '<div class="lunch-weeks">', "".join(_week_card(w, i, root) for i, w in enumerate(weeks)), "</div>"]
    if menu["note"] and not any(w["note"] for w in weeks):
        out.append(f"<p>{esc(menu['note'])}</p>")
    orderable = [w for w in weeks if can_order_online(w)]
    if api and orderable:
        out.append(order_form(menu, orderable, root, api))
    else:
        out.append(_forms_link(str(data.get("order_form_url") or "").strip()))
    return "".join(out) + contact

def _forms_link(form):
    if not form.startswith("https://"):
        return ""
    if form.split("/")[2].lower() in ("forms.office.com", "forms.cloud.microsoft", "forms.microsoft.com"):
        # Microsoft Forms can be shown right on the page; the button stays for full-screen use.
        sep = "&" if "?" in form else "?"
        return (f'<iframe class="lunch-form" src="{esc(form)}{sep}embed=true" title="Hot lunch order form" '
                'loading="lazy" allowfullscreen></iframe>'
                f'<p><a class="btn outline" href="{esc(form)}" target="_blank" rel="noopener">'
                "Open the order form in a new window</a></p>")
    return f'<p><a class="btn" href="{esc(form)}">Order Hot Lunch</a></p>'

def _day_box(d):
    """One day on the order form: a checkbox, plus a list of the day's choices when it has any."""
    box = (f'<label class="lunch-day"><input type="checkbox" name="days" value="{d["key"]}">'
           f'<span><b>{esc(d["day"])}</b> {esc(d["date_label"])}<br>{esc(d["meal"])}</span></label>')
    if not d["options"]:
        return box
    opts = '<option value="">Choose…</option>' + "".join(f"<option>{esc(o)}</option>" for o in d["options"])
    return (f'<div class="lunch-pick">{box}<label class="lunch-choice">Choice for {esc(d["day"])}'
            f'<select name="choice" data-day="{d["key"]}">{opts}</select></label></div>')

def order_form(menu, weeks, root, api):
    grade_options = '<option value="">Choose…</option>' + "".join(
        f'<option value="{esc(g["value"])}">{esc(g["label"])}</option>' for g in menu["grades"])
    # One set of day boxes per open week; lunch.js shows the week picked under "Week".
    day_sets = "".join(
        f'<div class="lunch-days" role="group" aria-label="Lunch days, week of {esc(w["week_label"])}" '
        f'data-week="{w["week_start"]}"{"" if i == 0 else " hidden"}>{"".join(_day_box(d) for d in w["days"])}'
        f'{_no_school_line(w)}</div>'
        for i, w in enumerate(weeks))
    week_options = "".join(
        f'<option value="{w["week_start"]}" data-deadline="{w["deadline"]}">Week of {esc(w["week_label"])} '
        f'(order by {esc(w["deadline_label"])})</option>' for w in weeks)
    child = (f'<fieldset class="lunch-child"><legend>Student</legend>'
             f'<label>Student’s full name<input name="student" autocomplete="off" maxlength="80" required></label>'
             f'<label>Grade<select name="grade" required>{grade_options}</select></label>'
             f'{day_sets}'
             f'<button type="button" class="lunch-remove" hidden>Remove this student</button></fieldset>')
    return (
        f'<form class="info lunch-order" id="lunch-order" data-api="{esc(api)}" '
        f'data-menu="{root}lunch/menu.json" novalidate>'
        '<h3>Order hot lunch online</h3>'
        '<div id="lunch-closed" class="notice" hidden>Online ordering is closed for now. '
        'To change or add an order, please call the office at <a href="tel:3018561616">(301) 856-1616</a>.</div>'
        f'<label class="lunch-week-pick">Week<select name="week">{week_options}</select></label>'
        '<fieldset><legend>Parent/Guardian</legend>'
        '<label>First name<input name="parent_first" autocomplete="given-name" maxlength="60" required></label>'
        '<label>Last name<input name="parent_last" autocomplete="family-name" maxlength="60" required></label>'
        '<label>Email<input type="email" name="email" autocomplete="email" maxlength="120" required></label>'
        '<label>Phone<input type="tel" name="phone" autocomplete="tel" maxlength="30" required></label>'
        '</fieldset>'
        f'<div id="lunch-children">{child}</div>'
        '<p><button type="button" class="btn outline" id="lunch-add">Add another student</button></p>'
        '<label class="hp" aria-hidden="true">Leave this field blank<input name="website" tabindex="-1" autocomplete="off"></label>'
        '<p class="lunch-total" aria-live="polite">Total: <b id="lunch-total">$0.00</b></p>'
        '<div id="lunch-status" role="status" aria-live="polite"></div>'
        '<p><button class="btn" type="submit">Place order</button></p>'
        '<p class="lunch-fine">You’ll get a confirmation email. Lunch charges are added to your Gradelink account. '
        'To change an order after placing it, please call the office.</p>'
        '</form>'
        '<noscript><p class="notice">Online ordering needs JavaScript. Please call the office at '
        '(301) 856-1616 to order.</p></noscript>'
        f'<script src="{root}lunch.js" defer></script>')
