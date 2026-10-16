"""Hot lunch: the menu table, the online order form and the public menu file.

The office edits content/lunch.json in the online editor (.pages.yml). build.py calls
lunch_block() for the Parent Portal and menu_file() for site/lunch/menu.json, which the
order service (worker/) reads to check every order: prices, open days and the deadline
always come from that file, never from the parent's browser.
"""
import datetime as dt
import html
import json
from zoneinfo import ZoneInfo

EASTERN = ZoneInfo("America/New_York")
USUAL_DEADLINE = dt.time(9, 0)  # orders close Friday at 9:00 AM the week before
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
# Grades a parent can pick. Preschool grades pay the preschool price.
GRADES = [("K3", "Preschool (K3)", True), ("K4", "Preschool (K4)", True), ("K5", "Kindergarten (K5)", False)] + [
    (str(n), f"{n}{'st' if n == 1 else 'nd' if n == 2 else 'rd' if n == 3 else 'th'} Grade", False)
    for n in range(1, 13)]

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

def menu_data(data):
    """The week's menu as plain data, shared by the page, menu.json and the order service."""
    start = _date(data.get("week_start"))
    days = []
    for item in data.get("menu") or []:
        day, meal = str(item.get("day") or "").strip(), str(item.get("meal") or "").strip()
        if not meal or day not in DAYS:
            continue
        date = start + dt.timedelta(days=DAYS.index(day)) if start else None
        days.append({"key": day[:3].lower(), "day": day, "meal": meal,
                     "date": date.isoformat() if date else "",
                     "date_label": f"{date:%b} {date.day}" if date else ""})
    deadline = _deadline(data.get("order_deadline"), start)
    week_label = str(data.get("week_of") or "").strip()
    if start and days and not week_label:
        week_label = _range_label(dt.date.fromisoformat(days[0]["date"]), dt.date.fromisoformat(days[-1]["date"]))
    return {
        "week_start": start.isoformat() if start else "",
        "week_label": week_label,
        "deadline": deadline.astimezone(dt.timezone.utc).isoformat() if deadline else "",
        "deadline_label": (f"{deadline:%A, %B} {deadline.day} at {deadline:%I:%M %p}".replace(" 0", " ")
                           if deadline else ""),
        "prices": {"regular": float(data.get("price_regular") or 8), "preschool": float(data.get("price_preschool") or 5)},
        "note": str(data.get("note") or "").strip(),
        "days": days,
        "grades": [{"value": v, "label": l, "preschool": p} for v, l, p in GRADES],
    }

def menu_file(data):
    return json.dumps(menu_data(data), indent=2) + "\n"

def can_order_online(menu):
    """Online ordering needs a week start date, a deadline and at least one meal."""
    return bool(menu["week_start"] and menu["deadline"] and menu["days"])

def lunch_block(data, root, api):
    menu = menu_data(data)
    price_line = (f"<p>Lunches are {money(menu['prices']['regular'])} "
                  f"({money(menu['prices']['preschool'])} for preschool).</p>")
    contact = ('<p>If your student is enrolled and you are not receiving the lunch form, please contact the '
               'school office at <a href="tel:3018561616">(301) 856-1616</a> or '
               '<a href="mailto:iba@ibcministries.org">iba@ibcministries.org</a>.</p>')
    if not menu["days"]:
        return ('<p>A lunch menu form is sent out via email/text one week in advance for students to select '
                'the days they would like to eat hot lunch. Orders are due by 9:00 AM on the Friday before.</p>' + price_line + contact +
                f'<img src="{root}img/hot-lunch-menu.jpg" alt="Sample weekly hot lunch menu with a different meal '
                'for each day, such as chicken parmesan, soft tacos, hot dogs, chili and burgers" loading="lazy" '
                'width="1000" height="562" style="border-radius:4px;margin-top:.8rem">')
    out = [f'<p class="lunch-week">Week of {esc(menu["week_label"])}</p>' if menu["week_label"] else ""]
    out.append('<table class="lunch"><tbody>' + "".join(
        f'<tr><th scope="row">{esc(d["day"])}'
        + (f' <span class="lunch-date">{esc(d["date_label"])}</span>' if d["date_label"] else "")
        + f'</th><td>{esc(d["meal"])}</td></tr>' for d in menu["days"]) + "</tbody></table>")
    out.append(price_line)
    if menu["note"]:
        out.append(f"<p>{esc(menu['note'])}</p>")
    if menu["deadline_label"]:
        out.append(f'<p class="lunch-by">Please order by {esc(menu["deadline_label"])}.</p>')
    if api and can_order_online(menu):
        out.append(order_form(menu, root, api))
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

def order_form(menu, root, api):
    grade_options = '<option value="">Choose…</option>' + "".join(
        f'<option value="{esc(g["value"])}">{esc(g["label"])}</option>' for g in menu["grades"])
    day_boxes = "".join(
        f'<label class="lunch-day"><input type="checkbox" name="days" value="{d["key"]}">'
        f'<span><b>{esc(d["day"])}</b> {esc(d["date_label"])}<br>{esc(d["meal"])}</span></label>'
        for d in menu["days"])
    child = (f'<fieldset class="lunch-child"><legend>Student</legend>'
             f'<label>Student’s full name<input name="student" autocomplete="off" maxlength="80" required></label>'
             f'<label>Grade<select name="grade" required>{grade_options}</select></label>'
             f'<div class="lunch-days" role="group" aria-label="Lunch days">{day_boxes}</div>'
             f'<button type="button" class="lunch-remove" hidden>Remove this student</button></fieldset>')
    return (
        f'<form class="info lunch-order" id="lunch-order" data-api="{esc(api)}" '
        f'data-menu="{root}lunch/menu.json" novalidate>'
        '<h3>Order hot lunch online</h3>'
        '<div id="lunch-closed" class="notice" hidden>Online ordering for this week is closed. '
        'To change or add an order, please call the office at <a href="tel:3018561616">(301) 856-1616</a>.</div>'
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
