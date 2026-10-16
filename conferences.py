"""Parent-teacher conferences: the booking page and the public schedule file.

The office edits content/conferences.json in the online editor (.pages.yml): the day, the hours,
the session length, the break between sessions and the teachers. build.py calls page_block() for
the Conferences page and schedule_file() for site/conferences/schedule.json, which the booking
service (worker/, the same service as hot lunch) reads to check every booking: teachers, times
and the deadline always come from that file, never from the parent's browser.
"""
import datetime as dt
import html
import json
import re
from zoneinfo import ZoneInfo

EASTERN = ZoneInfo("America/New_York")
USUAL_DEADLINE = dt.time(15, 0)  # bookings close at 3:00 PM on the school day before

def esc(value):
    return html.escape(str(value or "").strip())

def slug(value):
    return re.sub(r"[^a-z0-9]+", "-", str(value or "").lower()).strip("-")

def _date(value):
    try:
        return dt.date.fromisoformat(str(value or "").strip()[:10])
    except ValueError:
        return None

def _time(value):
    """'9:00 AM', '1 PM' or '13:00' -> minutes after midnight, else None."""
    m = re.fullmatch(r"\s*(\d{1,2})(?::(\d{2}))?\s*([AaPp])?\.?\s*[Mm]?\.?\s*", str(value or ""))
    if not m:
        return None
    h, mins, ampm = int(m.group(1)), int(m.group(2) or 0), (m.group(3) or "").upper()
    if ampm:
        if not 1 <= h <= 12:
            return None
        h = h % 12 + (12 if ampm == "P" else 0)
    if h > 23 or mins > 59:
        return None
    return h * 60 + mins

def time_label(minutes):
    h, m = divmod(minutes, 60)
    return f"{(h + 11) % 12 + 1}:{m:02d} {'AM' if h < 12 else 'PM'}"

def _deadline(value, day):
    """The office's deadline if they typed one, else 3:00 PM on the weekday before the conference day."""
    try:
        return dt.datetime.fromisoformat(str(value or "").strip()).replace(tzinfo=EASTERN)
    except ValueError:
        if not day:
            return None
        before = day - dt.timedelta(days=1)
        while before.weekday() > 4:
            before -= dt.timedelta(days=1)
        return dt.datetime.combine(before, USUAL_DEADLINE, tzinfo=EASTERN)

def _when(d):
    return f"{d:%A, %B} {d.day} at {time_label(d.hour * 60 + d.minute)}"

def schedule_data(data):
    """The conference day as plain data, shared by the page, schedule.json and the booking service."""
    day = _date(data.get("date"))
    start, end = _time(data.get("start")), _time(data.get("end"))
    length = int(data.get("minutes") or 0)
    gap = int(data.get("break_minutes") or 0)
    slots = []
    if day and start is not None and end is not None and length > 0:
        t = start
        while t + length <= end:
            slots.append({"id": f"{t // 60:02d}{t % 60:02d}", "label": time_label(t), "end": time_label(t + length)})
            t += length + max(gap, 0)
    teachers, seen = [], set()
    for item in data.get("teachers") or []:
        name = str(item.get("name") or "").strip()
        if not name:
            continue
        tid, n = slug(name) or "teacher", 2
        while tid in seen:
            tid, n = f"{slug(name)}-{n}", n + 1
        seen.add(tid)
        # A shared panel (High School: three teachers meet each family together) lists its teachers;
        # each of them gets the booking email. A one-teacher panel is just its own name.
        members = [m.strip() for m in re.split(r"[,;\n]", str(item.get("members") or "")) if m.strip()]
        teachers.append({"id": tid, "name": name, "subject": str(item.get("subject") or "").strip(),
                         "members": members or [name]})
    deadline = _deadline(data.get("booking_deadline"), day)
    return {
        "date": day.isoformat() if day else "",
        "date_label": f"{day:%A, %B} {day.day}, {day.year}" if day else "",
        "minutes": length,
        "deadline": deadline.astimezone(dt.timezone.utc).isoformat() if deadline else "",
        "deadline_label": _when(deadline) if deadline else "",
        "note": str(data.get("note") or "").strip(),
        "slots": slots,
        "teachers": teachers,
    }

def schedule_file(data):
    return json.dumps(schedule_data(data), indent=2) + "\n"

def is_open(s, now=None):
    now = now or dt.datetime.now(dt.timezone.utc)
    return bool(s["slots"] and s["teachers"] and s["deadline"]
                and now < dt.datetime.fromisoformat(s["deadline"]))

def page_block(data, root, api):
    s = schedule_data(data)
    contact = ('<p>Questions, or need a different time? Please call the school office at '
               '<a href="tel:3018561616">(301) 856-1616</a> or email '
               '<a href="mailto:iba@ibcministries.org">iba@ibcministries.org</a>.</p>')
    if not s["slots"] or not s["teachers"]:
        return '<p class="lead">Conference sign-ups aren’t open right now.</p>' + contact
    first, last = s["slots"][0], s["slots"][-1]
    out = [f'<p class="conf-when">{esc(s["date_label"])}, {esc(first["label"])} to {esc(last["end"])}</p>',
           f'<p class="lead">Each conference is {s["minutes"]} minutes. Choose your child’s teacher, then pick an open time.'
           + (" If you have children in more than one class, book a time with each teacher." if len(s["teachers"]) > 1 else "")
           + "</p>"]
    if s["note"]:
        out.append(f"<p>{esc(s['note'])}</p>")
    if s["deadline_label"]:
        out.append(f'<p class="lunch-by">Please sign up by {esc(s["deadline_label"])}.</p>')
    # Without the booking service (or without JavaScript) the page still lists the teachers and times.
    teachers = "".join(
        f'<li><a class="conf-teacher" href="#{t["id"]}" data-id="{t["id"]}"><b>{esc(t["name"])}</b>'
        + (f'<span>{esc(t["subject"])}</span>' if t["subject"] else "") + "</a></li>"
        for t in s["teachers"])
    out.append(f'<div id="conf-app" data-api="{esc(api)}" data-schedule="{root}conferences/schedule.json">'
               f'<h2 id="conf-pick">Choose a teacher</h2><ul class="conf-teachers">{teachers}</ul>'
               '<div id="conf-panel" hidden></div></div>')
    if api:
        out.append('<noscript><p class="notice">Signing up online needs JavaScript. Please call the office at '
                   '(301) 856-1616 to book a time.</p></noscript>'
                   f'<script src="{root}conferences.js" defer></script>')
    else:
        out.append('<p class="notice">To book a time, please call the office.</p>')
    return "".join(out) + contact
