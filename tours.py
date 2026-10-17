"""School tours: the sign-up page and the public tour list.

The office edits content/tours.json in the online editor (.pages.yml): each tour's date, time and
number of families (spots). build.py calls page_block() for the Book a Tour page and
schedule_file() for site/tours/schedule.json, which the sign-up service (worker/, the same service
as hot lunch and conferences) reads to check every sign-up: tours, spots and deadlines always come
from that file, never from the visitor's browser. A tour is full once it has as many sign-ups as spots.
"""
import datetime as dt
import html
import json

from conferences import EASTERN, USUAL_DEADLINE, _date, _time, time_label

DEFAULT_SPOTS = 5

def esc(value):
    return html.escape(str(value or "").strip())

def _deadline(value, day):
    """The office's deadline for a tour if typed, else 3:00 PM on the weekday before it."""
    try:
        return dt.datetime.fromisoformat(str(value or "").strip()).replace(tzinfo=EASTERN)
    except ValueError:
        before = day - dt.timedelta(days=1)
        while before.weekday() > 4:
            before -= dt.timedelta(days=1)
        return dt.datetime.combine(before, USUAL_DEADLINE, tzinfo=EASTERN)

def schedule_data(data):
    default_spots = int(data.get("spots") or DEFAULT_SPOTS)
    tours = []
    for item in data.get("tours") or []:
        day, start = _date(item.get("date")), _time(item.get("time"))
        if not day or start is None:
            continue
        starts = dt.datetime.combine(day, dt.time(start // 60, start % 60), tzinfo=EASTERN)
        deadline = _deadline(item.get("signup_deadline"), day)
        label = time_label(start)
        tours.append({
            "id": f"{day.isoformat()}-{start // 60:02d}{start % 60:02d}",
            "date": day.isoformat(), "time": label,
            "label": f"{day:%A, %B} {day.day}, {day.year} at {label}",
            "starts": starts.astimezone(dt.timezone.utc).isoformat(),
            "spots": max(int(item.get("spots") or default_spots), 1),
            "deadline": deadline.astimezone(dt.timezone.utc).isoformat(),
            "deadline_label": f"{deadline:%A, %B} {deadline.day} at {time_label(deadline.hour * 60 + deadline.minute)}",
            "note": str(item.get("note") or "").strip(),
        })
    tours.sort(key=lambda t: t["starts"])
    return {"note": str(data.get("note") or "").strip(), "tours": tours}

def schedule_file(data):
    return json.dumps(schedule_data(data), indent=2) + "\n"

def page_block(data, root, api):
    s = schedule_data(data)
    now = dt.datetime.now(dt.timezone.utc).isoformat()
    upcoming = [t for t in s["tours"] if t["starts"] > now]
    out = []
    if s["note"]:
        out.append(f"<p>{esc(s['note'])}</p>")
    if not upcoming:
        out.append('<p class="notice">No tours are scheduled right now. Please call the office to set up a visit.</p>')
        return "".join(out)
    # Without the sign-up service (or without JavaScript) the page still lists the tours.
    items = "".join(f'<li data-id="{t["id"]}"><b>{esc(t["label"])}</b>'
                    + (f'<span>{esc(t["note"])}</span>' if t["note"] else "") + "</li>" for t in upcoming)
    out.append(f'<div id="tour-app" data-api="{esc(api)}" data-schedule="{root}tours/schedule.json">'
               f'<h2>Upcoming tours</h2><ul class="tour-list">{items}</ul><div id="tour-panel"></div></div>')
    if api:
        out.append('<noscript><p class="notice">Signing up online needs JavaScript. Please call the office at '
                   '(301) 856-1616 to join a tour.</p></noscript>'
                   f'<script src="{root}tours.js" defer></script>')
    else:
        out.append('<p class="notice">To join a tour, please call the office.</p>')
    return "".join(out)
