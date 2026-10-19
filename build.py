#!/usr/bin/env python3
"""Build the site: wrap each page in src/pages/ with src/layout.html and write it to site/.

Run `python3 build.py` after editing anything in src/. The generated files in site/ are
committed, so what is in git is exactly what gets published.

Each page file starts with a few `key: value` lines, then a line with `---`, then the HTML body.
Keys: title, description, nav (which menu item to highlight), cta (no = hide the
"Explore the Possibilities" band), hero (image file in img/ for the page banner), banner (no = no banner),
fulltitle (the browser tab title, instead of "<title> | Independent Baptist Academy"), fonts (Google Fonts
families for this page only).
"""
import html, json, pathlib, re
from html.parser import HTMLParser

import conferences
import lunch
import tours

# The site's own domain. Leave empty while the site is only a preview at
# guarantee418.github.io/ibaclinton-site/. Set to "ibaclinton.com" on go-live day and rebuild:
# that writes site/CNAME, switches 404.html's links to "/", turns on Google Analytics and
# lets search engines index the pages (the preview asks them not to).
DOMAIN = "ibaclinton.com"

# Address of the hot lunch order service (worker/, a Cloudflare Worker; it also takes the conference, tour,
# transcript and Request Info forms), for example
# "https://ibaclinton-lunch.<account>.workers.dev". Empty = no online ordering: the page shows
# the menu and, if the office pasted one, the Microsoft Forms link instead.
LUNCH_API = "https://ibaclinton-lunch.guarantee418.workers.dev"

# Google tag from the WordPress site (added there by Site Kit). Reusing it keeps the
# Analytics history in one place. Only added to pages when DOMAIN is set.
GA_ID = "GT-PJ5L2Z5R"

# Address used in sitemap.xml and robots.txt; search engines only read those on the real domain.
SITE_URL = "https://ibaclinton.com/"

# Where the site is served from. Used only by 404.html, which GitHub Pages shows at any
# path, so its links can't be relative.
BASE_PATH = "/" if DOMAIN else "/ibaclinton-site/"

GTAG = f"""<script async src="https://www.googletagmanager.com/gtag/js?id={GA_ID}"></script>
<script>window.dataLayer=window.dataLayer||[];function gtag(){{dataLayer.push(arguments);}}gtag("js",new Date());gtag("config","{GA_ID}");</script>"""

ROOT = pathlib.Path(__file__).parent
SRC, OUT, CONTENT = ROOT / "src", ROOT / "site", ROOT / "content"
# styles.css is linked as styles.css?v=<this>, so browsers (Cloudflare tells them to keep the file 4 hours)
# load the new styles as soon as they change instead of drawing new pages with old styles.
CSS_VERSION = __import__("hashlib").sha256((OUT / "styles.css").read_bytes()).hexdigest()[:10]
NAV = [("parents", "parents/", "Parents"), ("about", "about/", "About"),
       ("staff", "about-copy/", "Staff"), ("academics", "academics/", "Academics"),
       ("admissions", "admissions/", "Admissions"), ("student-life", "student-life/", "Student Life"),
       ("give", "https://secure.anedot.com/ibc/academy-donations", "Give")]

# content/*.json is edited by the school office through the online editor (.pages.yml).
# Everything from it is escaped, and only https links are used.
def esc(value):
    return html.escape(str(value or "").strip())

def safe_url(value):
    value = str(value or "").strip()
    return value if value.startswith("https://") else ""

def docs_list(items, root):
    lis = []
    for item in items or []:
        title, url = esc(item.get("title")), safe_url(item.get("url"))
        if not url and item.get("file"):
            url = f"{root}files/" + html.escape(pathlib.PurePosixPath(str(item["file"])).name, quote=True)
        if title and url:
            lis.append(f'<li><a href="{url}">{title}</a></li>')
    return '<ul class="docs">' + "".join(lis) + "</ul>" if lis else ""

# The Kings Corner (parents page, content/parents.json): the office edits the hero, the tiles and every section
# in the online editor.
# Section text is rich text (HTML) from the editor; only simple formatting tags and https,
# mailto and tel links are kept.
ALLOWED_TAGS = {"p", "br", "b", "strong", "i", "em", "u", "ul", "ol", "li", "a", "h3", "h4", "blockquote"}

class _Clean(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out, self.skip = [], 0
    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self.skip += 1
        if self.skip or tag not in ALLOWED_TAGS:
            return
        if tag == "a":
            href = str(dict(attrs).get("href") or "").strip()
            ok = href.startswith(("https://", "mailto:", "tel:"))
            self.out.append(f'<a href="{html.escape(href, quote=True)}">' if ok else "<a>")
        else:
            self.out.append(f"<{tag}>")
    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self.skip = max(0, self.skip - 1)
        elif not self.skip and tag in ALLOWED_TAGS and tag != "br":
            self.out.append(f"</{tag}>")
    def handle_data(self, data):
        if not self.skip:
            self.out.append(html.escape(data, quote=False))

def clean_html(value):
    c = _Clean()
    c.feed(str(value or ""))
    c.close()
    return "".join(c.out).strip()

def slug(value):
    return re.sub(r"[^a-z0-9]+", "-", str(value or "").lower()).strip("-")

FEEDBACK_TOPICS = ["Academics", "Communication", "Discipline", "Activities & athletics", "Hot lunch",
                   "Facilities & safety", "Suggestion", "Praise", "Other"]  # same list as worker/src/feedback.js

def feedback_form(root):
    """The Feedback for Administration form; responses go to the Feedback tab of the office workbook."""
    topics = "".join(f"<option>{esc(t)}</option>" for t in FEEDBACK_TOPICS)
    return (
        '<div id="fb-status" role="status" aria-live="polite"></div>'
        f'<form class="info fb-form" id="fb-form" data-api="{esc(LUNCH_API)}" novalidate hidden>'
        '<fieldset><legend>Your feedback</legend>'
        f'<label class="full">Topic<select name="topic" required><option value="">Choose…</option>{topics}</select></label>'
        '<label class="full">Feedback<textarea name="message" maxlength="4000" required></textarea></label>'
        '</fieldset>'
        '<fieldset><legend>About you <span class="opt">(optional)</span></legend>'
        '<p class="full fb-note">Leave these blank to send your feedback without your name. '
        'Add your email or phone if you would like a reply.</p>'
        '<label>Name<input name="name" autocomplete="name" maxlength="120"></label>'
        '<label>Child’s grade(s)<input name="grades" maxlength="60" placeholder="for example 3rd and 7th"></label>'
        '<label>Email<input type="email" name="email" autocomplete="email" maxlength="120"></label>'
        '<label>Phone<input type="tel" name="phone" autocomplete="tel" maxlength="30"></label>'
        '</fieldset>'
        '<label class="hp" aria-hidden="true">Leave this field blank<input name="website" tabindex="-1" autocomplete="off"></label>'
        '<button class="btn" type="submit">Send feedback</button>'
        '</form>'
        '<noscript><p class="notice">To share feedback, please email '
        '<a href="mailto:iba@ibcministries.org">iba@ibcministries.org</a> or call (301) 856-1616.</p></noscript>'
        f'<script src="{root}feedback.js" defer></script>')

ARROW = ('<svg class="kc-arrow" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
         'stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
         '<path d="M4 12h16"/><path d="M14 6l6 6-6 6"/></svg>')

def parents_block(data, root):
    """The Kings Corner (parents page): hero and tile grid; each tile opens its section in a pop-up (site/kings.js).
    All text comes from content/parents.json. Without JavaScript the sections show below the tiles instead."""
    sections = [s for s in data.get("sections") or [] if str(s.get("title") or "").strip()]
    for s in sections:
        s["_id"] = slug(s.get("anchor")) or slug(s.get("title"))
    ids = {s["_id"] for s in sections}

    hero = data.get("hero") or {}
    button = ""
    if esc(hero.get("button_text")) and safe_url(hero.get("button_url")):
        button = f'<a class="kc-btn" href="{esc(safe_url(hero["button_url"]))}">{esc(hero["button_text"])}</a>'
    hero_html = (
        '<div class="kc-hero"><div class="kc-in kc-hero-in"><div class="kc-hero-text">'
        + (f'<p class="kc-script">{esc(hero.get("script"))}</p>' if esc(hero.get("script")) else "")
        + f'<h1>{esc(hero.get("title")) or "The Kings Corner"}</h1>'
        + (f'<p class="kc-lead">{esc(hero.get("text"))}</p>' if esc(hero.get("text")) else "")
        + button + '</div>'
        f'<img class="kc-crest" src="{root}img/kings-crest.png" alt="Independent Baptist Academy Kings crest" width="600" height="450">'
        '</div></div>')

    # A tile opens a section (its page link name) or goes to a full https:// address.
    # A tile whose section no longer exists is left out, so no tile is a dead link; a section that no tile opens
    # gets a plain tile at the end, so every section can be reached.
    cards, used = [], set()
    for t in data.get("tiles") or []:
        title, target = esc(t.get("title")), str(t.get("section") or "").strip()
        href = esc(safe_url(target)) or (f"#{slug(target)}" if slug(target) in ids else "")
        if not title or not href:
            if title:
                print(f"  tile {title!r} left out: no section {target!r}")
            continue
        used.add(slug(target))
        tag = f'<span class="kc-tag">{esc(t.get("tag"))}</span>' if esc(t.get("tag")) else ""
        desc = f'<span class="kc-desc">{esc(t.get("description"))}</span>' if esc(t.get("description")) else ""
        cards.append(f'<li><a class="kc-tile" href="{href}"><span class="kc-tile-top"><span class="kc-tile-title">{title}</span>'
                     f'{tag}</span>{desc}{ARROW}</a></li>')
    for s in sections:
        if s["_id"] not in used:
            cards.append(f'<li><a class="kc-tile" href="#{s["_id"]}"><span class="kc-tile-top"><span class="kc-tile-title">'
                         f'{esc(s["title"])}</span></span>{ARROW}</a></li>')
    tiles_html = ""
    if cards:
        tiles_html = (
            '<div class="kc-in kc-find">'
            + (f'<p class="kc-script kc-eyebrow">{esc(data.get("tiles_script"))}</p>' if esc(data.get("tiles_script")) else "")
            + (f'<h2 class="kc-h2">{esc(data.get("tiles_heading"))}</h2>' if esc(data.get("tiles_heading")) else "")
            + f'<ul class="kc-tiles">{"".join(cards)}</ul></div>')

    blocks = []
    for s in sections:
        extra = ""
        if s.get("show_lunch"):
            extra = lunch.lunch_block(json.loads((CONTENT / "lunch.json").read_text()), root, LUNCH_API)
        if s.get("show_feedback"):
            extra += feedback_form(root)
        blocks.append(f'<div class="block" id="{s["_id"]}"><h2>{esc(s["title"])}</h2>'
                      f'{clean_html(s.get("text"))}{docs_list(s.get("links"), root)}{extra}</div>')
    back = f'Back to {esc(hero.get("title")) or "The Kings Corner"}'
    return (f'<div class="kc" data-back="{back}">{hero_html}{tiles_html}'
            f'\n<div class="kc-in kc-sections"><div class="blocks">\n{"\n".join(blocks)}\n</div></div>\n</div>'
            f'\n<script src="{root}kings.js" defer></script>')

def render(path, text, layout):
    head, body = text.split("\n---\n", 1)
    meta = dict(line.split(": ", 1) for line in head.strip().splitlines())
    rel = path.relative_to(SRC / "pages").with_suffix("")
    out = OUT / ("index.html" if rel.name == "index" and rel.parent == pathlib.Path(".") else
                 "404.html" if rel.name == "404" else rel / "index.html")
    depth = len(out.relative_to(OUT).parts) - 1
    root = "../" * depth or "./"
    if rel.name == "404":
        root = BASE_PATH
    nav = "".join(
        f'<li><a href="{href if href.startswith("http") else root + href}"'
        f'{" aria-current=\"page\"" if key == meta.get("nav") else ""}>{label}</a></li>'
        for key, href, label in NAV)
    cta = "" if meta.get("cta") == "no" else (SRC / "cta.html").read_text()
    hero = meta.get("hero")
    banner = "" if meta.get("banner") == "no" else (
        f'<div class="page-hero"{f" style=\"--bg:url({root}img/{hero})\"" if hero else ""}>'
        f'<div class="wrap"><h1>{meta["title"]}</h1></div></div>')
    title = meta.get("fulltitle") or f'{meta["title"]} | Independent Baptist Academy'
    if "{{lunch}}" in body:
        body = body.replace("{{lunch}}", lunch.lunch_block(json.loads((CONTENT / "lunch.json").read_text()), root, LUNCH_API))
    if "{{conferences}}" in body:
        body = body.replace("{{conferences}}", conferences.page_block(
            json.loads((CONTENT / "conferences.json").read_text()), root, LUNCH_API))
    if "{{tours}}" in body:
        body = body.replace("{{tours}}", tours.page_block(json.loads((CONTENT / "tours.json").read_text()), root, LUNCH_API))
    body = body.replace("{{api}}", esc(LUNCH_API))
    if "{{parents}}" in body:
        body = body.replace("{{parents}}", parents_block(json.loads((CONTENT / "parents.json").read_text()), root))
    tracking = ("This website uses Google Analytics to count visits and see which pages are used. "
        "Google Analytics sets cookies to do this; it does not tell us who you are. "
        "The site does not use advertising cookies."
        if DOMAIN and GA_ID else "This website does not use advertising or tracking cookies of its own.")
    body = body.replace("{{tracking}}", tracking)
    head_extra = GTAG if DOMAIN and GA_ID else '<meta name="robots" content="noindex">'
    if meta.get("fonts"):  # Google Fonts for this page only, e.g. "fonts: Yellowtail&family=Nunito+Sans"
        head_extra += ('\n<link rel="preconnect" href="https://fonts.googleapis.com">'
                       '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>'
                       f'<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family={esc(meta["fonts"])}&display=swap">')
    html = (layout.replace("{{title}}", title).replace("{{head}}", head_extra).replace("{{description}}", meta["description"])
            .replace("{{nav}}", nav).replace("{{banner}}", banner)
            .replace("{{content}}", body.strip()).replace("{{cta}}", cta))
    html = html.replace("{{root}}", root).replace("{{css_version}}", CSS_VERSION)
    leftover = re.findall(r"\{\{\w+\}\}", html)
    if leftover:
        raise SystemExit(f"{path}: unfilled {leftover}")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html)
    print("wrote", out.relative_to(ROOT))
    return None if rel.name == "404" else out.relative_to(OUT).parent.as_posix()

def write(name, text):
    (OUT / name).write_text(text)
    print("wrote", (OUT / name).relative_to(ROOT))

if __name__ == "__main__":
    layout = (SRC / "layout.html").read_text()
    paths = [render(p, p.read_text(), layout) for p in sorted((SRC / "pages").rglob("*.html"))]
    urls = [SITE_URL + ("" if d == "." else d + "/") for d in paths if d]
    write("sitemap.xml", '<?xml version="1.0" encoding="UTF-8"?>\n'
          '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
          + "".join(f"<url><loc>{u}</loc></url>\n" for u in urls) + "</urlset>\n")
    (OUT / "lunch").mkdir(exist_ok=True)
    write("lunch/menu.json", lunch.menu_file(json.loads((CONTENT / "lunch.json").read_text())))
    (OUT / "conferences").mkdir(exist_ok=True)
    write("conferences/schedule.json", conferences.schedule_file(json.loads((CONTENT / "conferences.json").read_text())))
    (OUT / "tours").mkdir(exist_ok=True)
    write("tours/schedule.json", tours.schedule_file(json.loads((CONTENT / "tours.json").read_text())))
    write("robots.txt", f"User-agent: *\nAllow: /\n\nSitemap: {SITE_URL}sitemap.xml\n")
    cname = OUT / "CNAME"
    if DOMAIN:
        write("CNAME", DOMAIN + "\n")
    elif cname.exists():
        cname.unlink()
