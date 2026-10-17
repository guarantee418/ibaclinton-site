#!/usr/bin/env python3
"""Build the site: wrap each page in src/pages/ with src/layout.html and write it to site/.

Run `python3 build.py` after editing anything in src/. The generated files in site/ are
committed, so what is in git is exactly what gets published.

Each page file starts with a few `key: value` lines, then a line with `---`, then the HTML body.
Keys: title, description, nav (which menu item to highlight), cta (no = hide the
"Explore the Possibilities" band), hero (image file in img/ for the page banner).
"""
import html, json, pathlib, re
from html.parser import HTMLParser

import conferences
import lunch
import tours

# Formspree form ID for the Request Info form (the part after formspree.io/f/).
# Leave empty until the form exists; the page then shows a "call or email us" notice.
FORMSPREE_ID = "xaeqabaz"  # Formspree account: ggreene@ibcministries.org

# The site's own domain. Leave empty while the site is only a preview at
# guarantee418.github.io/ibaclinton-site/. Set to "ibaclinton.com" on go-live day and rebuild:
# that writes site/CNAME, switches 404.html's links to "/", turns on Google Analytics and
# lets search engines index the pages (the preview asks them not to).
DOMAIN = "ibaclinton.com"

# Address of the hot lunch order service (worker/, a Cloudflare Worker), for example
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

# Parent Portal (content/parents.json): the office edits every section in the online editor.
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

def parents_block(data, root):
    sections = [s for s in data.get("sections") or [] if str(s.get("title") or "").strip()]
    for s in sections:
        s["_id"] = slug(s.get("anchor")) or slug(s.get("menu_label")) or slug(s.get("title"))
        s["_label"] = str(s.get("menu_label") or "").strip() or str(s["title"]).strip()
    jump = "".join(f'<li><a href="#{s["_id"]}">{esc(s["_label"])}</a></li>'
                   for s in sorted(sections, key=lambda s: s["_label"].lower()))
    blocks = []
    for s in sections:
        extra = ""
        if s.get("show_lunch"):
            extra = lunch.lunch_block(json.loads((CONTENT / "lunch.json").read_text()), root, LUNCH_API)
        blocks.append(f'<div class="block" id="{s["_id"]}"><h2>{esc(s["title"])}</h2>'
                      f'{clean_html(s.get("text"))}{docs_list(s.get("links"), root)}{extra}</div>')
    return (f'<h2>{esc(data.get("heading"))}</h2>'
            + (f'<p class="lead">{esc(data.get("intro"))}</p>' if str(data.get("intro") or "").strip() else "")
            + (f'<ul class="jump">{jump}</ul>' if jump else "")
            + f'\n<div class="blocks">\n{"\n".join(blocks)}\n</div>')

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
    if "{{parents}}" in body:
        body = body.replace("{{parents}}", parents_block(json.loads((CONTENT / "parents.json").read_text()), root))
    if FORMSPREE_ID:
        body = body.replace("{{formspree}}", f"https://formspree.io/f/{FORMSPREE_ID}").replace("{{form_notice}}", "")
    else:
        body = body.replace("{{formspree}}", "#").replace("{{form_notice}}",
            '<p class="notice">This form is being set up. For now, please call '
            '<a href="tel:3018561616">(301) 856-1616</a> or email '
            '<a href="mailto:iba@ibcministries.org">iba@ibcministries.org</a>.</p>')
    tracking = ("This website uses Google Analytics to count visits and see which pages are used. "
        "Google Analytics sets cookies to do this; it does not tell us who you are. "
        "The site does not use advertising cookies."
        if DOMAIN and GA_ID else "This website does not use advertising or tracking cookies of its own.")
    body = body.replace("{{tracking}}", tracking)
    head_extra = GTAG if DOMAIN and GA_ID else '<meta name="robots" content="noindex">'
    html = (layout.replace("{{title}}", title).replace("{{head}}", head_extra).replace("{{description}}", meta["description"])
            .replace("{{nav}}", nav).replace("{{banner}}", banner)
            .replace("{{content}}", body.strip()).replace("{{cta}}", cta))
    html = html.replace("{{root}}", root)
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
