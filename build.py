#!/usr/bin/env python3
"""Build the site: wrap each page in src/pages/ with src/layout.html and write it to site/.

Run `python3 build.py` after editing anything in src/. The generated files in site/ are
committed, so what is in git is exactly what gets published.

Each page file starts with a few `key: value` lines, then a line with `---`, then the HTML body.
Keys: title, description, nav (which menu item to highlight), cta (no = hide the
"Explore the Possibilities" band), hero (image file in img/ for the page banner).
"""
import pathlib, re

# Formspree form ID for the Request Info form (the part after formspree.io/f/).
# Leave empty until the form exists; the page then shows a "call or email us" notice.
FORMSPREE_ID = ""

# Where the site is served from. Used only by 404.html, which GitHub Pages shows at any
# path, so its links can't be relative. Change to "/" when ibaclinton.com points here.
BASE_PATH = "/ibaclinton-site/"

ROOT = pathlib.Path(__file__).parent
SRC, OUT = ROOT / "src", ROOT / "site"
NAV = [("parents", "parents/", "Parents"), ("about", "about/", "About"),
       ("staff", "about-copy/", "Staff"), ("academics", "academics/", "Academics"),
       ("admissions", "admissions/", "Admissions"), ("student-life", "student-life/", "Student Life"),
       ("give", "https://secure.anedot.com/ibc/academy-donations", "Give")]

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
    if FORMSPREE_ID:
        body = body.replace("{{formspree}}", f"https://formspree.io/f/{FORMSPREE_ID}").replace("{{form_notice}}", "")
    else:
        body = body.replace("{{formspree}}", "#").replace("{{form_notice}}",
            '<p class="notice">This form is being set up. For now, please call '
            '<a href="tel:3018561616">(301) 856-1616</a> or email '
            '<a href="mailto:iba@ibcministries.org">iba@ibcministries.org</a>.</p>')
    html = (layout.replace("{{title}}", title).replace("{{description}}", meta["description"])
            .replace("{{nav}}", nav).replace("{{banner}}", banner)
            .replace("{{content}}", body.strip()).replace("{{cta}}", cta))
    html = html.replace("{{root}}", root)
    leftover = re.findall(r"\{\{\w+\}\}", html)
    if leftover:
        raise SystemExit(f"{path}: unfilled {leftover}")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html)
    print("wrote", out.relative_to(ROOT))

if __name__ == "__main__":
    layout = (SRC / "layout.html").read_text()
    for p in sorted((SRC / "pages").rglob("*.html")):
        render(p, p.read_text(), layout)
