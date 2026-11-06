#!/usr/bin/env python3
"""
build_site.py — Generate a searchable GitHub Pages site from Dragonfly plugin .docx manuals.

Usage:
    python3 build/build_site.py [--src /Users/user/DragonflyPlugins] [--out .]

What it does (idempotent, safe to re-run):
  1. Scans --src (recursively) for *.docx files.
  2. Converts each manual to a standalone HTML page in pages/<slug>.html:
       - cover block (title / subtitle / version)
       - auto-generated table of contents with anchors
       - headings, paragraphs, bullet lists, tables, embedded images
  3. DEDUPE: manuals whose full text content is byte-identical (after whitespace
     normalization) are collapsed into ONE page (canonical = alphabetically
     first source file); all source files are listed in the page footer and
     the homepage shows the duplicates as aliases instead of duplicate cards.
  4. Emits index.html (searchable catalogue), search-data.json (full-text
     search payloads), style.css, and refreshes README.md.

Adding new manuals in the future:
    drop new .docx files into the src dir, re-run this script, git push.
New files appear automatically; changed files get rebuilt; identical files
merge into an existing page.
"""

import argparse
import glob
import hashlib
import html
import json
import os
import re
import sys
import zipfile

try:
    import docx
    from docx.oxml.ns import qn
except ImportError:
    sys.exit("python-docx is required:  pip install python-docx")

MAIN_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def slugify(text: str) -> str:
    s = text.lower().replace(".docx", "")
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return re.sub(r"-{2,}", "-", s) or "page"


def anchorify(text: str) -> str:
    s = re.sub(r"\s+", " ", text).strip().lower()
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s or "sec"


def norm_ws(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def extract_content_hash(d: "docx.Document") -> str:
    """Stable hash of the manual's full textual content (dedupe key)."""
    parts = []
    for p in d.paragraphs:
        t = norm_ws(p.text)
        if t:
            parts.append(t)
    for tb in d.tables:
        for row in tb.rows:
            for c in row.cells:
                t = norm_ws(c.text)
                if t:
                    parts.append(t)
    return hashlib.sha256("\n".join(parts).encode("utf-8")).hexdigest()


def runs_to_html(p, manual: "Manual | None" = None) -> str:
    """Convert a paragraph's runs to HTML with bold/italic; emits <figure> for embedded images."""
    out = []
    for run in p.runs:
        blips = run._element.findall(f".//{{{A_NS}}}blip")
        if blips:
            for blip in blips:
                rid = blip.get(f"{{{R_NS}}}embed")
                if not rid or manual is None:
                    continue
                rel, width = manual.emit_image(rid)
                if not rel:
                    continue
                wattr = f' width="{width}"' if width else ""
                out.append(f'<figure><img src="{rel}" alt=""{wattr} loading="lazy"></figure>')
            continue
        text = run.text
        if not text:
            brs = run._element.findall(f"{{{MAIN_NS}}}br")
            if brs:
                out.append("<br>" * len(brs))
            continue
        frag = html.escape(text)
        if run.bold:
            frag = f"<strong>{frag}</strong>"
        if run.italic:
            frag = f"<em>{frag}</em>"
        out.append(frag)
    return "".join(out)


def _draw_containing_blip(doc, rid: str):
    """Return the wp:anchor / wp:inline drawing element whose blip uses rid."""
    wp_ns = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    for tag in (f"{{{wp_ns}}}anchor", f"{{{wp_ns}}}inline", f"{{{wp_ns}}}pict"):
        for el in doc.element.body.iter(tag):
            for blip in el.iter(f"{{{A_NS}}}blip"):
                if blip.get(f"{{{R_NS}}}embed") == rid:
                    return el
    return None


def image_display_width_px(draw_el) -> int | None:
    """Read a:ext cx from a drawing's first spPr/extent, EMU → px @96dpi."""
    for ext in draw_el.findall(f".//{{{A_NS}}}ext"):
        cx = ext.get("cx")
        if cx:
            return max(16, min(760, int(cx) // 9525))
    return None


class Manual:
    def __init__(self, path: str):
        self.path = path
        self.fname = os.path.basename(path)
        self.slug = slugify(self.fname)
        self.doc = docx.Document(path)
        self.asset_dir = None      # set by build(): <out>/assets/<slug>
        self._img_counter = 0
        self._seen_rids: set[str] = set()

    def set_out(self, out_dir: str):
        self.asset_dir = os.path.join(out_dir, "assets", self.slug)
        os.makedirs(self.asset_dir, exist_ok=True)

    def emit_image(self, rid: str) -> tuple[str | None, int | None]:
        """Write the image blob to assets/<slug>/ and return (relpath, width_px)."""
        if rid in self._seen_rids:
            return None, None  # same image referenced again (rare); skip duplicate
        self._seen_rids.add(rid)
        try:
            part = self.doc.part.rels[rid].target_part
        except KeyError:
            return None, None
        name = part.partname  # e.g. /word/media/manual_image_1.png
        ext = os.path.splitext(str(name))[1].lower() or ".png"
        self._img_counter += 1
        dest = os.path.join(self.asset_dir, f"img{self._img_counter}{ext}")
        try:
            with open(dest, "wb") as fh:
                fh.write(part.blob)
        except Exception:
            return None, None
        width = None
        draw = _draw_containing_blip(self.doc, rid)
        if draw is not None:
            width = image_display_width_px(draw)
        return f"../assets/{self.slug}/img{self._img_counter}{ext}", width

    def cover(self):
        ne = [p for p in self.doc.paragraphs if p.text.strip()]
        title = ne[0].text.strip() if ne else self.fname
        subtitle = ne[1].text.strip() if len(ne) > 1 else ""
        version = ne[2].text.strip() if len(ne) > 2 else ""
        # strip trailing " - User Manual" for a cleaner card title
        short = re.sub(r"\s*[-–—]\s*User Manual\s*$", "", title)
        return short, subtitle, version

    # ---- body ----------------------------------------------------------
    def iter_body_items(self):
        """Yield ('p', Paragraph) / ('tbl', Table) in true document order."""
        for child in self.doc.element.body.iterchildren():
            if child.tag == f"{{{MAIN_NS}}}p":
                yield "p", docx.text.paragraph.Paragraph(child, self.doc)
            elif child.tag == f"{{{MAIN_NS}}}tbl":
                yield "tbl", docx.table.Table(child, self.doc)

    def image_relpath(self, rid: str) -> str | None:
        try:
            part = self.doc.part.rels[rid].target_part
        except KeyError:
            return None
        name = part.partname  # /word/media/manual_image_1.png
        ext = os.path.splitext(name)[1] or ".png"
        self._img_counter += 1
        rel = f"../assets/{self.slug}/img{self._img_counter}{ext}"
        self.images.append(rel)
        try:
            data = part.blob
        except Exception:
            return None
        return rel, data

    def plain_text(self) -> str:
        """Full normalized body text for the search index."""
        parts = []
        for kind, el in self.iter_body_items():
            if kind == "p":
                t = norm_ws(el.text)
                if t:
                    parts.append(t)
            else:
                for row in el.rows:
                    for c in row.cells:
                        t = norm_ws(c.text)
                        if t:
                            parts.append(t)
        return " ".join(parts)

    def render_body(self) -> tuple[str, list[dict]]:
        """Render body HTML; returns (html, toc_entries[{level,text,anchor}])."""
        body = []
        toc = []
        in_toc_block = False
        in_ul = False

        def close_ul():
            nonlocal in_ul
            if in_ul:
                body.append("</ul>")
                in_ul = False

        for kind, el in self.iter_body_items():
            if kind == "tbl":
                close_ul()
                body.append(render_table(el))
                continue

            p = el
            text = p.text.strip()
            style = p.style.name

            # skip the docx TOC block (we regenerate a real one with anchors)
            if text == "Contents":
                close_ul()
                in_toc_block = True
                continue
            if in_toc_block:
                if style.startswith("Heading"):
                    in_toc_block = False
                else:
                    continue

            if style.startswith("Heading 1") and text.lower().startswith("english"):
                continue  # part label, already represented in the cover

            if style.startswith("Heading"):
                close_ul()
                level = int(style.split()[-1])
                anchor = anchorify(text)
                # unique anchors
                seen = [t["anchor"] for t in toc]
                a, i = anchor, 2
                while a in seen:
                    a = f"{anchor}-{i}"
                    i += 1
                toc.append({"level": level, "text": text, "anchor": a})
                body.append(f'<h{level} id="{a}">{html.escape(text)}</h{level}>')
            elif style.startswith("List Bullet") or style.startswith("List Number"):
                if not in_ul:
                    body.append("<ul>")
                    in_ul = True
                body.append(f"<li>{runs_to_html(p, self)}</li>")
            elif text:
                close_ul()
                body.append(f"<p>{runs_to_html(p, self)}</p>")
            else:
                close_ul()
                frag = runs_to_html(p, self)  # paragraphs that hold only images
                if frag:
                    body.append(f"<p>{frag}</p>")
        close_ul()
        return "\n".join(body), toc


def render_table(tbl) -> str:
    rows = []
    for i, row in enumerate(tbl.rows):
        cells = []
        for c in row.cells:
            paras = [norm_ws(p.text) for p in c.paragraphs if p.text.strip()]
            cell_html = "<br>".join(html.escape(t) for t in paras) or "&nbsp;"
            cells.append(cell_html)
        tag = "th" if i == 0 else "td"
        rows.append("<tr>" + "".join(f"<{tag}>{x}</{tag}>" for x in cells) + "</tr>")
    n_cols = len(tbl.columns)
    return f'<div class="tablewrap"><table class="param"><colgroup>{"".join(f"<col>" for _ in range(n_cols))}</colgroup>{"".join(rows)}</table></div>'


PAGE_TMPL = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title} · Dragonfly Prototype Apps</title>
<meta name="description" content="{meta_desc}">
<link rel="stylesheet" href="../style.css">
</head>
<body>
<nav class="topnav">
  <a class="brand" href="../">Dragonfly <span>Prototype Apps</span></a>
  <span class="crumb">{title}</span>
</nav>
<div class="page">
  <header class="cover">
    <div class="overline">{subtitle}</div>
    <h1>{title}</h1>
    <div class="meta"><span class="badge">{version}</span><span class="badge badge-soft">Manual</span></div>
    {sources}
  </header>
  {toc}
  <main class="manual">{body}</main>
  <footer class="pagefoot">
    <p>Generated from <code>{sources_code}</code> · <a href="../">&larr; All manuals</a></p>
  </footer>
</div>
</body>
</html>
"""


STUB_TMPL = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title} (duplicate) · Dragonfly Prototype Apps</title>
<link rel="canonical" href="{canonical_url}">
<link rel="stylesheet" href="../style.css">
</head>
<body>
<nav class="topnav">
  <a class="brand" href="../">Dragonfly <span>Prototype Apps</span></a>
  <span class="crumb">{title}</span>
</nav>
<div class="page">
  <div class="stub">
    <h1>Duplicate manual</h1>
    <p><code>{fname}</code> contains content that is <strong>identical</strong> to the
    manual published as <a href="{canonical_url}">{canonical_title}</a>.</p>
    <p>To avoid two copies of the same text, only the canonical page is indexed in the
    search catalogue. <a href="{canonical_url}">Read it here &rarr;</a></p>
    <p><a href="../">&larr; All manuals</a></p>
  </div>
</div>
</body>
</html>
"""


def toc_html(toc: list[dict]) -> str:
    if not toc:
        return ""
    items = []
    for t in toc:
        cls = "t2" if t["level"] <= 2 else "t3"
        items.append(f'<li class="{cls}"><a href="#{t["anchor"]}">{html.escape(t["text"])}</a></li>')
    return f'<nav class="toc"><h2>Contents</h2><ul>{"".join(items)}</ul></nav>'


def build(src_dir: str, out_dir: str):
    files = sorted(
        glob.glob(os.path.join(src_dir, "**", "*.docx"), recursive=True)
        + glob.glob(os.path.join(src_dir, "*.docx")),
    )
    files = sorted(set(files))
    if not files:
        sys.exit(f"no .docx files found under {src_dir}")

    print(f"found {len(files)} .docx files under {src_dir}")

    # parse + dedupe by content hash
    manuals = []
    for f in files:
        try:
            m = Manual(f)
        except Exception as e:
            print(f"  ! skipping {os.path.basename(f)}: {e}")
            continue
        m.hash = extract_content_hash(m.doc)
        manuals.append(m)

    # stable canonical choice: if a page was already published for this hash
    # in a previous build, keep the same source file as canonical (so renaming
    # or adding a copy never flips the "real" page into a duplicate stub)
    old_canon = {}
    old_data = os.path.join(out_dir, "search-data.json")
    if os.path.exists(old_data):
        try:
            with open(old_data, encoding="utf-8") as fh:
                for c in json.load(fh):
                    if "hash" in c and "slug" in c:
                        old_canon[c["hash"]] = c["slug"]
        except Exception:
            pass

    groups: dict[str, list[Manual]] = {}
    for m in manuals:
        groups.setdefault(m.hash, []).append(m)
    for g in groups.values():
        g.sort(key=lambda m: m.path)
        prev = old_canon.get(g[0].hash)
        for m in g:
            if m.slug == prev:
                g.remove(m)
                g.insert(0, m)
                break

    os.makedirs(os.path.join(out_dir, "pages"), exist_ok=True)
    os.makedirs(os.path.join(out_dir, "assets"), exist_ok=True)

    cards = []
    dup_merged = 0
    for m in manuals:
        g = groups[m.hash]
        canonical = g[0]
        is_dup = m is not canonical
        dup_merged += 1 if is_dup else 0

        short, subtitle, version = m.cover()

        if is_dup:
            canonical_title = re.sub(r"\s*[-–—]\s*User Manual\s*$", "", canonical.cover()[0])
            page_html = STUB_TMPL.format(
                title=html.escape(short),
                fname=html.escape(m.fname),
                canonical_url=html.escape(f"{canonical.slug}.html"),
                canonical_title=html.escape(canonical_title),
            )
            with open(os.path.join(out_dir, "pages", f"{m.slug}.html"), "w", encoding="utf-8") as fh:
                fh.write(page_html)
            print(f"  = {m.slug}  (duplicate of {canonical.slug}, stub written)")
            continue

        cover_title = short
        m.set_out(out_dir)
        body_html, toc = m.render_body()
        sources_code = ", ".join(f"<code>{html.escape(x.fname)}</code>" for x in g)
        sources_html = ""
        if len(g) > 1:
            names = ", ".join(html.escape(x.fname) for x in g)
            sources_html = f'<p class="dupnote">Identical content also shipped in: {names}</p>'

        meta_desc = html.escape((m.plain_text()[:155] + "…") if len(m.plain_text()) > 158 else m.plain_text())
        page_html = PAGE_TMPL.format(
            title=html.escape(cover_title),
            meta_desc=meta_desc,
            subtitle=html.escape(subtitle),
            version=html.escape(version),
            sources=sources_html,
            toc=toc_html(toc),
            body=body_html,
            sources_code=sources_code,
        )
        page_path = os.path.join(out_dir, "pages", f"{m.slug}.html")
        with open(page_path, "w", encoding="utf-8") as fh:
            fh.write(page_html)

        if not is_dup:
            desc = m.plain_text()[:200]
            cards.append({
                "slug": m.slug,
                "hash": m.hash,
                "file": f"pages/{m.slug}.html",
                "title": short,
                "version": version,
                "desc": desc,
                "aliases": [x.fname for x in g[1:]],
                "text": m.plain_text(),
            })
            print(f"  + {m.slug}  ({m.fname}{', = ' + str(len(g) - 1) + ' dup(s)' if len(g) > 1 else ''})")

    cards.sort(key=lambda c: c["title"].lower())

    # remove pages/assets for manuals that no longer exist in the source dir
    wanted = {m.slug for m in manuals}
    pages_dir = os.path.join(out_dir, "pages")
    for f in os.listdir(pages_dir):
        if f.endswith(".html") and f[:-5] not in wanted:
            os.remove(os.path.join(pages_dir, f))
            print(f"  - removed stale page {f}")
    assets_dir = os.path.join(out_dir, "assets")
    if os.path.isdir(assets_dir):
        for d in os.listdir(assets_dir):
            if d not in wanted:
                import shutil
                shutil.rmtree(os.path.join(assets_dir, d))
                print(f"  - removed stale assets/{d}")

    # search-data.json: full payloads (no need to hide — public site)
    with open(os.path.join(out_dir, "search-data.json"), "w", encoding="utf-8") as fh:
        json.dump(cards, fh, ensure_ascii=False, separators=(",", ":"))

    # index.html
    index = INDEX_TMPL.format(
        count=len(cards),
        cards_json=json.dumps(cards, ensure_ascii=False, separators=(",", ":")),
        year=datetime_date(),
    )
    with open(os.path.join(out_dir, "index.html"), "w", encoding="utf-8") as fh:
        fh.write(index)

    with open(os.path.join(out_dir, "style.css"), "w", encoding="utf-8") as fh:
        fh.write(CSS)

    write_readme(out_dir, src_dir, len(cards), dup_merged)
    merged_files = len(manuals) - len(cards)
    print(f"done: {len(cards)} unique manuals from {len(manuals)} files, {merged_files} duplicate(s) merged")
    print(f"site written to {out_dir}")


def datetime_date() -> str:
    import datetime
    return datetime.date.today().isoformat()


def write_readme(out_dir, src_dir, n_pages, merged):
    readme = f"""# Dragonfly Prototype Apps — User Manuals

A searchable catalogue of {n_pages} plugin user manuals, generated from the
`.docx` sources in `{src_dir}`.

## Layout

| path | what |
|---|---|
| `index.html` | homepage with search (client-side, full-text) |
| `pages/` | one HTML page per manual |
| `assets/` | images extracted from the manuals |
| `search-data.json` | search payloads (titles, descriptions, full text) |
| `build/build_site.py` | the generator |

## Adding / updating manuals

1. Put new or updated `.docx` files into `{src_dir}` (any subfolder works).
2. Run the generator:

   ```bash
   git clone git@github.com:howpy/imagefly_app.git
   cd imagefly_app
   pip install python-docx
   python3 build/build_site.py --src {src_dir}
   ```

3. `git add . && git commit && git push` — GitHub Pages rebuilds automatically.

The generator is idempotent: files whose content is **identical** are deduped
into one page (duplicates are listed as aliases on the page and homepage);
changed files are rebuilt in place, so slugs stay stable.
"""
    with open(os.path.join(out_dir, "README.md"), "w", encoding="utf-8") as fh:
        fh.write(readme)


CSS = """
:root{
  --bg:#fafaf9; --panel:#ffffff; --ink:#1c1917; --muted:#78716c;
  --line:#e7e5e4; --accent:#0f766e; --accent-soft:#ccfbf1;
  --code-bg:#f5f5f4; --mark:#fde68a;
}
@media (prefers-color-scheme: dark){
  :root{
    --bg:#0c0a09; --panel:#161311; --ink:#f5f5f4; --muted:#a8a29e;
    --line:#292524; --accent:#2dd4bf; --accent-soft:#042f2e;
    --code-bg:#1c1917; --mark:#854d0e;
  }
}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;background:var(--bg);color:var(--ink);
  font:16px/1.65 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"Helvetica Neue",Arial,"PingFang SC","Microsoft YaHei",sans-serif;
  -webkit-font-smoothing:antialiased}
a{color:var(--accent);text-decoration:none}
a:hover{text-decoration:underline}
code{background:var(--code-bg);border:1px solid var(--line);border-radius:4px;
  padding:.08em .35em;font-size:.88em;font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace}
mark{background:var(--mark);color:inherit;border-radius:2px;padding:0 1px}

.topnav{position:sticky;top:0;z-index:10;display:flex;align-items:center;gap:14px;
  padding:12px clamp(16px,4vw,40px);background:color-mix(in srgb,var(--panel) 88%,transparent);
  backdrop-filter:blur(10px);border-bottom:1px solid var(--line)}
.topnav .brand{font-weight:700;font-size:15px;color:var(--ink)}
.topnav .brand span{color:var(--accent)}
.topnav .crumb{color:var(--muted);font-size:14px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}

/* ---------- index ---------- */
.index{max-width:1200px;margin:0 auto;padding:clamp(20px,4vw,44px) clamp(16px,4vw,40px)}
.index h1{font-size:clamp(26px,4vw,38px);margin:0 0 6px}
.index .sub{color:var(--muted);margin:0 0 22px}
.searchbox{position:relative;margin-bottom:10px}
.searchbox input{width:100%;padding:14px 44px 14px 44px;border:1px solid var(--line);border-radius:12px;
  background:var(--panel);color:var(--ink);font-size:16px;outline:none}
.searchbox input:focus{border-color:var(--accent)}
.searchbox .icon{position:absolute;left:15px;top:50%;transform:translateY(-50%);color:var(--muted)}
.searchbox kbd{position:absolute;right:12px;top:50%;transform:translateY(-50%)}
.resultcount{color:var(--muted);font-size:14px;margin:0 0 16px}
kbd{background:var(--code-bg);border:1px solid var(--line);border-bottom-width:2px;border-radius:5px;
  padding:1px 6px;font-size:11px;font-family:ui-monospace,monospace;color:var(--muted)}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:14px}
.card{display:block;background:var(--panel);border:1px solid var(--line);border-radius:12px;
  padding:18px;color:var(--ink);transition:border-color .12s,transform .12s}
.card:hover{border-color:var(--accent);text-decoration:none;transform:translateY(-2px)}
.card h3{margin:0 0 6px;font-size:16.5px;line-height:1.35}
.card .ver{color:var(--accent);font-size:12px;font-weight:600;margin-bottom:8px;display:block}
.card p{margin:0;color:var(--muted);font-size:13.5px;line-height:1.55;
  display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden}
.card .snip{display:none;margin-top:8px;font-size:13px}
.card.hit .snip{display:block}
.card .aliases{margin-top:8px;font-size:11.5px;color:var(--muted)}
.index footer{text-align:center;color:var(--muted);font-size:13px;margin-top:44px}
.empty{color:var(--muted);padding:40px;text-align:center}

/* ---------- manual pages ---------- */
.page{max-width:900px;margin:0 auto;padding:28px clamp(16px,4vw,40px) 60px}
.cover{border-bottom:1px solid var(--line);padding-bottom:24px;margin-bottom:28px}
.cover .overline{color:var(--accent);font-weight:600;font-size:13px;letter-spacing:.04em;text-transform:uppercase;margin-bottom:8px}
.cover h1{font-size:clamp(24px,4vw,34px);margin:0 0 12px;line-height:1.2}
.cover .meta{display:flex;gap:8px;flex-wrap:wrap}
.badge{background:var(--accent-soft);color:var(--accent);font-size:12px;font-weight:600;
  padding:3px 10px;border-radius:999px}
.badge-soft{background:var(--code-bg);color:var(--muted)}
.dupnote{margin-top:14px;font-size:13px;color:var(--muted)}
.toc{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:18px 22px;margin:0 0 34px}
.toc h2{font-size:13px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin:0 0 10px}
.toc ul{list-style:none;margin:0;padding:0}
.toc li{margin:2px 0}
.toc li.t3 a{padding-left:20px;color:var(--muted)}
.toc a{color:var(--ink);font-size:14.5px}
.toc a:hover{color:var(--accent);text-decoration:none}

.manual h2{font-size:23px;margin:44px 0 14px;padding-top:10px;border-top:1px solid var(--line)}
.manual h3{font-size:18px;margin:28px 0 10px}
.manual h4{font-size:15.5px;margin:22px 0 8px}
.manual p{margin:0 0 14px}
.manual ul{margin:0 0 16px;padding-left:24px}
.manual li{margin:4px 0}
.manual img{max-width:100%;border:1px solid var(--line);border-radius:8px}
.manual figure{margin:18px 0}
.tablewrap{overflow-x:auto;margin:0 0 20px;border:1px solid var(--line);border-radius:10px}
table.param{border-collapse:collapse;width:100%;font-size:13.5px;background:var(--panel)}
table.param th,table.param td{border-bottom:1px solid var(--line);padding:8px 12px;text-align:left;vertical-align:top}
table.param th{background:var(--code-bg);font-weight:600;white-space:nowrap}
table.param tr:last-child td{border-bottom:none}
.pagefoot{margin-top:52px;padding-top:18px;border-top:1px solid var(--line);color:var(--muted);font-size:13.5px}
.stub{margin:56px 0}
.stub h1{font-size:24px;margin:0 0 18px}
.stub p{margin:0 0 14px}
"""


INDEX_TMPL = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dragonfly Prototype Apps — User Manuals</title>
<meta name="description" content="Searchable catalogue of Dragonfly prototype app plugin user manuals.">
<link rel="stylesheet" href="style.css">
</head>
<body>
<nav class="topnav">
  <a class="brand" href="./">Dragonfly <span>Prototype Apps</span></a>
  <span class="crumb">User manuals</span>
</nav>
<div class="index">
  <h1>User Manuals</h1>
  <p class="sub">{count} Dragonfly prototype apps · full-text search over every manual</p>
  <div class="searchbox">
    <span class="icon">&#128269;</span>
    <input id="q" type="search" placeholder="Search manuals… (e.g. denoise, porosity, EBSD, ROI)" autocomplete="off">
    <kbd>/</kbd>
  </div>
  <p class="resultcount" id="count"></p>
  <div class="grid" id="grid"></div>
  <div class="empty" id="empty" hidden>No manuals match — try a different term.</div>
  <footer>Generated {year} from the plugin manual sources.</footer>
</div>
<script>
const CARDS = {cards_json};
const grid = document.getElementById('grid');
const q = document.getElementById('q');
const countEl = document.getElementById('count');
const emptyEl = document.getElementById('empty');

function escapeRe(s){{ return s.replace(/[.*+?^${{()\\[\\]}}]/g, '\\\\$&'); }}

function snippet(text, terms){{
  for (const t of terms){{
    const i = text.toLowerCase().indexOf(t.toLowerCase());
    if (i >= 0){{
      const start = Math.max(0, i - 60);
      const end = Math.min(text.length, i + t.length + 110);
      let s = (start > 0 ? '…' : '') + text.slice(start, end) + (end < text.length ? '…' : '');
      const re = new RegExp('(' + terms.map(escapeRe).join('|') + ')', 'gi');
      s = s.replace(re, '<mark>$1</mark>');
      return s;
    }}
  }}
  return null;
}}

function render(){{
  const query = q.value.trim();
  const terms = query ? query.split(/\\s+/).filter(Boolean) : [];
  let n = 0;
  for (const c of CARDS){{
    const hay = (c.title + ' ' + c.version + ' ' + c.text).toLowerCase();
    const hit = terms.every(t => hay.includes(t.toLowerCase()));
    let el = document.getElementById('card-' + c.slug);
    if (!el){{
      el = document.createElement('a');
      el.className = 'card';
      el.id = 'card-' + c.slug;
      el.href = c.file;
      el.innerHTML =
        '<h3></h3><span class="ver"></span><p class="desc"></p><p class="snip"></p>' +
        (c.aliases.length ? '<div class="aliases">also: ' + c.aliases.join(', ') + '</div>' : '');
      el.querySelector('h3').textContent = c.title;
      el.querySelector('.ver').textContent = c.version;
      el.querySelector('.desc').textContent = c.desc;
      grid.appendChild(el);
    }}
    if (hit){{
      n++;
      el.style.display = '';
      el.classList.toggle('hit', !!terms.length);
      const snip = el.querySelector('.snip');
      snip.innerHTML = terms.length ? (snippet(c.text, terms) || '') : '';
    }} else {{
      el.style.display = 'none';
    }}
  }}
  countEl.textContent = terms.length ? (n + ' of ' + CARDS.length + ' manuals') : (CARDS.length + ' manuals');
  emptyEl.hidden = n !== 0;
}}

q.addEventListener('input', render);
document.addEventListener('keydown', e => {{
  if (e.key === '/' && document.activeElement !== q){{ e.preventDefault(); q.focus(); }}
  if (e.key === 'Escape'){{ q.value = ''; render(); q.blur(); }}
}});
render();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--src", default="/Users/user/DragonflyPlugins", help="directory containing .docx manuals")
    ap.add_argument("--out", default=".", help="output directory for the generated site")
    args = ap.parse_args()
    build(args.src, args.out)
