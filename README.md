# Dragonfly Prototype Apps — User Manuals

A searchable catalogue of 97 plugin user manuals, generated from the
`.docx` sources in `/Users/user/DragonflyPlugins`.

## Layout

| path | what |
|---|---|
| `index.html` | homepage with search (client-side, full-text) |
| `pages/` | one HTML page per manual |
| `assets/` | images extracted from the manuals |
| `search-data.json` | search payloads (titles, descriptions, full text) |
| `build/build_site.py` | the generator |

## Adding / updating manuals

1. Put new or updated `.docx` files into `/Users/user/DragonflyPlugins` (any subfolder works).
2. Run the generator:

   ```bash
   git clone git@github.com:howpy/imagefly_app.git
   cd imagefly_app
   pip install python-docx
   python3 build/build_site.py --src /Users/user/DragonflyPlugins
   ```

3. `git add . && git commit && git push` — GitHub Pages rebuilds automatically.

The generator is idempotent: files whose content is **identical** are deduped
into one page (duplicates are listed as aliases on the page and homepage);
changed files are rebuilt in place, so slugs stay stable.
