# Vendored dependencies

## sql.js 1.13.0

SQLite compiled to WebAssembly. Used to query `itc_rates.db` inside the page.

- Upstream: https://github.com/sql-js/sql.js
- Files: `sql-wasm.js` (loader), `sql-wasm.wasm` (engine)
- Licence: MIT (sql.js); SQLite itself is public domain

Vendored rather than loaded from a CDN because the page must work with no outbound
requests — `embed_db.py` inlines both files as base64 at build time. To upgrade, replace
both files with a matching pair from the same release and re-run `python3 embed_db.py`.
