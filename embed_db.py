#!/usr/bin/env python3
"""Build dcr-calculator.html from dcr-calculator.src.html.

Inlines the sql.js loader, the SQLite WASM binary and itc_rates.db as base64 so
the published page is fully self-contained -- an Artifact CSP blocks every
external host, so nothing may be fetched at runtime.

    python3 build_db.py     # xlsx  -> itc_rates.db
    python3 embed_db.py     # src + db -> dcr-calculator.html

Output is pure ASCII: the page is served inside a wrapper whose charset we do
not control, so every non-ASCII character is escaped per context (HTML entity,
CSS escape or JS \\u escape).
"""
import base64
import json
import os
import re
import sqlite3
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "dcr-calculator.src.html")
OUT = os.path.join(HERE, "dcr-calculator.html")
DB = os.path.join(HERE, "itc_rates.db")
SQLJS = os.path.join(HERE, "vendor", "sql-wasm.js")
WASM = os.path.join(HERE, "vendor", "sql-wasm.wasm")


def escape_by_context(html):
    """Escape non-ASCII per region: CSS inside <style>, \\u inside <script>,
    HTML entities everywhere else."""
    spans = []          # (start, end, mode) for style/script bodies
    for tag, mode in (("style", "css"), ("script", "js")):
        for m in re.finditer(rf"<{tag}>(.*?)</{tag}>", html, re.S):
            spans.append((m.start(1), m.end(1), mode))
    spans.sort()

    def esc(text, mode):
        out = []
        for ch in text:
            if ord(ch) < 128:
                out.append(ch)
            elif mode == "css":
                out.append("\\%04X " % ord(ch))
            elif mode == "js":
                out.append("\\u%04X" % ord(ch))
            else:
                out.append("&#x%04X;" % ord(ch))
        return "".join(out)

    parts, pos = [], 0
    for start, end, mode in spans:
        parts.append(esc(html[pos:start], "html"))
        parts.append(esc(html[start:end], mode))
        pos = end
    parts.append(esc(html[pos:], "html"))
    return "".join(parts)


def snapshot(db_path):
    """Export the tables the page needs as JSON.

    A published page runs under a CSP we do not control, and WebAssembly needs
    'wasm-unsafe-eval' in script-src. If that is absent, sql.js cannot start --
    so the page also carries this snapshot of the same database and falls back
    to it, keeping the calculator usable. One source, two encodings.
    """
    con = sqlite3.connect(db_path)
    con.row_factory = sqlite3.Row
    rate_cols = ("id, period, class_code, class_name, class_group, code, variant, "
                 "sites, trip_am, trip_noon, trip_pm, trip_phg, "
                 "rate_employee, rate_visitor, rate_truck, rate_total, "
                 "unit_label, driver_label, driver_kind, conversion")
    data = {
        "itc_rate": [dict(r) for r in
                     con.execute(f"SELECT {rate_cols} FROM itc_rate ORDER BY id")],
        "land_use": [dict(r) for r in con.execute(
            "SELECT id, category, designation, name, far, coverage, remarks, source "
            "FROM land_use ORDER BY id")],
        "plot": [dict(r) for r in con.execute(
            "SELECT id, plot_number, sector_plot_id, district, primary_use, "
            "devcode, devcode_cat, area FROM plot ORDER BY sector_plot_id")],
        # 2,000-odd rows across 20 plots; the picker reads only the selected
        # plot's, but the snapshot has to carry them all.
        "plot_approved": [dict(r) for r in con.execute(
            "SELECT sector_plot_id, region, itc_loc, district, dev_code, planned_area, "
            "area, gfa, far, coverage_pct FROM plot_approved ORDER BY area")],
        "uad_map": [dict(r) for r in con.execute(
            "SELECT uad_code, uad_category, uad_category_inspection, itc_equivalent, "
            "itc_class, note FROM uad_map ORDER BY uad_category")],
        "plot_uad": [dict(r) for r in con.execute(
            "SELECT sector_plot_id, activity_code, activity_name, variant_en, variant_ar, "
            "licence_type, luc_code, category, transport_mode, found_walk_700m, "
            "found_drive_10m, inclusion, proposed, rational, identifier, identifier_logic "
            "FROM plot_uad ORDER BY sector_plot_id, id")],
        "weighting": [dict(r) for r in
                      con.execute("SELECT period, days, share FROM weighting")],
        # 3,900-odd rows; the picker searches these when SQLite is unavailable
        "ded_activity": [dict(r) for r in con.execute(
            "SELECT activity_id, name, division, isic_class, itc_class, map_confidence, "
            "map_reason FROM ded_activity ORDER BY activity_id")],
        "coefficient": {r["key"]: r["value"] for r in
                        con.execute("SELECT key, value FROM coefficient")},
        "meta": {r["key"]: r["value"] for r in con.execute("SELECT key, value FROM meta")},
    }
    con.close()
    return json.dumps(data, separators=(",", ":"), ensure_ascii=True)


def main():
    for path in (SRC, DB, SQLJS, WASM):
        if not os.path.exists(path):
            sys.exit(f"missing: {path}")

    html = open(SRC, encoding="utf-8").read()
    sqljs = open(SQLJS, encoding="utf-8").read()
    wasm_b64 = base64.b64encode(open(WASM, "rb").read()).decode("ascii")
    db_b64 = base64.b64encode(open(DB, "rb").read()).decode("ascii")

    if "</script" in sqljs.lower():
        sys.exit("sql-wasm.js contains a closing script tag; cannot inline as-is")

    # Escape the authored page first, then splice in the already-ASCII payloads
    # so ~1 MB of base64 is not walked character by character.
    html = escape_by_context(html)
    snap = snapshot(DB)
    for token, value, label in (("/*__SQLJS__*/", sqljs, "sql.js loader"),
                                ("__WASM_B64__", wasm_b64, "wasm binary"),
                                ("__DB_B64__", db_b64, "sqlite database"),
                                ("__SNAPSHOT__", snap, "fallback snapshot")):
        if token not in html:
            sys.exit(f"placeholder {token} not found in source")
        html = html.replace(token, value)
        print(f"  inlined {label:16s} {len(value):>9,d} chars")

    open(OUT, "w", encoding="utf-8").write(html)
    non_ascii = sum(1 for c in html if ord(c) > 127)
    print(f"{OUT}")
    print(f"  {len(html.encode()):,d} bytes  ·  non-ASCII chars: {non_ascii}")
    if non_ascii:
        sys.exit("output is not pure ASCII")


if __name__ == "__main__":
    main()
