#!/usr/bin/env python3
"""Post-build checks on dcr-calculator.html. Run by CI, useful locally too.

Guards the four ways this build can fail quietly:
  1. a payload placeholder never got substituted;
  2. a non-ASCII character leaked in, which mojibakes m2 because the page is
     served inside a wrapper whose charset we do not control;
  3. plot-specific client data finding its way back into the build;
  4. client data files being *tracked by git*. That is how a geodatabase once
     reached a public remote: the commit was about something else and a
     `git add -A` swept the directory in. Checking only the built HTML could
     never have caught it, because the data was never in the HTML.
"""
import re
import subprocess
import sys

# Patterns, not literals: naming the plot here would put the very data we are
# purging back into the repository. These also catch *any* plot, not just one.
PLOT_DATA = (
    (r"\b\d{2}-\d-\d{3}-\d{3}\b",        "plot number"),
    (r"\b3[0-9]{5}\.\d{2,3}\b",           "UTM easting"),
    (r"\b2[0-9]{6}\.\d{2,3}\b",           "UTM northing"),
)
PLACEHOLDERS = ("__WASM_B64__", "__DB_B64__", "__SNAPSHOT__", "/*__SQLJS__*/")

# Paths that must never be tracked. itc_rates.db is deliberately absent from this
# list: the plot list it carries is published on purpose. Geometry, GIS ids and
# the source workbooks are not.
FORBIDDEN_TRACKED = (
    (r"(^|/)data/",         "the data/ directory"),
    (r"\.gdb/",             "a file geodatabase"),
    (r"\.(xls|xlsx|pdf)$",  "a source workbook or report"),
)


def tracked_client_data():
    """Tracked paths that look like client data. Empty if git is unavailable."""
    try:
        out = subprocess.run(["git", "ls-files"], capture_output=True, text=True,
                             timeout=30, check=True).stdout
    except (OSError, subprocess.SubprocessError):
        return []
    bad = []
    for path in out.splitlines():
        for pattern, label in FORBIDDEN_TRACKED:
            if re.search(pattern, path):
                bad.append((path, label))
                break
    return bad

def main():
    try:
        html = open("dcr-calculator.html", encoding="utf-8").read()
    except FileNotFoundError:
        sys.exit("dcr-calculator.html not found - run `python3 embed_db.py` first")

    errors = []
    for token in PLACEHOLDERS:
        if token in html:
            errors.append(f"placeholder {token} was not substituted")
    if len(html) < 1_000_000:
        errors.append(f"output only {len(html):,} bytes - a payload is missing")
    non_ascii = [(i, c) for i, c in enumerate(html) if ord(c) > 127]
    if non_ascii:
        i, c = non_ascii[0]
        errors.append(f"{len(non_ascii)} non-ASCII chars, first {c!r} at offset {i}")
    for pattern, label in PLOT_DATA:
        hits = re.findall(pattern, html)
        if hits:
            errors.append(f"looks like {label} in build: {sorted(set(hits))[:3]}")

    tracked = tracked_client_data()
    if tracked:
        kinds = ", ".join(sorted({label for _, label in tracked}))
        errors.append(f"{len(tracked)} client-data files tracked by git ({kinds}); "
                      f"first is {tracked[0][0]}")

    if errors:
        for e in errors:
            print("FAIL:", e, file=sys.stderr)
        sys.exit(1)
    print(f"ok: {len(html):,} bytes, pure ASCII, no plot-specific data, "
          f"nothing client-side tracked")

if __name__ == "__main__":
    main()
