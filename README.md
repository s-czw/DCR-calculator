# DCR Plot Yield Calculator

Derives Max GFA, plot coverage, GLA and the permitted activity count from
a plot area, a land-use designation and an ITC category. The designation schedule and the
ITC rate matrix are both embedded as a SQLite database and queried live in the page.

## Files

| File | Role |
|---|---|
| `dcr-calculator.html` | **The deliverable.** Self-contained, ~1.1 MB. Open directly or publish. |
| `dcr-calculator.src.html` | Source. Edit this, never the built file. |
| `itc_rates.db` | SQLite: 282 ITC rate rows (141 codes × 2 periods), 59 land-use designations, the week weighting and coefficients. |
| `build_db.py` | ITC sheet + `Designation_Index.xlsx` &rarr; `itc_rates.db` |
| `embed_db.py` | `dcr-calculator.src.html` + db + sql.js &rarr; `dcr-calculator.html` |
| `gdb_export.py` | Runs the derivation over a file geodatabase's plot layer, to Excel. |
| `verify_build.py` | Post-build checks: placeholders substituted, output pure ASCII, no plot data. |
| `vendor/sql-wasm.{js,wasm}` | sql.js 1.13.0 — SQLite compiled to WebAssembly. See `vendor/README.md`. |
| `.github/workflows/build.yml` | CI: build, verify, then deploy to whichever target `DEPLOY_TARGET` selects. |

## Rebuilding

When the matrix is reissued:

```bash
python3 build_db.py     # workbooks -> itc_rates.db   (needs xlrd + openpyxl)
python3 embed_db.py     # src + db  -> dcr-calculator.html
python3 verify_build.py # sanity-check the output
```

`build_db.py` looks for the workbooks in `./sources`, or wherever `DCR_SOURCE_DIR` points,
or at two paths you pass as arguments. **The workbooks are client data and are not in this
repository** — `sources/` is gitignored.

`dcr-calculator.html` is generated and gitignored too: clone and run `embed_db.py`, or take
it from a CI run's artifacts. `itc_rates.db` *is* committed, which is what keeps `embed_db.py`
dependency-free and lets CI install nothing.

`build_db.py` takes the ITC sheet as the first optional argument and the Designation Index
as the second, if either moves. Both scripts are
idempotent; `build_db.py` drops and recreates every table. The current source is a legacy
`.xls`, which needs **xlrd** (`pip install xlrd`); `.xlsx` still works through openpyxl.
Both tab spellings are accepted (`Rates-Weekend` / `Rates-Weekends`).

Class and Class Name are **merged cells** spanning each class's variant rows in the `.xls`,
so the parser carries them forward. Requiring a value in the Class column instead silently
drops all 71 variant rows, leaving 70 of 141 codes — worth remembering on the next reissue.

The build inlines sql.js, the WASM binary, the database and a JSON snapshot as base64 or
literals, and escapes all non-ASCII per context (HTML entity / CSS escape / JS `\u`). All
of it is required: a published Artifact runs under a CSP that blocks every external host,
and inside a document wrapper whose charset declaration we do not control.

### Why there is a snapshot as well as a database

SQLite in the browser is sql.js — SQLite compiled to WebAssembly — and WebAssembly needs
`wasm-unsafe-eval` in the host's `script-src`. That could not be verified for the Artifact
host, so the page carries both: it opens the real database and queries it, and if
WebAssembly is blocked it falls back to a JSON snapshot of the same tables exported by
`embed_db.py` at build time. Same source, two encodings — figures are identical either
way. If the fallback ever engages the only trace is a `console.warn` — a masthead chip used to
report which path ran, but the distinction does not affect any figure, so it was not worth the
space.

## Deploying

CI builds and verifies on every push and pull request, and uploads the page as a run
artifact. Deployment is **off until you opt in**, so nothing fails while it is unconfigured.

Pick a target by setting a repository variable — *Settings → Secrets and variables → Actions
→ Variables* — named `DEPLOY_TARGET`:

| `DEPLOY_TARGET` | Also needs | Result |
|---|---|---|
| unset | — | build + verify only; download the page from the run summary |
| `azure` | secret `AZURE_STATIC_WEB_APPS_TOKEN` | Azure Static Web Apps; repo stays private, Entra ID can restrict access to invited people |
| `pages` | *Settings → Pages → Source: GitHub Actions* | GitHub Pages |

### A caveat on GitHub Pages

This repository is private. Pages on a private repository requires a paid plan, and the
published site is **public** regardless — the repository stays private, the site does not.
Access control for Pages exists only on GitHub Enterprise Cloud. So Pages is the wrong
choice if the point is to limit the audience; use `azure` (or put Cloudflare Access in front
of a static host) for that.

The page carries no secrets, and no plot-specific data — `verify_build.py` enforces that. But
it does hold the land-use designation schedule and the full ITC rate matrix, and several
figures are not yet confirmed with DMT. Read *Confirmed vs. assumed* before making it
reachable by anyone who might quote it.

### Adding another target

`build` produces `dist/` and uploads it as the `dcr-calculator` artifact. Any new deploy job
downloads that artifact and ships it — for an internal box, that is an rsync over SSH:

```yaml
  deploy-ssh:
    needs: build
    if: vars.DEPLOY_TARGET == 'ssh' && github.ref == 'refs/heads/main'
    runs-on: ubuntu-latest
    steps:
      - uses: actions/download-artifact@v4
        with: { name: dcr-calculator, path: dist }
      - run: |
          install -m 600 /dev/stdin key <<< "${{ secrets.DEPLOY_KEY }}"
          rsync -az --delete -e "ssh -i key -o StrictHostKeyChecking=accept-new" \
            dist/ "${{ secrets.DEPLOY_TARGET_HOST }}"
```

## Override values

No field is pre-filled from a table. The original shows as placeholder text instead —
`Original value: 75%` — and the calculation uses it while the field is empty. Typing in a field
does three things:

1. an amber strip above the columns lists what changed, `GLA share 75% → 80%`, with a masthead
   chip counting them and a Clear button;
2. every figure that depended on it gains a line naming the override,
   `Recalculated using GLA share = 80% (original 75%)`;
3. clearing the field restores the original and the indicators disappear.

Overridable: FAR, max plot coverage, GLA share, unit area per activity, area per parking space,
usable share of a basement floor, basement floor area, and the required number of spaces — the last
for a UPPC or ITC reduction, defaulting to the count the activity schedule produces.

Every overridable field shows its resolved value directly beneath itself, with where that value
came from: `Coverage 90%  overridden`, `Space 32.5 m²  standard`, `Spaces 26  from the schedule`.
The readout belongs to one field rather than trailing a group of them. The collapsible block
holding the non-Code parameters is titled **Parameters**. The three percentage fields
(coverage, GLA share, usable per basement floor) accept **0–100 with at most two decimals**:
letters, signs and exponents are refused, a third decimal is truncated, and anything above 100
or below 0 is clamped as you type.

The point is that an entered value always reads as a deviation, so a figure derived from Code
values never looks like one derived from a hand-entered assumption.


## The sign-in gate

One account: username `dmtdcr`. The password is not in this repository — it is compared as a
salted SHA-256 digest, and only the digest is committed.

**This is not a security control, and should not be relied on as one.** Say so to anyone who asks.
The page is a single static file: the plot register, the ITC matrix and every rate travel inside
it, so anyone who can load the file already has the data, whatever the gate does. Deleting the
overlay from the DOM, or reading the source, gets straight past it. A known password can be
brute-forced against a published digest in seconds.

What it *is* good for: keeping the tool out of the way of people who have no business in it, and
making it obvious that it is not for general use. If the contents ever genuinely need protecting,
that has to happen at the server — private hosting with real authentication in front of it, which
is what the `azure` deploy target in the workflow is there for.

The digest is used instead of a literal for one specific reason: a plaintext password committed to
a public repository is exposed permanently, is indexed, and stays in the history after it is
changed — and people reuse passwords. Hashing keeps the literal out of the history. It does not
make the gate strong.

Mechanics: a `locked` class goes on the root element in the head script, before the stylesheet
loads, so the app is never painted and then covered. The state lives in `sessionStorage`, so it
survives a reload but not a new tab, and **Sign out** in the masthead clears it. A failed attempt
pauses briefly, which costs a person nothing and makes scripted guessing through the form tedious.

## Visual design

Two palettes. **Light is the default** — a first visit is light whatever the operating system
is set to, because `prefers-color-scheme` is deliberately not consulted. The masthead carries a
Light/Dark segmented control; the choice is remembered in `localStorage` under `dcr.theme` and
re-applied by a small inline script in the head, ahead of the stylesheet, so a stored dark
choice does not flash light first.

Dark is the project Figma palette, applied deliberately rather than derived:

| Token | Value | Use |
|---|---|---|
| page | `#1C1C1C` | ground |
| panel | `#242425` | surfaces |
| white | `#FFFFFF` | headings, values, table text |
| grey | `#7D7D81` | labels, hints, sub-lines, units |
| blue | `#4786C9` | Max GFA and GLA figures, links, primary actions |
| yellow | `#DC9530` | the activity count, and warnings |

Light inverts the depth relationship — cards sit *above* the page rather than below it, so the
surface is white and the page a light grey:

| Token | Value | Use |
|---|---|---|
| page | `#F4F4F5` | ground |
| panel | `#FFFFFF` | surfaces |
| ink | `#1C1C1C` | headings, values, table text |
| grey | `#5C5C61` | labels, hints, sub-lines, units |
| blue | `#2C6098` | figures and links at text size; `#4786C9` for large display and borders |
| yellow | `#8A5A0C` | the activity count, and warnings |
| border | `#8D8D91` | field borders and structural rules |

Type is **Manrope** throughout in both themes, with IBM Plex Mono kept for formula chips and
small numeric annotations. Every colour is painted explicitly, so neither theme inherits
anything from the viewer.

### Contrast, on the record

**Light mode clears WCAG AA in full** — an in-browser audit over every element with its own
text found 0 failures out of 203, checked against the effective background and the real font
size, at 4.5:1 for normal text and 3.0:1 for large. Two token splits are what make that work,
and both are load-bearing:

- `--accent` (`#4786C9`) is only 3.80:1 on white, so it is confined to borders and the large
  display figures, where 3.0 applies. Text-size accent uses `--accent-ink` (`#2C6098`, 6.49:1).
- `--accent-solid` (`#2C6098` in light) is the fill behind white button labels. On `#4786C9`
  white is 3.80:1; on `#2C6098` it is 6.49:1.

`--line-strong` is `#8D8D91` rather than a hairline because it borders form fields, which WCAG
1.4.11 holds to 3:1 — it clears it at 3.31:1 on a card and 3.01:1 on the page.

**Dark mode does not, and is unchanged.** Two of its four text colours fall short on `#242425`:

| Colour | Ratio | Status |
|---|---|---|
| `#FFFFFF` | 15.51 | passes |
| `#DC9530` | 6.19 | passes |
| `#4786C9` | 4.08 | large text only (AA needs 4.5 for normal, 3.0 for large) |
| `#7D7D81` | 3.78 | large text only |

The same audit run in dark mode reports 158 failures on those two colours. That affects labels,
hints, sub-lines, table headers, chips and the status badges. The large figures pass, because at
weight 700 and 18.66 px or more they only need 3.0.

**This is a deliberate design decision, not an oversight.** Do not "fix" it without checking
first. If it is ever revisited, the minimum change that clears AA is `#7D7D81` -> `#8A8A8E`
and `#4786C9` -> `#4F8ED1`, both visually near-identical, plus pointing `--accent-solid` at a
deeper blue in dark as it already is in light.

The one dark-mode contrast fix that *was* made is the theme control itself, since it is new
rather than specified: its active option is `--ink` on the accent tint (11.77:1 dark, 14.79:1
light) rather than the accent blue, which would have been 3.09:1 in dark. The inactive option
stays on `--ink-3`, matching the chips beside it.

## The plot register, and saving a configuration

`plot` in the database holds the 66 plots from `DCR_plots.xlsx`, keyed by **sector/plot ID**
(`19_1_106_118_A`). The Inputs panel leads with a **Plot** picker: choosing a plot fills the area
from the register, and typing an area by hand returns the picker to *Manual entry* — a figure is
never left filed against a plot it does not belong to.

Picking a plot deliberately does **not** change the designation. The register carries a DevCode,
but FAR and coverage follow the land-use selection, and moving that silently would change the
derivation behind you. The DevCode is shown in each option so it can be matched by hand.

### Save

**Save** writes one JSON document keyed by sector/plot address, so re-saving a plot replaces its
entry rather than appending a second one:

```json
{
  "19_1_106_118_A": {
    "savedAt": "...", "plot": { … register record … },
    "inputs": { "plotArea": 401.72, "designation": "NR", "raw": { … every field … } },
    "overrides": [ … ], "derived": { … }, "activities": [ … ], "parking": { … }
  }
}
```

`inputs.raw` is every form field verbatim, so an entry is enough to reconstruct the exact state,
while the sibling blocks are readable without knowing the field ids. `activities` records each
row's ITC code, rate, what it is charged per, its quantity where charged per unit, and its
unrounded bay contribution.

Saving needs a plot: with *Manual entry* selected there is no address to key on, and Save says so
rather than inventing one. Where the browser supports the File System Access API the same file is
rewritten in place — and its existing contents are merged first, so entries saved on an earlier
visit survive. Otherwise it downloads `DCR_<plot id>.json`. The accumulated document is also held
in `localStorage`, so a failed or cancelled write never loses the entry.

### Uploading to Google Drive — currently disabled

Save writes locally only. The Drive upload is **commented out**, not deleted: the plumbing, the
panel markup, its styles and the wiring in `wire()` are all still in `dcr-calculator.src.html`,
each behind a `DISABLED:` marker naming the other pieces to bring back with it.

The JS is commented with `//` per line rather than one `/* */` block, because the code contains
block comments of its own and those cannot nest. The CSS could not be wrapped for the same reason,
so its rules are preserved as text inside a single comment.

What it did, for whoever revives it: an OAuth client id and a Drive folder link, entered per
browser; `response_type=token` in a popup, so no client secret; scope `drive.file`, which reaches
only files the page created; and find-then-patch on save, because Drive has no write-by-path and a
blind create would leave two files of the same name in the folder. The sign-in leg was never
exercised against a live client id — everything downstream of the token was.

### Local delivery

There are two ways the file can be handed over locally, because a plain download link is inert
inside the claude.ai artifact viewer. The page asks the host first (`claude.use("downloads")`, declared as the
`downloads` capability), and falls back to an ordinary anchor where there is no host to ask — which
is the case on GitHub Pages and any local server. A viewer declining the host prompt is reported as
declined, not as a failure.

## Batch export from a geodatabase

The page cannot read a file geodatabase — it is a directory of Esri binary tables, and parsing
it needs GDAL. So batch work is a companion script rather than a page feature:

```bash
pip install pyogrio openpyxl          # GDAL ships inside the pyogrio wheel
python3 gdb_export.py DMT_Plot_Entrance.gdb -o plots.xlsx
```

Plot areas come from `UDM_Plot.PLOTCALCULATEDAREA`. **FAR** comes from the layer's own
`DevCode_FAR` where it has one and from the designation schedule otherwise, with a column
recording which. **Max plot coverage** is not a geodatabase field, so it always comes from the
schedule, or the fallback default where the Code publishes none — again recorded per row.

Three sheets: **Plots** (a row per feature, filterable, frozen panes), **Parameters** (every
coefficient and source file used, so the output is self-documenting) and **Notes** (how each
column was produced and what is still unconfirmed).

Options: `--layer`, `--gla`, `--unit`, `--coverage-default`, `--space`, `--floor-use`.

### It does not compute parking per plot

The ITC rate depends on which activities occupy the plot, and no geodatabase field carries
that. Inventing a default activity mix would produce authoritative-looking numbers with no
basis. Instead the sheet gives the parking **envelope**, all of it derivable from the plot
alone:

| Column | |
| --- | --- |
| Open ground | plot area less the coverage — what is left outside the footprint |
| Spaces on open ground | how many of those the open ground holds, at the area per space |
| Usable per basement floor | what one basement floor yields, after the usable share |
| Spaces before L-3/L-4 | Max GFA x 2 / 100 — the demand ceiling the limitations set |

The last one is worth the column because the cap is a ratio against GFA, so the ceiling is
exact even though the demand that would meet it is not knowable without the activity schedule.
Below 1,600 m2 limitation L-4 bars a use that exceeds the ceiling; at or above it, L-3 permits
the use but makes the excess the developer's responsibility and grants no additional GFA. Each
row's note says which of the two applies to that plot.

### Validation against the sample geodatabase

`DMT_Plot_Entrance.gdb` carries a `DevCode_MaxGFA` field, which is an independent check on
step 1. Across its 66 plots: **57 have a usable FAR and all 57 agree exactly** with
`plot area x FAR`; the other 9 publish `N/A` in both sources and correctly derive no GFA.
One plot has a genuine `FAR = 0` (no buildable area) which is treated as a real value rather
than missing data, and its computed GFA of 0 matches the field.

Four designations in the geodatabase are absent from the Code schedule — `MU-13`,
`PLANNED DEVELOPMENT`, `RE-10`, `RE-7` — so those rows fall back to the default coverage and
say so. `PLANNED DEVELOPMENT` still gets its FAR, because the geodatabase supplies it directly.

## Schema

```sql
itc_rate(id, period, class_code, class_name, class_group, code, variant, sites,
         trip_am, trip_am_in, trip_noon, trip_noon_in, trip_pm, trip_pm_in,
         trip_phg, trip_phg_in,
         rate_employee, rate_visitor, rate_truck, rate_total,
         unit_label, driver_label, driver_kind, conversion)
weighting(period, days, share)
land_use(id, category, designation, name, far, coverage, remarks, source)
coefficient(key, value, note)
meta(key, value)

VIEW itc_combined(... , weekday_total, weekend_total,
                  rate_employee, rate_visitor, rate_truck, rate_total)
```

`period` is `'weekday'` or `'weekend'` — 141 codes per tab, identical code sets, matching
units and conversion factors. The build fails loudly if a code appears in one tab but not
the other, since the combined view would silently drop it.

The workbook also carries **trip generation** rates (AM / Noon / PM / PHG with inbound
percentage splits) and the surveyed-site count per class. These are stored but unused by
the calculator, which is a parking tool. Query `itc_rates.db` directly to reach them.

`driver_kind` is derived from the workbook's *Variables to be used* column and is what lets
the calculator bind the right quantity:

| `driver_kind` | rows | bound to |
|---|---|---|
| `gfa` | 62 | Max GFA (step 1) |
| `gla` | 1 | GLA (step 3) |
| `site` | 5 | plot area |
| `count` | 73 | manual entry — seats, units, bedrooms, students, beds, doctors, taxi bays, berths, invitees, fuelling positions |

## Land use designations

`land_use` holds all 59 base-district designations from the **Abu Dhabi Capital Development
Code** (via `Designation_Index.xlsx`), each with Maximum FAR and Maximum Plot Coverage.
Picking one from the Inputs panel fills both; either can be typed over, and blanking the
coverage field returns it to the published value.

`coverage` is stored as a **fraction** (0.6 = 60%) exactly as the Code publishes it, and
displayed as a percentage. `far` and `coverage` are `NULL` where the Code publishes no
number — **15 designations have no FAR**, 12 no coverage — governed by note reference
(N-15, N-25, N-28 …) instead. The `remarks` column carries the reason, surfaced as an
advisory when such a designation is selected; the calculator refuses to compute rather than
guess.

This confirmed the whiteboard: **NR = NEIGHBOURHOOD RETAIL, FAR 1.05, Max Plot Coverage
0.6**. The 60% was never an arbitrary coefficient — it is NR's published coverage.

Coverage spread across the schedule: 27 designations at 100% (towers and mid-rise blocks),
the rest between 30% (AG) and 88% (R15-E).

## The week-weighted rate

A week is five weekdays and two weekend days, so the default parking rate is the
day-weighted blend of the two tabs:

```
share(weekday) = 5 / 7 = 71.43%
share(weekend) = 2 / 7 = 28.57%
rate           = weekday_rate × 0.7143 + weekend_rate × 0.2857
```

applied to each of the three components (employee/resident, visitor, truck/bus) and to the
total. The *Rate basis* control switches between **Week** (the blend, default), **Weekday**
and **Weekend**. The day split is editable, and the `weighting` table is updated in step so
the `itc_combined` view always matches what is on screen.

**Only 19 of the 141 codes actually differ between the two tabs** — mostly schools,
nurseries and warehousing, where weekend demand is higher. For the other 122 the blend
returns the weekday figure unchanged. The rate panel shows both inputs, the resulting week
figure, whether the tabs differ, and which basis is in use.

Worked example — Warehousing `713` on a 1,000 m² plot at FAR 1.05 (GFA 1,050 m²):

| Basis | Rate | Bays |
|---|---|---|
| Weekday | 0.515 | 5.408 → **6** |
| Weekend | 1.028 | 10.794 → **11** |
| Week (5:2) | 0.515 × 0.7143 + 1.028 × 0.2857 = **0.661571** | 6.946 → **7** |

## The calculation

```
1  Max GFA        = plot size x FAR                  FAR from the designation
2  Plot coverage  = plot size x coverage%            GFA allowed on the ground floor
3  GLA            = Max GFA x GLA%                   GLA% an input, default 75%
4  Activities     = floor( GLA / unit area )         unit area an input, default 60 m2

   Spaces         = sum over the activity schedule   or a UPPC/ITC override
   Parking        = 32.5 m2 each, open ground first then basement
```

Parking is no longer a derivation step. It depends on *which* activities occupy the plot,
so it comes from the activity schedule and has its own panel.

Coverage drives the footprint in step 2;
so it is a single input, auto-filled from the designation and editable.

Worked example — 1,000 m&sup2; plot, code `NR` (FAR 1.05), ITC `112` Local Shopping Centre:

| Step | Result |
|---|---|
| Max GFA | 1,050 m&sup2; |
| Plot coverage | 1,000 x 0.6 = **600 m&sup2;** |
| GLA | 787.5 m&sup2; |
| Parking | (1,050 / 100) x 1.318 = 13.839 &rarr; **14 bays** (one ceiling, on the total) |
| Activities | 787.5 / 60 = 13.125 → **13 activities** |
| Parking (13-activity example) | **26 bays** = 26 &times; 32.5 = **845 m&sup2;** |
| On open ground | 1,000 &minus; 600 = 400 m&sup2; holds &lfloor;400 / 32.5&rfloor; = **12 spaces** |
| In basement | 26 &minus; 12 = **14 spaces** = 455 m&sup2;, over 750 m&sup2; usable &rarr; **1 floor** |
| L-3 / L-4 | 26 bays is 2.48 per 100 m&sup2; GFA, over the cap of 2 &mdash; and the plot is under 1,600 m&sup2;, so **L-4 bars it** |

Steps are ordered by dependency rather than as drawn on the source whiteboard: GLA moves
ahead of parking because one category (`111` Regional Shopping Centre) is charged
against GLA, not GFA.

## Activity schedule

Step 5 gives how many activities the GLA supports. The **Activity schedule** panel fills those
slots with real uses and prices the parking per activity, which is finer than charging the whole
plot at one ITC rate.

Each row is an activity plus the number of slots it occupies:

- **Activity** — searched from the DED activity list (3,892 active activities from
  `DED قائمة الانشطة`, deduplicated by `ACTIVITY_ID`).
- **ITC category** — auto-mapped from the activity, overridable per row. The confidence tag
  says how far to trust the mapping.
- **Unit area** — per row. The field always shows the effective figure, so the spinner nudges it
  from there (60 → 61) rather than jumping to the input's minimum. Slots, unit area and quantity
  accept whole numbers only, and editing one repaints just that row's derived cells rather than
  rebuilding the table, so the field keeps focus while you type. A row follows the Fixed-values
  default until you change it, and clearing the field hands it back. Each row is tagged
  `inherited` or `entered` so a hand-set area is visible at a glance, and the footer counts how
  many were entered. Because rows can differ, the footer also totals the floor area the schedule
  consumes, which the slot count alone no longer tells you.
- **Bays** — `slots × unit area × conversion × rate`, carried **unrounded per row**. The ceiling
  is applied **once, to the total**. Rounding each row up first charges for part spaces the scheme
  never needs: `2.372 + 5.536 + 5.99` is `13.898 → 14` spaces, but per-row ceilings make it
  `3 + 6 + 6 = 15`. The error compounds with the number of rows, so a long schedule drifts further
  from the true figure.
  The slots an activity holds are its share of GLA, and for GFA/GLA-charged categories that share
  is the driver. **73 of the 141 classes are not charged by area at all** — they are charged per
  seat, unit, bedroom, student, bed, doctor, berth, taxi bay, invitee or fuelling position — and
  those rows show a quantity box in the *Rate / units* column instead of a rate, with the per-unit
  rate underneath it. The box is keyed on the class's `driver_kind`, so it stays visible and
  editable once a count is entered; a row whose quantity is still blank raises the `QTY` flag and
  contributes no parking rather than being silently charged by floor area.
- **ITC location** (Abu Dhabi/Al Ain × CBD/non-CBD) picks which variant of a class applies,
  falling back progressively when a class is not split that finely.

The footer totals slots against the permitted count and flags an overrun. Bays are totalled but
not capped there: whether the demand clears L-3/L-4 is a ratio against GFA, so it is raised as a
flag rather than a column total.

### The floor area the schedule consumes

The slot count stops standing in for floor area as soon as unit areas vary. Six slots at a
hand-entered 260 m² can sit inside the permitted count and still overrun the plot, so the floor
area is checked in its own right:

| | |
|---|---|
| over **Max GFA** | a `GFA` **restriction** — the FAR does not permit the floor area, full stop |
| over **GLA** but within GFA | a `GLA` note — it fits the building but not the leasable area |
| within both | the footer shows `floor area x / y m²` with no flag |

Both are needed because they answer different questions, and the slot counter answers neither: a
schedule can read *spare 2 slots* while already exceeding the leasable area. GLA is the tighter
bound — tenancies are let from leasable area, not gross — but only GFA is treated as a hard cap,
since GLA is itself a 75% assumption rather than a published figure.

### The DED → ITC crosswalk

`ITC_RULES` in `build_db.py` is a **heuristic mapping written for this tool, not a client
crosswalk**. It keys on the numeric ISIC levels, because the workbook's `LEVEL_0` letters are
re-lettered and do not follow ISIC sections. Coverage of the 3,892 activities:

| Confidence | Count | Meaning |
|---|---|---|
| high | 162 | the ISIC class and the ITC category describe the same thing |
| medium | 2,722 | division-level match, e.g. all specialised retail → Local Shopping Centre |
| low | 1,008 | closest available analogue; review before relying on it |

Validated against the worked example — flower shop, book shop, commercial bank and fast food
restaurant all auto-map to the ITC categories whose rates the brief quotes (1.318, 1.318,
4.992, 7.702). Any row can be reassigned by hand, which then shows as *set by hand*.

### Parking: open ground first, then basement

A space needs its own area, so the ground left outside the footprint holds only so many. The
remainder goes to basement, where a floor is only partly usable — ramps, cores, plant — which is
what sets how many floors it takes.

A basement floor defaults to the **whole plot**, not the footprint: unlike a storey above ground
it is not held to the coverage limit. That is why `parkFloorArea` originates from the plot area
and not from the coverage.

```
total parking area required   = spaces x 32.5 m2
open ground area              = plot area - plot coverage
spaces on ground              = min(required, floor(open ground / 32.5))
offset spaces required        = required - on ground
offset parking area required  = total parking area - on-ground area
basement floors               = ceil(offset area / (basement floor area x 75%))
```

The panel reads as two rows of four: the requirement and what the open ground absorbs, then
what is left of the requirement and where it goes. *Offset* is the residual after the ground
has taken what it can, so **offset spaces** and **in basement** always carry the same number —
one states it as a remaining requirement, the other as a placement. That is deliberate, not a
duplicated calculation: `total area - on-ground area` is identically `basement spaces x 32.5`,
because the on-ground count is itself a whole number of spaces.

Both cases from the brief, reproduced exactly:

| | general plot | Community Retail |
|---|---|---|
| Spaces required | 30 | 120 |
| Total parking area | 975 m² | 3,900 m² |
| Plot / coverage | 1,000 m² / 60% | 1,000 m² / **100%** |
| Open ground | 400 m² → holds 12 | **0 m² → holds 0** |
| On ground | 12 | 0 |
| In basement | 18 → 585 m² | 120 → 3,900 m² |
| Floors at 75% usable | 1 | **6** |

**Community Retail forces the floors calculation, and it follows from the Code rather than a
special rule:** `CR` is published at **100% plot coverage**, so there is no open ground and every
space goes to basement. `3,900 / (1,000 x 0.75) = 5.2 -> 6 floors`.

The number of spaces required comes from the activity schedule, or from the *Required spaces*
override where UPPC or ITC has granted a reduction.

This replaced an earlier model that capped parking at 50% of plot area and made basement parking
mandatory above 16% of coverage. Those two *thresholds* were stand-ins with no source behind them
— the basement itself was never in doubt, only the invented trigger for reaching it. The area per
space (32.5 m²) and the 75% floor efficiency come from the standard, so the two assumptions
flagged in earlier versions are now closed.


## Folded by default

**Plot & footprint** starts collapsed — it is a sanity check on the geometry rather than an
output, so it sits behind its own header until clicked. Note the CSS trap if you add another
folding panel: the rule that hides a closed `<details>`' children is a *user-agent* rule, so any
author `display` on a child (here `.plan { display: grid }`) wins and the panel stays visibly
open. `.panel.fold:not([open]) > *:not(summary){display:none !important}` is what actually closes
it.

## Removed from the page

Taken out on request: the **ITC rate conversion matrix** browser, the editable **Land use
designations** table, the **Query the database** console, the SQL readout under the ITC rate
card, and the **DCR report sample** view (an A3 facsimile of an issued report, with its plot
drawing, coordination table, setback ring and boilerplate notes).

The reference data those panels exposed is untouched in `itc_rates.db`. What went with the
report view was the sample plot's identity and survey coordinates, which were only ever its
default input values — see *Client data* below.

Consequence worth knowing: there is no longer a way to add a designation or import a FAR
schedule from the page, so corrections go through `build_db.py`.

## Client data

This repository holds no plot-specific data. The DCR report sample previously carried a real
parcel's number, district, sector and cadastral UTM coordinates as its default field values;
all of it was removed along with that view, and the codebase was swept to confirm none
remains.

What is here, and is fine to hold: the **land-use designation schedule** (published in the
Abu Dhabi Capital Development Code) and the **ITC trip and parking rates**. Neither is
plot-specific.

The source workbooks live outside this tree and should stay out of git. Before publishing
this anywhere, re-read *Confirmed vs. assumed* below — the figures are engineering
work-in-progress, not cleared outputs.

## Confirmed vs. assumed

Confirmed from source:

- ITC rates — all 141 codes × weekday and weekend, from
  `ITC_Trip & Parking Calculation Sheet_V1.xls`.
- **All 59 land-use designations** with Maximum FAR and Maximum Plot Coverage, from the Abu
  Dhabi Capital Development Code via `Designation_Index.xlsx`. `NR` &rarr; FAR 1.05,
  coverage 60% — which is where the whiteboard's 60% came from.
- The whiteboard's ITC "1.3" is `112`'s three components summed: 0.107 + 1.204 + 0.007 =
  **1.318**. The calculator uses the unrounded figure; both round up to 14 bays.

Still needs sign-off:

1. **An issued DCR report gave `AVG. REQUIRED PARKING = 0`** for a small commercial plot, where
   the ITC rate for that land use implies 3 spaces. Not a rounding artefact. The most important
   one — it decides whether that report field can be trusted at all.
2. **The Arabic note on the whiteboard** ("50% of the PCT") most likely points at limitation
   **L-2** of the issued report: at least 50% of GFA to the specified uses, and not less than
   50% of leasable area. Surfaced as an advisory, not applied as a rule.
3. **The GLA 75% share** is not a Code citation — it is the worked example's value and may vary
   by scheme. Overridable, and seeded from the `coefficient` table.
4. **The DED → ITC crosswalk** is heuristic: 162 high, 2,722 medium, 1,008 low confidence across
   3,892 activities. Review the low-confidence rows.
5. **15 designations publish no FAR** and 12 no coverage; the Code governs them by note
   reference. Those note definitions live on the district pages of the Code PDF and were not
   extracted, so those designations need an override to calculate.

Closed since earlier versions:

- **Plot coverage basis** — it is `plot area x coverage%`, the GFA allowed on the ground floor.
- **Step 5** — a floored count, `floor(GLA / unit area)`, not an area.
- **Area per parking space** — 32.5 m², from the standard, replacing an invented 25 m² bay.
- **Where parking goes** — open ground first, then basement, replacing an invented 50% cap and
  a 16%-of-coverage trigger for going below.
- **Which parking figure governs** — the activity schedule; the whole-plot single-category basis
  was removed.
