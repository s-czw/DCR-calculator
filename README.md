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
way. The status chip in the header reads `sqlite` or `snapshot` so you can tell which
path ran.

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

Options: `--layer`, `--gla`, `--unit`, `--coverage-default`, `--bay`, `--cap`, `--ug`.

### It does not compute parking per plot

The ITC rate depends on which activities occupy the plot, and no geodatabase field carries
that. Inventing a default activity mix would produce authoritative-looking numbers with no
basis. Instead the sheet gives the parking **budget** — how much surface parking the cap
allows, and the area above which parking must go underground — both derivable from the plot
alone.

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

   Parking        = sum over the activity schedule   see below
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
| Parking | (1,050 / 100) x 1.318 = 13.839 &rarr; **14 bays** |
| Activities | 787.5 / 60 = 13.125 → **13 activities** |
| Parking (13-activity example) | **26 bays** = 650 m² = 65% of plot → **over the 50% cap** |
| Share of coverage | 650 / 630 = **103%** → **underground parking mandatory** (trigger 16%) |

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
- **Bays** — `slots × unit area × conversion × rate`, **rounded up per row**: each activity is a
  separate tenancy, so it takes whole bays. The raw figure is shown underneath. Rounding per row
  rather than once at the end costs a little more parking — 26 bays against 24 on the example.
  The slots an activity holds are its share of GLA, and for GFA/GLA-charged categories that share
  is the driver. Categories charged per seat, unit or student instead ask for a quantity.
- **ITC location** (Abu Dhabi/Al Ain × CBD/non-CBD) picks which variant of a class applies,
  falling back progressively when a class is not split that finely.

The footer totals slots against the permitted count, and bays against the parking cap.

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

### Underground parking

Above a share of the **allowed plot coverage** (step 2), surface parking stops being viable and
the panel states that underground parking is mandatory within the plot. The trigger is an input,
default **16%**.

That is a low bar in practice. On the worked example — 600 m² of allowed coverage — 16% is
96 m², which at 25 m² a bay is **under four bays**. So almost any multi-tenant scheme
will require underground parking. That may well be the intent for dense plots, but it is worth
confirming the rule means 16% *of coverage* rather than of plot area, and that it compares an
**area** rather than a bay count.

The requirement is styled apart from the parking-cap breach: exceeding the cap is an error,
needing underground parking is a requirement the design can satisfy.

### The parking cap

The rule is *"if the chosen activities need more than 50% of the plot as parking, error"*. ITC
rates give a **bay count**, and the cap is an **area**, so converting between them needs an area
per bay. That figure is not in any source, so it is an input: **25 m² by default**, a typical
bay-plus-aisle allowance. **Confirm it** — the example lands at 600 m² against a 500 m² cap, so
the verdict flips somewhere near 21 m² a bay.

The **Parking requirement** panel below the schedule reports bays, the area they need, that area
as a share of the plot, and the cap — turning red when the cap is breached. The masthead chip
carries the same figure.

The earlier whole-plot basis (one ITC category applied to total GFA) has been removed, along with
the ITC category picker and the ITC rate card, so there is now a single parking figure rather than
two that disagreed.

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

1. ~~**Plot coverage basis.**~~ **Resolved 2026-08-22:** plot coverage is the GFA allowed on
   the ground floor, `plot size x max plot coverage%`. The whiteboard's `Max GFA x coverage%` was
   wrong, which the issued report had hinted at by printing `MAX. PLOT COVERAGE 60%` as a bare
   percentage with no area. On the worked example this moves coverage from 630 m² to 600 m², and
   the footprint/plot ratio from 63% to exactly 60%.
2. **The handwritten Arabic note** reading roughly "take 50% of the PCT" now has a likely
   referent: limitation **L-2** of the issued report requires "a minimum of 50% of the
   plot's total GFA … allocated to the specified uses and … not less than 50% of the total
   leasable floor area." The calculator points at L-2 and shows both 50% figures, but still
   treats it as an advisory, not a rule.
3. **The GLA 75% share** is not a Code citation — it is the worked example's value and may
   vary by scheme. It is an input, seeded from the `coefficient` table.
4. **Parking bay area (25 m²).** Introduced here so a bay count can be compared with an area
   cap; not from any source. The example's pass/fail flips near 21 m² a bay.
5. **The DED → ITC crosswalk** is heuristic — 4% high confidence, 26% low. Review low-confidence
   rows.
6. **Which parking figure governs** — the whole-plot ITC category or the activity schedule.
7. ~~**Step 5's unit change.**~~ Resolved: reverted to the whiteboard's count, `floor(GLA / 60)`.
   Superseded note: `GLA × coverage%` returned an area, where the whiteboard
   returned a tenancy count. See *The calculation* above.
5. **15 designations publish no FAR** and 12 no coverage; the Code governs them by note
   reference. Those note definitions live on the district pages of the Code PDF and were not
   extracted, so those designations cannot be calculated without an override.

Form state persists to `localStorage`. Both reference tables are read-only in the page —
edit `Designation_Index.xlsx` or the ITC sheet and re-run the build to change them.

## Removed from the page

Three panels were taken out on request: the **ITC rate conversion matrix** browser, the
editable **Land use designations** table, and the **Query the database** console. The data
they exposed is untouched in `itc_rates.db`; only the in-page UI for browsing, editing and
querying it is gone. The consequence worth knowing: there is no longer a way to add a
designation or import a FAR schedule from the page, so corrections have to go through
`build_db.py`. The SQL readout under the ITC rate card was dropped at the same time.
