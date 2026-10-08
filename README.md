# Prior Auth RPA Dashboard

An interactive dashboard reporting RPA (bot) coverage and success on prior
authorizations, built to be embedded in Notion via GitHub Pages.

```
Orders_and_Appointments_Dashboard.xlsx   <- the periodic export (NEVER committed)
          |
          |  Build-Dashboard.ps1
          |    1. sanitize   <- strip PPI at ingest, before anything else
          |    2. normalize
          |    3. aggregate
          |    4. publish gate
          v
sanitized/*.sanitized.csv                 <- PPI-free row-level copy (local only)
docs/data.json                            <- aggregate counts only (publishable)
docs/index.html                           <- the dashboard (static, no build step)
          |
          v
GitHub Pages  ->  Notion /embed block
```

---

### PPI is removed at ingest

The build strips identifiers **first**, before normalizing or aggregating
anything, so no later stage ever sees them. This is not optional and there is no
flag to turn it off.

**`RC Assigned`** is collapsed to one of three literals. Only
`svc-quickbase` is kept as itself; every human assignment becomes `(human)` and
every blank becomes `(unassigned)`. Names and user IDs are discarded, not masked
— the dashboard only ever needed "bot or not":

| Raw value (shape only) | Becomes |
|---|---|
| `svc-quickbase svc-quickbase <SERVICE-ACCOUNT-ID>` | `svc-quickbase` |
| `<First Last> <First Last> <USER-ID>` | `(human)` |
| *(empty)* | `(unassigned)` |

**`Bot Run Fail Reason`** is redacted in place. Dates become `[date]`, people
become `[name]`, and identifiers become `[id]`:

```
Member Id is not found-<ID>            ->  Member Id is not found-[id]
The authorization number is: <ID>      ->  The authorization number is: [id]
Due date <MM/DD/YYYY> expired          ->  Due date [date] expired
```

The name list is **harvested from the file's own `RC Assigned` column** on every
run, so it stays current as staff change — there is no roster to maintain by
hand. Redaction order matters: digits glued to a word are stripped before the
whole-token rule, so a token like `matches` followed immediately by a digit run
becomes `matches[id]` rather than bare `[id]`, which keeps the
failure-category rules below matching.

A PPI-free row-level copy is written to
`sanitized/<name>.sanitized.csv` so you can inspect or share the data without
touching the original export (which is never modified). Pass
`-SkipSanitizedCsv` to skip writing it. That folder is gitignored.

### Four safeguards. Do not disable them.

1. **Ingest sanitization**, above — identifiers are destroyed before the data
   is used for anything.
2. **`.gitignore` blocks `*.xlsx` / `*.csv` / `sanitized/`** so neither the
   source export nor a row-level copy can be committed by accident.
3. **The published label set is closed.** `Normalize-Reason` returns one of
   ~35 fixed category labels or the literal `Other / unclassified`. Input text
   is **never echoed through**, so a brand-new failure message carrying an
   identifier cannot reach `data.json` even in principle — it lands in
   `Other / unclassified` instead. And `data.json` holds aggregate counts only;
   there is no row-level field for an identifier to occupy.
4. **A publish gate** asserts every emitted label is in the closed set, then
   scans all labels for digit runs, ID phrasing, and any harvested coordinator
   name. If anything suspicious survives, the build **deletes its own output and
   fails** rather than emit a publishable file. (Verified non-trivial: the gate
   rejects **774 of the 905** distinct raw reason strings in the current export,
   and **0** of the 40 published labels.)

If `Other / unclassified` starts growing, that is the signal to add a rule to
`$REASON_RULES` in `Build-Dashboard.ps1` — the data is still safe, just less
informative. It currently holds **117 of 5,917** failure rows (2.0%).

---

## Refreshing the data

When a new export arrives:

1. Drop it in the project root. Run bare, the build picks up the newest file
   named `Orders_and_Appointments_Dashboard.xlsx` / `.xls` / `.csv`; any other
   name needs `-Source`. (`.xlsx` is read through Excel, `.csv` directly.)
2. Run the build:

   ```powershell
   .\Build-Dashboard.ps1 -Source "Orders_and_AppointmentsV1.05.csv"
   ```

   It prints what it redacted, which date column it bucketed on, and the
   headline metrics, so you can sanity-check all three before publishing:

   ```
   Sanitized: RC Assigned -> 13496 bot / 91944 human / 0 unassigned
              (67 coordinator names discarded)
   Sanitized: 834 fail reasons had names/identifiers redacted
   Date basis     : Date Auth Completed Roll Up
   Rows aggregated: 105417  (dropped for unparsable date: 0,
                             dated after the 2026-10-08 pull: 23)
   Publish gate: no identifier-like text in 40 reason categories
   ```

   **Check the `Date basis` line.** If it ever says anything other than
   `Date Auth Completed Roll Up`, every chart has silently changed meaning —
   see *The date basis* below.

3. Preview locally (optional):

   ```powershell
   .\Serve-Dashboard.ps1
   ```

   This starts a localhost server and opens a browser. A local server is
   required — opening `docs/index.html` straight from disk makes the browser
   block the data fetch (`file://` security), which looks like a broken
   dashboard but isn't.

4. Publish **only** `docs/data.json` (see below). `docs/index.html` never needs
   to change on a refresh.

The build accepts explicit paths if your file is named differently:

```powershell
.\Build-Dashboard.ps1 -Source "C:\path\to\export.xlsx"
```

Other switches:

| Switch | Effect |
|---|---|
| `-OutFile <path>` | Write the aggregate JSON somewhere other than `docs\data.json`. |
| `-SkipSanitizedCsv` | Don't write the `sanitized/` row-level copy. Sanitization itself still happens — it cannot be skipped. |

---

## One-time setup: GitHub Pages

Your repo: <https://github.com/NicholasWilcoxAkido/Prior_Auth_Dashboard>

`git` is not installed on this machine, so these steps use the GitHub web UI.
(If you'd rather use the CLI, install Git for Windows and the commands are the
obvious ones — the layout is already correct.)

### 1. Upload the dashboard

1. Open the repo and click **Add file → Upload files**.
2. Drag in the **`docs` folder** (both `index.html` and `data.json`).
   Keep them inside a folder named `docs` — Pages is configured to serve it.
3. Also upload `README.md`, `.gitignore`, `Build-Dashboard.ps1`, and
   `Serve-Dashboard.ps1` so the pipeline is reproducible.
4. Commit to `main`.

> Do **not** upload the `.xlsx` / `.csv` export, the `sanitized` folder, or
> the `_shots` folder.

### 2. Turn on Pages

1. Repo **Settings → Pages**.
2. **Source:** `Deploy from a branch`.
3. **Branch:** `main`, **Folder:** `/docs`. Save.
4. Wait ~1–2 minutes for the first deploy. Your URL will be:

   ```
   https://nicholaswilcoxakido.github.io/Prior_Auth_Dashboard/
   ```

Open it directly in a browser first and confirm it loads before embedding.

### 3. Refreshing later

Go to `docs/data.json` in the repo → pencil (**Edit**) or **Upload files** to
replace it → commit. The dashboard picks it up on next load; the page already
cache-busts the fetch so you won't see stale numbers.

---

## Embedding in Notion

1. In your Notion page, type `/embed` and press **Enter**.
2. Paste your Pages URL and click **Embed link**:

   ```
   https://nicholaswilcoxakido.github.io/Prior_Auth_Dashboard/
   ```

3. **Resize it.** The default embed is far too short. Drag the handle at the
   bottom edge down until the whole dashboard is visible — roughly 1,600–2,000px
   for the full view. Notion embeds don't auto-size to content.
4. Optional: click the `⋮⋮` handle → **Full width** so the charts get more room.

### Useful URL parameters

Append these to the embed URL to control the initial view:

| Parameter | Values | Effect |
|---|---|---|
| `?theme=` | `dark`, `light` | Pins the theme. Use `dark` to match a dark Notion workspace — the embed can't detect Notion's theme on its own. |
| `?range=` | `14d`, `1m`, `1q`, `ytd`, `all` | Sets the opening date range. Older values (`7d`, `30d`, `90d`, `mtd`) still work — they map to the nearest current preset, so existing embeds don't break. |
| `?type=` | `Appointment`, `Diagnostic`, `Procedure`, `Referral` | Sets the opening Type filter. |

Combine with `&`:

```
https://nicholaswilcoxakido.github.io/Prior_Auth_Dashboard/?theme=dark&range=1q
```

A good pattern is several embeds on one Notion page, each pinned to a different
range — e.g. a 14-day operational view near the top and a YTD trend below.

### Notion caveats

- **Interactivity works** inside the embed (filters, hover tooltips, table
  toggles) because it's a real iframe.
- **Notion caches embeds aggressively.** If a refresh doesn't show, click the
  embed and use **⋮⋮ → Reload**, or reload the Notion page.
- **Mobile** — Notion's mobile app renders embeds short. The layout is
  responsive and will stack, but the desktop view is the intended one.

---

## The date basis — read this first

Every date on this dashboard is **`Date Auth Completed Roll Up`**. The dashboard
reports **throughput**: what got *finished* in a period. It is not an intake
report, and the distinction is not cosmetic — a September *completion* cohort and
a September *creation* cohort overlap by barely half. September 2026 reads
**11,993** across all types on this basis; on the old creation basis the same
month read **7,768**. Both were correct; they answer different questions.

The source export must carry all **seven** columns, and the build fails loudly if
any are absent:

```
Date Created (Roll-Up)          Bot Run Fail Date
Type                            Bot Run Fail Reason
Status (Referral Coordinator)   Date Auth Completed Roll Up   <- the date basis
RC Assigned
```

`Date Created (Roll-Up)` is still required and still kept in the sanitized copy,
but nothing is bucketed on it. Switching back is a one-line change —
`$COL_BASIS` in `Build-Dashboard.ps1` — and it rewrites the meaning of every
chart, so don't do it quietly.

### Three things this basis costs you

1. **History starts 2026-01-01, hard.** Not one row in the export has a
   completion date in 2025, even though 10,657 rows were *created* between July
   and December 2025 — those all carry 2026 completion dates, decaying from 4,588
   in January to a few dozen by autumn. Whether that is a real cleared backlog or
   an artifact of when the field started being populated is a **Quickbase
   question, not a dashboard one**. Either way the x-axis cannot go back further,
   so `YTD` and `All` currently show the same window.
2. **The open backlog is invisible.** `Date Auth Completed Roll Up` is **100%
   populated** — zero blanks in 105,440 rows. A genuine completion field would be
   blank for anything still in flight, so in-progress authorizations either aren't
   in this export or aren't distinguishable in it. Nothing here can tell you how
   much work is currently open.
3. **About 6% of rows complete before they're created.** 6,536 rows have a
   completion date earlier than their creation date, which is impossible for a
   true timestamp pair. It doesn't affect any current metric — nothing computes a
   duration — but it does mean **turnaround time cannot be built on these two
   columns** until it's explained.

A fourth issue was found and handled: **23 rows carry completion dates after the
day the export was pulled**, the furthest 2027-03-30. Date presets anchor to the
latest date in the file, so that single 2027 row would have dragged the "14 Days"
window five months into empty calendar and opened the dashboard blank. The build
now drops anything dated past the export's own pull date and prints the count.

---

## Metric definitions

Scoped to the current Type, Status, and date filters. "In the period" means
**completed** in the period. Type is matched
case-insensitively, because the export mixes `Appointment` with `DIAGNOSTIC`
and `REFERRAL`. `RC Assigned` is tested with *contains* `svc-quickbase` against
the raw value, before sanitization collapses it — so the metrics are unaffected
by the redaction. (Confirmed: the headline figures are byte-identical with and
without the sanitize stage.)

| Metric | Definition |
|---|---|
| **Total Schedule Authorizations** | Count of authorizations **completed** in the period where `Type = Appointment`. |
| **Touched by RPA** | **Distinct** rows where `RC Assigned` contains `svc-quickbase` **or** `Bot Run Fail Date` is populated. |
| **Percent Touched by RPA** | Touched ÷ Total. |
| **RPA Success Rate** | Rows where `RC Assigned` contains `svc-quickbase` ÷ Touched. |

### Three decisions worth knowing about

**Overlapping rows are counted once.** 462 Appointment rows have *both* an
`svc-quickbase` assignment and a `Bot Run Fail Date`. Summing the two counts
would double-count them, so "touched" uses the de-duplicated union. These rows
are treated as **bot failed, then succeeded on retry** — they count as touched
*and* as successes. Across the full V1.05 export:

```
12,894  bot succeeded
   462  bot succeeded after a retry
 5,009  bot failed
------
18,365  touched by RPA   (not 18,827 — that figure double-counts the 462)
```

This makes all-time coverage 30.5% rather than 31.2%, and success rate 72.7%.
Note that all-time coverage is essentially **unchanged** by the switch to the
completion basis (30.41% → 30.45%), which is the expected result: re-dating rows
moves them between periods but doesn't add or remove any. The basis changes
*period* figures, not the grand total.

**A delta is only shown when the comparison period is fully present.** The
▲▼ figures compare against the preceding window of equal length. If that window
starts before the first row in the export, no delta is shown at all — an earlier
version compared YTD against a partly-empty 2025 window and reported "+401%
growth" that was really just missing history.

**Date presets anchor to the latest date in the export, not to today.** A
"last 7 days" window measured from today would come up nearly empty whenever the
export is a few days stale. Anchoring to the data keeps the default view
populated. The exact window is always printed at the right of the filter row.

---

## Reading the dashboard

- **Filters** sit in one row and scope everything below them. Presets are
  **14 Days · 1M · 1Q · YTD · All**; the default is **1Q** with
  `Type = Appointment`. `1M` and `1Q` are inclusive calendar months (a `1M`
  window ending Oct 8 starts Sep 9), and `YTD` runs from Jan 1 of the latest
  year in the export. A quarter is wide enough that the charts bucket by week,
  which is why the opening view is weekly rather than daily. **`YTD` and `All`
  currently show the same window**, because the earliest completion date in the
  export is Jan 1, 2026.
- **Every number is a count of authorizations *completed* in the window**, not
  requested in it. See *The date basis* above — this is the single most
  misreadable thing about the dashboard.
- **Every chart has a `Table` toggle** — the accessible, copy-pasteable twin.
  The failure-reason table lists *all* categories, not just the charted top 12.
- **"Authorization Volume vs. Volume Touched by RPA" plots two counts on one
  shared axis.** Touched is a strict subset of total, so the filled band is
  RPA's share and **the gap between the two lines is the volume still handled
  manually** — that gap is the thing to watch. Because both series are counts in
  the same unit, one axis is correct here; no second y-axis is used anywhere in
  this dashboard, since a dual-axis chart lets the apparent crossover point be
  moved anywhere just by rescaling. Percent touched and success rate are still
  in that chart's tooltip and its `Table` view, which also breaks out a
  Handled Manually column.
- **The two bottom charts deliberately ignore their own filter** so the full
  Type / Status mix stays visible while you filter the rest of the dashboard.
- **`Patterns`** adds directional fills so series stay distinguishable without
  relying on color (colorblind-safe, print-safe).
- **Daily percentages swing hard on low-volume days.** A day with 2 auths can
  read 0% or 100%. Hover any point for the underlying counts, or widen the range
  — past 70 days the charts bucket by week automatically.

---

## Possible next steps

Ideas beyond the current scope, roughly in order of value:

1. **Turnaround time** — the export now has both a creation and a completion
   date, so median days from created → completed, split bot vs manual, is within
   reach and is probably the strongest ROI metric available. **Blocked on one
   data question:** 6% of rows complete before they're created (see *The date
   basis*), so the pairing isn't yet trustworthy enough to publish a duration.
2. **Open backlog / aging** — needs an export where
   `Date Auth Completed Roll Up` is blank for in-flight work. It is currently
   100% populated, so open authorizations can't be counted at all.
2. **Hours saved** — touched-and-succeeded × an agreed minutes-per-auth figure.
   Needs one number from you to be credible.
3. **Failure-reason trend** — which categories are growing, not just which are
   biggest. Tells you where to point bot fixes.
4. **A target line** on the success-rate chart, once there's an agreed SLA.
5. **Retry depth** — currently a row is retried-or-not. If the source tracks
   attempt counts, retry distribution would show where the bot thrashes.
6. **Scheduled refresh** — a Task Scheduler job running `Build-Dashboard.ps1`
   against a network-drop folder, plus `git push`, would remove the manual step.

---

## Files

| File | Purpose |
|---|---|
| `Build-Dashboard.ps1` | Export → `docs/data.json`. Sanitizes, normalizes, aggregates, gates. |
| `Serve-Dashboard.ps1` | Local preview server (no Python/Node needed). |
| `docs/index.html` | The dashboard. Static, dependency-free, never regenerated. |
| `docs/data.json` | Generated aggregate data. The only data file that is published. |
| `sanitized/` | PPI-free row-level copies of each export. Local only, gitignored. |
| `.gitignore` | Blocks source exports and sanitized copies from being committed. |

No build toolchain, package manager, or CDN dependency — `docs/index.html` is
plain HTML, CSS, and vanilla JS, and the charts are hand-rolled SVG. It will
still work years from now.
