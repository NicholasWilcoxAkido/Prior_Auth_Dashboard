# Prior Auth RPA Dashboard

An interactive dashboard reporting RPA (bot) coverage and success on prior
authorizations, built to be embedded in Notion via GitHub Pages.

```
Orders_and_Appointments_Dashboard.xlsx   <- the periodic export (NEVER committed)
providers.hard.txt                       <- hard-schedule roster (NEVER committed)
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

## ⚠️ Read this before you push anything

The raw export is **not** safe to publish. `Bot Run Fail Reason` carries member
identifiers and authorization numbers inside the error text, and `RC Assigned`
carries coordinator full names and user IDs. The shapes to watch for, written
here with **invented** values — every example in this file is synthetic, never
copied from an export:

```
The Referral is already submitted Member ID: <11-digit number>
Member Id is not found-<digits + letter + digits>
There is a previous authorization on file.The authorization number is: A<8 digits>
<First Last> <First Last> <8 digits>.<4 chars>
```

Member IDs are HIPAA identifiers. Publishing them to a public GitHub Pages site
is effectively irreversible — crawlers and CDN caches retain content even after a
force-push or repo deletion. That applies to this README too: **do not paste real
error text into it** when adding a rule. Describe the pattern instead.

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

**The two columns V1.07 added are treated differently from each other.**
`Insurance Portal` is published as a dashboard dimension — the values are payer
and vendor brand names (`Aerial`, `IEHP`, `Medpoint`, …), which carry nothing
about a member. Unlike failure reasons, that label set is **open**, because the
payer roster isn't ours to close, so the publish gate scans portal labels for
digit runs and coordinator names exactly as it scans the rest.
`Encounter Provider` holds provider **names**, the same class of personal data as
the coordinator names in `RC Assigned`, so it is never published in any form.
It is read for exactly one purpose — deciding whether an Appointment is an *easy*
or a *hard* schedule — and that decision is made during the sanitize stage, which
overwrites the column in place with the single word `Easy`, `Hard` or `(n/a)`.
Nothing downstream of ingest can see a provider name, so nothing downstream can
publish one.

**The roster that drives the split is deliberately not in the code.**
`Build-Dashboard.ps1` is pushed to this public repo; a list of provider surnames
hardcoded in it would be published just as surely as if it were in `data.json`.
The patterns live in `providers.hard.txt`, which is **gitignored**, and the build
reads it at run time. The consequence is worth knowing: a fresh clone of this
repo has no roster, so it builds a single undivided `Appointment` type and says
so on the console. That is the intended degradation, not a bug — the published
artifact is identical either way apart from the two labels.

A PPI-free row-level copy is written to
`sanitized/<name>.sanitized.csv` so you can inspect or share the data without
touching the original export (which is never modified). It includes
`Insurance Portal` as-is, and includes the Easy/Hard verdict under the renamed
header `Schedule Class` — renamed because by that point the column holds a
verdict, not a provider, and a column labelled `Encounter Provider` containing
the word `Hard` invites someone to re-join it to a name later. Pass
`-SkipSanitizedCsv` to skip writing it. That folder is gitignored.

### Five safeguards. Do not disable them.

1. **Ingest sanitization**, above — identifiers are destroyed before the data
   is used for anything.
2. **`.gitignore` blocks `*.xlsx` / `*.csv` / `sanitized/` / `providers.hard.txt`**
   so neither the source export, a row-level copy, nor the provider roster can be
   committed by accident.
3. **The published label set is closed.** `Normalize-Reason` returns one of
   ~35 fixed category labels or the literal `Other / unclassified`. Input text
   is **never echoed through**, so a brand-new failure message carrying an
   identifier cannot reach `data.json` even in principle — it lands in
   `Other / unclassified` instead. And `data.json` holds aggregate counts only;
   there is no row-level field for an identifier to occupy.
4. **A publish gate** asserts every emitted label is in the closed set, then
   scans all labels for digit runs, ID phrasing, and any harvested coordinator
   or provider name. If anything suspicious survives, the build **deletes its own
   output and fails** rather than emit a publishable file. (Verified non-trivial:
   the gate rejects **766 of the 911** distinct raw reason strings in the current
   export, and **0** of the 40 published labels or 11 published portals.)
5. **A repo gate** scans the six publishable files — `README.md` and
   `Build-Dashboard.ps1` included — against the identifiers, coordinator names and
   89 provider names present in *this* export, plus the roster patterns
   themselves. Documenting a feature is the moment someone pastes a real value in
   as an example, and this gate is not theoretical: it caught a comment in
   `Build-Dashboard.ps1` that illustrated the hyphen-matching rule with a real
   provider's surname, and refused to finish the build until it was rewritten as a
   shape.

If `Other / unclassified` starts growing, that is the signal to add a rule to
`$REASON_RULES` in `Build-Dashboard.ps1` — the data is still safe, just less
informative. It currently holds **119 of 6,004** failure rows (2.0%).

---

## Refreshing the data

When a new export arrives:

1. Drop it in the project root. Run bare, the build picks up the newest file
   named `Orders_and_Appointments_Dashboard.xlsx` / `.xls` / `.csv`; any other
   name needs `-Source`. (`.xlsx` is read through Excel, `.csv` directly.)
2. Run the build:

   ```powershell
   .\Build-Dashboard.ps1 -Source "Orders_and_AppointmentsV1.07.csv"
   ```

   It prints what it redacted, which date column it bucketed on, and the
   headline metrics, so you can sanity-check all three before publishing:

   ```
   Insurance Portal: present
   Encounter Provider: present - splitting Appointment on 7 roster pattern(s)
   Sanitized: RC Assigned -> 13952 bot / 93458 human / 4 unassigned
              (67 coordinator names discarded)
   Sanitized: 840 fail reasons had names/identifiers redacted
   Sanitized: Encounter Provider -> 14739 hard / 46641 easy / 46034 n-a
              (89 provider names discarded)
     roster pattern #1: 2,086 rows
     roster pattern #2: 2,645 rows
     ...
   Date basis     : Date Auth Completed Roll Up
   Rows aggregated: 107393  (dropped for unparsable date: 0,
                             dated after the 2026-10-13 pull: 21)
   Publish gate: no identifier-like text in 40 reason categories / 11 portals.
   ```

   **Check the `Insurance Portal` line.** `absent` means the export predates
   V1.07 or the column was renamed; the build still succeeds and the dashboard
   simply hides the PORTAL filter rather than failing.

   **Check the roster pattern counts.** They are printed by position, never by
   name — line them up against `providers.hard.txt`. A pattern showing **0 rows**
   is the thing to act on: either that provider has left, or their name is spelled
   differently in this export, and in both cases their appointments are silently
   counting as *easy*. A pattern whose count looks far too large is the opposite
   failure — it's short enough to be matching providers you didn't intend, since
   matching is "contains", not "equals".

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
| `-NoEncrypt` | Write `data.json` as plaintext. **Local preview only** — see below. |

---

## The access gate

GitHub Pages has no server, so there is nothing to log in *to*. A JavaScript
password prompt alone would be decoration: `data.json` is fetched by the
browser, so anyone could skip the page and request
`…/Prior_Auth_Dashboard/data.json` directly.

So the file itself is the lock. `data.json` is published as **AES-256-GCM
ciphertext**; the passphrase derives the key in the browser via PBKDF2-SHA256
(310,000 iterations) and decrypts it there. The published file is unreadable
without the passphrase, and GitHub Pages is still just serving a static file.

### Setting the passphrase

Put it on the first line of `dashboard.passphrase.txt` beside the build script:

```powershell
Set-Content -LiteralPath .\dashboard.passphrase.txt -Value '<passphrase>' -NoNewline
```

The file is gitignored, same as `providers.hard.txt` — committing it would
publish the lock and the key together. Use **four or more random words**: the
ciphertext is downloadable, so a short passphrase can be attacked offline at
the attacker's own pace. PBKDF2 raises the cost per guess; it cannot rescue a
weak passphrase.

Without that file the build **fails** rather than falling back to plaintext.
That is deliberate: the portal column and the provider roster both degrade to a
working dashboard when absent, but degrading *this* would silently publish
every figure to a public URL. `-NoEncrypt` is the explicit opt-out, for local
preview — never push its output.

### Rotating it

Edit `dashboard.passphrase.txt`, re-run the build, publish the new
`data.json`. Everyone's saved session stops working on the next load and they
are prompted again. Note that rotation is **not retroactive**: a `data.json`
that was already published in the clear stays readable in git history and in
whatever caches picked it up.

### What this does and does not buy you

| | |
|---|---|
| ✅ | Keeps the aggregates off the open web and out of search indexes. |
| ✅ | Works inside a Notion embed, which real SSO does not (see below). |
| ✅ | No hosting change, no new service, no cost. |
| ❌ | One shared passphrase — no per-person access and no audit of who looked. |
| ❌ | No revocation for one person; rotating cuts off everyone at once. |
| ❌ | Not a substitute for the publish/repo gates. Those keep identifiers out of the aggregates; this only keeps the aggregates private. An identifier that slipped past them would be encrypted, published, and still a leak to anyone holding the passphrase. |

If you later need per-person access and an audit trail, that means moving off
GitHub Pages to a host with real auth (e.g. Cloudflare Pages + Cloudflare
Access) — and accepting that the Notion embed degrades, because an identity
provider's login page sets `X-Frame-Options` and cannot render in an iframe.

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

> Do **not** upload the `.xlsx` / `.csv` export, the `sanitized` folder, the
> `_shots` folder, `providers.hard.txt`, or `dashboard.passphrase.txt`.
> The last one is the key to `data.json` — uploading it alongside the
> encrypted file would undo the access gate entirely.

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

### Unlocking it inside Notion

Viewers get the passphrase prompt in the embed itself and tick **Remember on
this browser for 30 days**. This works in an iframe precisely because nothing
leaves the page — there is no redirect to an identity provider, and an IdP
login page could not render in an iframe anyway.

Two browser behaviours to expect, neither of which the page can change:

- **The 30 days are per browser, and Notion counts as its own context.**
  Browsers partition storage for embedded iframes, so unlocking inside Notion
  and unlocking the dashboard's own URL are two separate sessions. Someone who
  uses both will be asked twice.
- **Safari may not remember it at all.** Safari blocks storage for third-party
  iframes by default, so the tick box silently has no effect there and the
  prompt returns on each load. The dashboard still works — it just asks every
  time. Chrome and Edge persist it normally.

Make the embed block tall enough that the prompt isn't clipped; the gate needs
roughly 420px even though the dashboard itself wants 1,600–2,000px.

### Useful URL parameters

Append these to the embed URL to control the initial view:

| Parameter | Values | Effect |
|---|---|---|
| `?theme=` | `dark`, `light` | Pins the theme. Use `dark` to match a dark Notion workspace — the embed can't detect Notion's theme on its own. |
| `?range=` | `14d`, `1m`, `1q`, `ytd`, `all` | Sets the opening date range; the default without this parameter is `ytd`. Older values (`7d`, `30d`, `90d`, `mtd`) still work — they map to the nearest current preset, so existing embeds don't break. |
| `?type=` | `Easy Schedules`, `Hard Schedules`, `Diagnostic`, `Procedure`, `Referral`, `all` — or a comma-separated list. Shorthands: `easy`, `hard`, and `appointment`/`schedules` for both halves | Sets the opening Type filter. All three filters are multi-select, so `?type=Hard Schedules,Referral` opens with two Types included. `?type=appointment` predates the Easy/Hard split and still works — it selects both halves, so existing embeds are unchanged. A value this export doesn't contain is ignored and the default (both schedule types) stands. |
| `?status=` | `Auth Approved`, `Auth Not Required`, `Auth Denied`, `Auth Blocked`, `all` — or a comma-separated list | Sets the opening Status filter; the default is `all`. Spaces in a status name need URL-encoding, which the browser does for you if you type the URL in the address bar. |
| `?portal=` | Any portal in the export (`Aerial`, `IEHP`, `Medpoint`, …), `(Blank)`, `all` — or a comma-separated list | Sets the opening Portal filter; the default is `all`. Ignored on an export with no `Insurance Portal` column, where the filter doesn't exist. |

Combine with `&`:

```
https://nicholaswilcoxakido.github.io/Prior_Auth_Dashboard/?theme=dark&range=1q
```

A good pattern is several embeds on one Notion page, each pinned to a different
range — e.g. a 14-day operational view near the top and the default YTD trend
below. The 14-day view is also the one that shows ▲▼ deltas, which the YTD
default can't yet (see *Three decisions worth knowing about*).

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
   completion date in 2025, even though 10,661 rows were *created* between July
   and December 2025 — those all carry 2026 completion dates, decaying from 4,588
   in January to a few dozen by autumn. Whether that is a real cleared backlog or
   an artifact of when the field started being populated is a **Quickbase
   question, not a dashboard one**. Either way the x-axis cannot go back further,
   so `YTD` and `All` currently show the same window.
2. **The open backlog is invisible.** `Date Auth Completed Roll Up` is **100%
   populated** — zero blanks in 107,414 rows. A genuine completion field would be
   blank for anything still in flight, so in-progress authorizations either aren't
   in this export or aren't distinguishable in it. Nothing here can tell you how
   much work is currently open.
3. **About 6% of rows complete before they're created.** 6,587 rows have a
   completion date earlier than their creation date, which is impossible for a
   true timestamp pair. It doesn't affect any current metric — nothing computes a
   duration — but it does mean **turnaround time cannot be built on these two
   columns** until it's explained.

A fourth issue was found and handled: **21 rows carry completion dates after the
day the export was pulled**, the furthest 2027-03-30. Date presets anchor to the
latest date in the file, so that single 2027 row would have dragged the "14 Days"
window five months into empty calendar and opened the dashboard blank. The build
now drops anything dated past the export's own pull date and prints the count.

---

## Metric definitions

Scoped to the current Type, Status, Portal, and date filters. "In the period" means
**completed** in the period. Type is matched
case-insensitively, because the export mixes `Appointment` with `DIAGNOSTIC`
and `REFERRAL`; `Appointment` rows are then relabelled `Easy Schedules` or
`Hard Schedules` before anything is counted, so the Type dimension the dashboard
filters on is the post-split one. `RC Assigned` is tested with *contains* `svc-quickbase` against
the raw value, before sanitization collapses it — so the metrics are unaffected
by the redaction. (Confirmed: the headline figures are byte-identical with and
without the sanitize stage.)

| Metric | Definition |
|---|---|
| **Total Schedule Authorizations** | Count of authorizations **completed** in the period whose `Type` is a scheduling type — i.e. `Easy Schedules` + `Hard Schedules`, which is the same row set the single `Appointment` type used to cover. Tick only one half and the KPI label changes to *Total Easy/Hard Schedule Authorizations*. |
| **Touched by RPA** | **Distinct** rows where `RC Assigned` contains `svc-quickbase` **or** `Bot Run Fail Date` is populated. |
| **Percent Touched by RPA** | Touched ÷ Total. |
| **RPA Success Rate** | Rows where `RC Assigned` contains `svc-quickbase` ÷ Touched. |

### Three decisions worth knowing about

**Overlapping rows are counted once.** 469 scheduling rows (Easy + Hard) have *both* an
`svc-quickbase` assignment and a `Bot Run Fail Date`. Summing the two counts
would double-count them, so "touched" uses the de-duplicated union. These rows
are treated as **bot failed, then succeeded on retry** — they count as touched
*and* as successes. Across the full V1.07 export:

```
13,341  bot succeeded
   469  bot succeeded after a retry
 5,093  bot failed
------
18,903  touched by RPA   (not 19,372 — that figure double-counts the 469)
```

This makes all-time coverage 30.8% rather than 31.6%, and success rate 73.1%.
Note that all-time coverage is essentially **unchanged** by the switch to the
completion basis (30.81% → 30.81%), which is the expected result: re-dating rows
moves them between periods but doesn't add or remove any. The basis changes
*period* figures, not the grand total.

**A delta is only shown when the comparison period is fully present.** The
▲▼ figures compare against the preceding window of equal length. If that window
starts before the first row in the export, no delta is shown at all — an earlier
version compared YTD against a partly-empty 2025 window and reported "+401%
growth" that was really just missing history.

> Because the default preset is now **YTD** and the export starts Jan 1, 2026,
> the opening view has **no ▲▼ deltas** — there is no 2025 to compare against, and
> the tiles say so in place of a number. This is the safeguard working, not a
> bug. Deltas appear as soon as you pick a shorter preset, and the default will
> start showing them once the export holds more than a year of completions.
> If you want deltas visible on first load, embed with `?range=1q`.

**Date presets anchor to the latest date in the export, not to today.** A
"last 7 days" window measured from today would come up nearly empty whenever the
export is a few days stale. Anchoring to the data keeps the default view
populated. The exact window is always printed at the right of the filter row.

---

## Reading the dashboard

- **Filters** sit in one row and scope everything below them. Presets are
  **14 Days · 1M · 1Q · YTD · All**; the default is **YTD** with both scheduling
  types ticked (`Easy Schedules` + `Hard Schedules`, i.e. the old
  `Type = Appointment` slice). `1M` and `1Q` are inclusive calendar months (a `1M`
  window ending Oct 8 starts Sep 9), and `YTD` runs from Jan 1 of the latest
  year in the export. Anything past 70 days buckets by week, so the opening view
  is weekly rather than daily. **`YTD` and `All` currently show the same
  window**, because the earliest completion date in the export is Jan 1, 2026 —
  so today the default view is also the whole history. That will diverge in
  January, when `YTD` resets and `All` keeps 2026.
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
- **Hatched = incomplete period; don't read the dip.** Both time charts shade
  any end bucket that holds fewer days than it spans, and the lines run dashed
  through it with a hollow end marker. The subtitle spells out the shortfall
  (e.g. *"Oct 5–Oct 8, 4 of 7 days"*), the tooltip repeats it, and the `Table`
  view appends **(incomplete)** to that row. Two separate causes:
  - **The tail** — the export is a snapshot, so the day it was pulled is a
    part-day and the week containing it is a part-week.
  - **The head**, weekly views only — buckets are ISO weeks starting Monday, so
    `All` opens on Mon Dec 29 because that's the week holding Jan 1, and only
    four of its days have data.

  The counts shown are still the true counts for the days present; it's the
  *finality* that would be wrong to imply. A bucket that is low because the work
  was genuinely low is **not** shaded — only a calendar shortfall triggers it, so
  the shading stays meaningful.
- **Type, Status and Portal are multi-select.** Each one opens a panel of
  checkboxes — tick as many or as few values as you want, and every ticked value
  is included in every KPI and chart below. The button reports the selection
  (`Hard Schedules`, `Easy Schedules, Referral`, `3 of 5 types`), and
  **All types** / **All statuses** / **All portals** at the top of the panel
  ticks everything at once. Hovering a row reveals an **only** button, which
  isolates that one value in a single click — the fast path for "just Aerial",
  which would otherwise mean unticking ten boxes. The selection is also printed
  in the small line under the KPI row, because the filter row is hidden when the
  page is printed. One value always stays ticked: unticking the last one is a
  no-op, since an empty selection would draw an empty dashboard that reads as a
  broken build rather than as a filter choice.
- **`Easy Schedules` / `Hard Schedules` are the two halves of the old
  `Appointment` type.** An Appointment is a **hard schedule** when its
  `Encounter Provider` matches the roster in `providers.hard.txt`, and an **easy
  schedule** otherwise. Both are ticked by default, and together they are exactly
  the old `Appointment` total — 46,630 + 14,733 = 61,363 in the current export,
  with every headline figure (30.81% touched, 73.06% success) unchanged by the
  split. Untick one and the hero KPI renames itself to match, so a screenshot
  can't misrepresent which half it's showing.

  The two halves behave genuinely differently, which is the point of separating
  them: **easy** schedules are 33.7% RPA-touched at a 74.5% success rate,
  **hard** ones only 21.6% touched at 66.2%. On an export with no
  `Encounter Provider` column, or with no roster file, the type stays a single
  undivided `Appointment` and nothing else changes.
- **`Portal`** is `Insurance Portal`, the payer portal the auth was worked in.
  It appears only on exports from **V1.07** on; on an older one the whole
  control is hidden rather than shown empty. `(Blank)` is a real answer, not
  missing data — it's ~32% of rows, mostly `Auth Not Required` work that never
  needed a portal, and it holds **2** of the ~19.5k RPA-touched rows. Expect
  portal-level RPA coverage to look wildly different per portal: `Aerial` is
  about a third of all volume and carries most of the bot work, while `Citrus`
  has none at all.
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
