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
sanitized/*.sanitized.csv                 <- identifier-free row-level copy (local only)
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

**The two columns V1.07 added are treated differently from each other.**
`Insurance Portal` is published as a dashboard dimension — the values are payer
and vendor brand names (`Aerial`, `IEHP`, `Medpoint`, …), which carry nothing
about a member. Unlike failure reasons, that label set is **open**, because the
payer roster isn't ours to close, so the publish gate scans portal labels for
digit runs and coordinator names exactly as it scans the rest. Two of those
values are two names for one portal and publish merged — see
[Two names, one portal](#two-names-one-portal) below.
`Encounter Provider` holds provider **names**, and as of **V1.09 those names are
published** as the `Provider` filter dimension, at the repo owner's explicit
request.
This reverses the earlier rule and the reversal is worth stating plainly rather
than burying: through V1.08 the column was destroyed at ingest and nothing
downstream could see a provider name. It now serves two purposes — the Easy/Hard
verdict, as before, **and** a published dimension of 78 names plus `(Blank)`,
collapsed from the 93 distinct spellings the export actually contains.

Two consequences follow, neither of them a bug:

- **The hard-schedule roster becomes derivable.** It is still not written down in
  the repo or in `data.json`, but anyone with the dashboard can filter to one
  provider and read off whether their appointments land in `Easy Schedules` or
  `Hard Schedules`. Treat `providers.hard.txt` as obscured, not secret.
- **`RC Assigned` is unaffected.** Coordinator names are still destroyed at
  ingest. The authorization covers `Encounter Provider` only, and the publish
  gate still scans every other dimension's labels for coordinator names.

Provider names are **not** member data, which is why this is a different decision
from publishing anything out of `Bot Run Fail Reason`. They are still names, so
the label set gets the same digit-run, length and ID-phrasing checks as portals
(see the publish gate below).

How those names are de-duplicated before publishing is its own topic — see
[One clinician, two spellings](#one-clinician-two-spellings) below.

**V1.10 added two more columns, and both are now used.**
`# of times pending notes checked` is a plain integer count with nothing
identifying in it, so it is summed into the fact tuples and published as the
second KPI row — see
[Metric definitions](#metric-definitions). `Worked by
Automation` is a yes/no flag meaning the bot worked the row, and since the
V1.11-era rename it is the **third** way a row can count as attempted, on the
same optional terms as the pending-notes column: absent from the export, the
build says so and *Attempted* falls back to `RC Assigned` + `Bot Run Fail Date`.

It is a **successor field, not a cross-check** — it is blank for January through
September, covers 21.9% of October's attempts and 91.6% of November's, so its
reach grows as the column fills rather than describing a change in bot behaviour.
Do not read a jump in *Attempted* across that boundary as the bot doing more
work. It also turned out to be the only record of a whole program: of the 54 rows
that qualify as attempts **solely** through this flag, 44 are `DIAGNOSTIC` /
`REFERRAL` order types, where RPA coverage otherwise measures 1–3%. Without the
third clause the dashboard would have been blind to that automation entirely, not
merely 54 rows short.

**The roster that drives the split is deliberately not in the code.**
`Build-Dashboard.ps1` is pushed to this public repo; a list of provider surnames
hardcoded in it would be published just as surely as if it were in `data.json`.
The patterns live in `providers.hard.txt`, which is **gitignored**, and the build
reads it at run time. The consequence is worth knowing: a fresh clone of this
repo has no roster, so it builds a single undivided `Appointment` type and says
so on the console. That is the intended degradation, not a bug — the published
artifact is identical either way apart from the two labels.

An identifier-free row-level copy is written to
`sanitized/<name>.sanitized.csv` so you can inspect the data without touching the
original export (which is never modified). Member identifiers, authorization
numbers and coordinator names are gone from it. Provider names are **not** — it
carries them in a `Provider Name` column, and the Easy/Hard verdict separately
under `Schedule Class`, so the two sit side by side and the roster can be read
straight off that file. It is gitignored and it stays local; don't hand it out as
"the de-identified copy" the way you could through V1.08. Pass
`-SkipSanitizedCsv` to skip writing it.

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
   name. If anything suspicious survives, the build **deletes its own output and
   fails** rather than emit a publishable file. (Verified non-trivial: the gate
   rejects **833 of the 937** distinct raw reason strings in the current export,
   and **0** of the 40 published labels, 10 portals or 79 providers.) Since V1.09
   this is the *only* coordinator-name check that runs against `data.json`, for
   the reason in safeguard 5.
5. **A repo gate** scans the seven publishable files — `README.md` and
   `Build-Dashboard.ps1` included — against the identifiers, coordinator names and
   93 provider names present in *this* export — all the **raw** spellings, not
   just the 78 published after collapsing — plus the roster patterns
   themselves. `data.json` is exempt from the provider-name and coordinator-name
   passes, and only from those: it publishes 78 provider names by design, and the
   coordinator check is a blunt token match on single name parts ≥4 characters,
   which collided with three published provider names on the first V1.09 build —
   different people who happen to share a name part. Every other file, and the
   export-identifier pass on every file including `data.json`, is unchanged.
   Documenting a feature is the moment someone pastes a real value in
   as an example, and this gate is not theoretical: it caught a comment in
   `Build-Dashboard.ps1` that illustrated the hyphen-matching rule with a real
   provider's surname, and refused to finish the build until it was rewritten as a
   shape.

If `Other / unclassified` starts growing, that is the signal to add a rule to
`$REASON_RULES` in `Build-Dashboard.ps1` — the data is still safe, just less
informative. It currently holds **129 of 6,349** failure rows (2.0%).

---

## One clinician, two spellings

The export does not spell provider names consistently, and the pattern is not
random: in V1.09 **appointment rows carry a short form and order rows the full
legal name**, with zero overlap across all fourteen affected people. The same
clinician therefore arrived as two values and appeared **twice in the `Provider`
filter**, which is what `Get-ProviderAliasMap` in `Build-Dashboard.ps1` exists to
fix. It runs as its own pass before sanitize — picking which spelling to publish
needs to have seen every row, so a one-row-at-a-time normalizer cannot do it —
and collapses 93 spellings onto 78 published labels by two rules:

1. **Middle names.** Same first and last token, one spelling carrying a middle
   name or initial and the other not. `.` and `,` are stripped before comparing,
   so an initial written `R.` matches one written `R`. **14 collapsed in V1.09.**
2. **Short first names.** Same surname, where one first name is a strict prefix
   of the other — the nickname case, invisible to rule 1 because both spellings
   are two tokens. **1 collapsed in V1.09**, at 1 row against 5,782 for the
   spelling it folded into, with the same appointment-vs-order signature.

The published label is the **fewest-token** spelling, not the most frequent one.
That ordering matters: for two of the fourteen pairs the long form is the *more*
common, so choosing by count would publish some people with a middle initial and
others without — inconsistent in exactly the way this is meant to fix. Frequency
only breaks ties between spellings of equal length, which is what decides rule 2.

**Both rules refuse rather than guess.** Rule 1 leaves a first+last name alone if
it carries two *different* middle names: one person spelled two ways has at most
one middle form plus a bare version, and two rival middle forms is more likely two
clinicians who share a first and last name. Rule 2 refuses any surname where three
or more first names chain off one prefix, which are far more likely colleagues.
Each refusal prints a warning instead of merging.

Rule 2 is the weaker of the two and is kept separate in the code so it can be
switched off on its own. Its residual risk is a pair like `Ann`/`Anna` — same
surname, true prefix, two different people — which no guard here can detect, so
**every merge it makes prints at build time** and those lines are meant to be read
rather than trusted. Build output names nobody: merges and refusals both report
initials and row counts only, because that output gets pasted into this file and
into chat.

Two things the collapse deliberately does **not** touch:

- **The Easy/Hard roster match still runs against the raw value**, so the split
  cannot come to depend on the published label set. Safe here because the short
  spelling is the one that appears on `Appointment` rows in the first place — the
  V1.09 Easy/Hard figures are identical before and after the collapse.
- **The repo gate still scans against all 93 raw spellings**, not the 78 published
  ones, so a pre-collapse spelling pasted into a publishable file is still caught.

Consequences worth knowing: `?provider=` takes the collapsed label only, and
`sanitized/*.sanitized.csv` carries the collapsed label too, so cross-checking how
a name was originally written means going back to the export itself.

---

## Two names, one portal

`Insurance Portal` has the same duplicate-label problem as the provider column,
for a completely different reason, and so it gets a completely different fix.

`Curo` is the product and `Optum` its owner. Work done in either lands in the
same place, so publishing them as two filter entries splits one portal's numbers
across two rows for no reason. They publish as a single label, **`Optum Curo`**:

| Export value | Rows | Published as |
| --- | --- | --- |
| `Curo` | 4,394 | `Optum Curo` |
| `Optum` | 286 | `Optum Curo` |

That takes the portal dimension from 11 labels to 10 (9 named plus `(Blank)`),
and the merged label carries **4,665** rows in the dashboard — 4,680 minus the
15 `Curo` rows the [future-date guard](#the-date-basis--read-this-first) drops,
like any other row dated after the pull.

**The merge is an explicit table, not a rule.** This is the important difference
from the provider collapse, and it cuts the other way from what you might guess:

- Two spellings of a person's name *look alike*, so an algorithm can find them.
  The provider collapse exploits that, and because it is algorithmic it keeps
  working on names nobody has seen yet.
- Nothing about the strings `Curo` and `Optum` says they belong together. Only
  someone who knows the payer landscape does. There is no rule to generalize, so
  there is a hand-maintained table in `Build-Dashboard.ps1` (`$PORTAL_ALIAS`,
  keyed on the lowercased, whitespace-collapsed value) and nothing else.

The consequence is the thing to remember: **a future export inventing a third
spelling will not be folded in automatically.** It has to be added to the table
on purpose. Two console lines are there to make that visible rather than silent:

```
Portals        : 10 distinct value(s) -> 9 published label(s) plus (Blank)
  merged Curo (4,394 rows) + Optum (286 rows) -> 'Optum Curo'
```

and, if an alias key stops matching anything — the upstream-rename case, which
would otherwise just grow a filter entry nobody notices:

```
  WARNING: alias 'curo' matched no rows - renamed upstream? the merge is not happening
```

Three things the merge deliberately does **not** touch:

- **It moves rows between labels and nowhere else.** No count changes: summing
  every portal label still reproduces `rowsUsed` exactly, which is the check
  that would catch a row landing under both the old and the new label.
  `Distinct cells` does drop (50,923 → 50,830) because cells that differed only
  by portal now coincide — that is the merge working, not data going missing.
- **`sanitized/*.sanitized.csv` keeps the raw spellings.** Same as the provider
  collapse: the sanitized copy mirrors the export, not the published labels.
- **The repo gate still scans the raw values.** A pre-merge spelling pasted into
  a publishable file is caught exactly as before.

One place the merge *is* visible beyond the filter: `?portal=Curo` and
`?portal=Optum` are kept as aliases in the dashboard so links written before the
merge still resolve to `Optum Curo`. Without that they would match no label, and
an unmatched `?portal=` is **ignored** rather than treated as an empty selection
— the link would look like it worked while showing every portal.

---

## Refreshing the data

When a new export arrives:

1. Drop it in the project root. Run bare, the build picks up the newest file
   named `Orders_and_Appointments_Dashboard.xlsx` / `.xls` / `.csv`; any other
   name needs `-Source`. (`.xlsx` is read through Excel, `.csv` directly.)
2. Run the build:

   ```powershell
   .\Build-Dashboard.ps1 -Source "Orders_and_AppointmentsV1.10.csv"
   ```

   It prints what it redacted, which date column it bucketed on, and the
   headline metrics, so you can sanity-check all three before publishing:

   ```
   Insurance Portal: present
   Pending notes column: present
   Encounter Provider: present - splitting Appointment on 7 roster pattern(s)
   Provider names: 93 distinct spelling(s) -> 78 published label(s);
                   14 middle-name and 1 short-first-name variant(s) collapsed
     short-first-name merge: CH, 1 row(s), folded into the more frequent
       spelling of the same surname - verify this is one person, not two
   Sanitized: RC Assigned -> 15720 bot / 99300 human / 0 unassigned
              (69 coordinator names discarded)
   Sanitized: 867 fail reasons had names/identifiers redacted
   Sanitized: Encounter Provider -> 15615 hard / 50206 easy / 49199 n-a
              (78 provider names PUBLISHED as the Provider dimension)
     roster pattern #1: 2,315 rows
     roster pattern #2: 2,786 rows
     ...
   Portals        : 10 distinct value(s) -> 9 published label(s) plus (Blank)
     merged Curo (4,394 rows) + Optum (286 rows) -> 'Optum Curo'
   Date basis     : Date Auth Completed Roll Up
   Rows aggregated: 114987  (dropped for unparsable date: 0,
                             dated after the 2026-11-03 pull: 33)
   Distinct cells : 50830
   Pending notes  : 145 checks over 22 row(s) = 145 min (2.4 h) saved
     18 of those sit on 6 row(s) not completed as of the pull; counted anyway,
       dated 2026-11-03
     5 zero-count cell(s) added to carry them - they move no other figure
     matches the raw column sum of 145 exactly
   Publish gate: no identifier-like text in 40 reason categories / 10 portals
                 / 79 providers.
   Repo gate: 7 files clean against 795 export identifiers / 44 name forms
              / 93 provider names (data.json exempt from the provider scan).
   ```

   **Read the `Provider names:` lines on every refresh.** A new export can spell
   a name a third way, or introduce a genuine same-surname collision, and these
   are the only lines that will tell you. A `NOT collapsed:` warning means the
   build found two people it refuses to merge on its own — go and look.

   **The word `PUBLISHED` on the `Encounter Provider` line is load-bearing.** If
   it ever reads `discarded` instead, the build is running pre-V1.09 code and the
   `Provider` filter will be missing from the dashboard.

   **Check the `Insurance Portal` line.** `absent` means the export predates
   V1.07 or the column was renamed; the build still succeeds and the dashboard
   simply hides the PORTAL filter rather than failing.

   **Check the `Portals` lines too.** The `merged … -> '…'` line is the receipt
   for the alias table described in [Two names, one portal](#two-names-one-portal);
   if a merge you expect isn't listed, it isn't happening. A red
   `WARNING: alias 'x' matched no rows` means the export renamed a value the
   table still looks for — the portal has silently un-merged and the dashboard
   will just grow an extra filter entry, which is easy to miss. Portal values
   are brand names, so unlike the provider lines these print in full.

   **Check the `Pending notes` lines.** `absent` on the column line means the
   export predates V1.10 or the column was renamed — the build still succeeds and
   the dashboard hides the second KPI row rather than showing two zeros.

   The line that matters most is the last one. **`matches the raw column sum of N
   exactly`** is the build asserting that the number on the tile equals what you
   would get by summing the column in Excel. If it ever prints
   **`WARNING: … unaccounted for`** in red instead, the tile is understating the
   column — find out why before publishing. See
   *Why this one metric ignores the future-date guard* below.

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

   **Expect `data.json` to be noticeably larger since V1.09.** Adding a 79-value
   dimension multiplies the number of distinct cells, so the file went from
   ~489 KB to **~1.4 MB** encrypted. That is a one-time step change, not drift.
   It is the one real cost of the `Provider` filter: the Notion embed takes
   correspondingly longer to download and decrypt before the first paint.

   V1.10's pending-notes total costs almost nothing on top of that, because the
   build appends it **per cell and only when it is non-zero** — 18 of 50,923
   cells in the current export. Writing a `0` into the other 50,905 would have
   added roughly 100 KB of plaintext to a payload that already strains the Notion
   embed. The dashboard reads the slot as "absent means zero", which is also the
   correct reading of a pre-V1.10 `data.json`.

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

   > **The V1.10 release is the exception: push both files.** The second KPI row
   > is new markup, so a V1.10 `data.json` beside the old `index.html` shows the
   > four original tiles and silently drops the pending-notes total. The pairing
   > is safe in both directions — the new dimension was appended *after* the count
   > in each fact tuple precisely so a stale pairing degrades instead of
   > misreading an index — but it is not what you want on screen.
   >
   > **The V1.11-era metric rename is the same kind of exception: push both
   > files.** It renamed every RPA tile and buckets are written by the *build*,
   > so the new labels need the new `data.json` behind them to be true.
   >
   > **Bucket 4 is retired, and the slot is still reserved.** The rename briefly
   > carried a fifth bucket (`bot_worked` = 4) for rows that are attempts only by
   > `Worked by Automation`; those rows now go straight into `bot_fail` =
   > *attempted, not submitted*, and nothing emits a `4` any more. `index.html`
   > still **reads** slot 4 and folds it into `fail`, because Pages serves the
   > previous `data.json` for ~10 minutes after a push and that file does hold
   > `4`s — folding them means the page reports the current rule either way.
   > Two rules follow from this and neither is optional:
   > - **Keep every bucket array in `index.html` sized to five slots.** Size one
   >   to four and `b[4]` is `undefined`, which turns `total` into `NaN` and takes
   >   the whole page down for exactly that stale window.
   > - **Never reuse index 4 for a new meaning.** During that window the new
   >   meaning would be applied to old rows, silently.

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
| `?range=` | `14d`, `1m`, `1q`, `ytd`, `all` | Sets the opening date range; the default without this parameter is `1m`. Older values (`7d`, `30d`, `90d`, `mtd`) still work — they map to the nearest current preset, so existing embeds don't break. `mtd` maps to `1m`, which is also now the default. |
| `?type=` | `Easy Schedules`, `Hard Schedules`, `Diagnostic`, `Procedure`, `Referral`, `Immunization`, `all` — or a comma-separated list. Shorthands: `easy`, `hard`, and `appointment`/`schedules` for both halves | Sets the opening Type filter. All four filters are multi-select, so `?type=Hard Schedules,Referral` opens with two Types included. `?type=appointment` predates the Easy/Hard split and still works — it selects both halves, so existing embeds are unchanged. A value this export doesn't contain is ignored and the default (both schedule types) stands. |
| `?status=` | `Auth Approved`, `Auth Partially Approved`, `Auth Not Required`, `Auth Denied`, `Auth Blocked`, `all` — or a comma-separated list | Sets the opening Status filter; the default is `all`. Spaces in a status name need URL-encoding, which the browser does for you if you type the URL in the address bar. `Auth Partially Approved` appears from the V1.11-era exports on; it is its **own** outcome, not a variant of `Auth Approved`, because a partial approval means some requested units or services were granted and others were not — folding the two together would overstate approvals. A value this export doesn't contain is ignored, so a link carrying it against an older `data.json` falls back to `all`. |
| `?portal=` | Any published portal label (`Aerial`, `IEHP`, `Medpoint`, `Optum Curo`, …), `(Blank)`, `all` — or a comma-separated list. `Curo` and `Optum` are kept as aliases for `Optum Curo` | Sets the opening Portal filter; the default is `all`. An unrecognized value is **ignored**, leaving `all` — so a typo shows every portal rather than an empty dashboard. That is why the two pre-merge spellings are aliased rather than dropped (see [Two names, one portal](#two-names-one-portal)): an old `?portal=Curo` link would otherwise look fine while quietly showing everything. Ignored on an export with no `Insurance Portal` column, where the filter doesn't exist. |
| `?provider=` | Any **published** provider label, `(Blank)`, `all` — or a comma-separated list | Sets the opening Provider filter; the default is `all`. Names need URL-encoding for their spaces. Must be the collapsed label, not a middle-name spelling the build merged away — those are no longer in the payload and are ignored, exactly like a name this export doesn't contain, so an embed pinned to a provider who has since left falls back to `all` rather than showing an empty dashboard. Ignored on an export with no `Encounter Provider` column. |

Combine with `&`:

```
https://nicholaswilcoxakido.github.io/Prior_Auth_Dashboard/?theme=dark&range=1q
```

A good pattern is several embeds on one Notion page, each pinned to a different
range — e.g. the default 1M operational view near the top and `?range=ytd` or
`?range=all` below it for the full-year ramp. Pin the long view explicitly: it is
no longer what you get for free, and it is the one place the year-to-date climb
from 5.4% to 40.2% is visible.

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
| **Attempted to Submit by RPA** | **Distinct** rows where `RC Assigned` contains `svc-quickbase` **or** `Bot Run Fail Date` is populated **or** `Worked by Automation` is checked. The clauses are evaluated in that order, so a row with a real outcome is never reclassified by the bare flag. |
| **Successfully Submitted by RPA** | Rows where `RC Assigned` contains `svc-quickbase` — buckets `bot_ok` + `bot_retry_ok`, so a run that failed once and succeeded on retry counts here. |
| **Attempted, Not Submitted** | Bucket `bot_fail`: a `Bot Run Fail Date`, **or** the `Worked by Automation` flag with no bot assignment and no fail date. The second kind has no recorded outcome, and the business rule reads it as one — automation worked it, we are no longer assigned to it, so it was attempted and not submitted. Those rows carry **no fail reason**, so they never appear in the Top Bot Failure Reasons chart; the build prints their count on every run so the clause stays visible. |
| **Percent Attempted to Submit by RPA** | Attempted ÷ Total. *Coverage* — how much of the work the bot reached at all. |
| **Percent of Total Authorizations Successfully Submitted by RPA** | Successfully Submitted ÷ **Total**. The program's goal metric, and the product of the two ratios either side of it. |
| **Success Rate of Authorizations Attempted by RPA** | Successfully Submitted ÷ **Attempted**. *Conversion* — how well the bot did once it got there. |

**Most of these are renames, not new measurements** — the exceptions are the
third clause on *Attempted*, and the *Successfully Submitted* and *Percent of
Total … Successfully Submitted* tiles. Both of those are figures that already
existed inside the other ratios and are now shown on their own; neither needs
anything new in `data.json`, because both are computed from the same facts in
`index.html`, so they appear the moment the new page loads whichever `data.json`
Pages happens to be serving. The old names were
`Touched by RPA`, `Percent Touched by RPA` and `RPA Success Rate`; "touched" and
"worked" did not survive contact with a reader who had not built the thing, which
is the whole reason for the rename. The internal field names in `data.json` and
in `index.html` (`touched`, `pctTouched`) deliberately **did not** change: they
are the stale-pairing contract between a new `index.html` and an old `data.json`,
and renaming them would buy nothing but a window in which the page reads
`undefined`.

Note what the success-rate denominator is: **Attempted** — every row the bot
touched, including the ones whose outcome the export never recorded. That is
deliberate. The tile says "of authorizations attempted by RPA", so dividing by a
quietly smaller number would make its own label false. The same logic is why a
`Worked by Automation`-only row sits in *Attempted, Not Submitted* rather than in
a bucket of its own: it is an attempt, and no submission is on record, so it
belongs on the failure side of the ratio like any other.

Also note what does **not** get credit: a deliberate handoff (`Routed to UM
Workflow`), a referral already submitted, a prior authorization already on file,
and a cancellation in IEHP all stay in the failure side of the ratio. The bot
attempted and no submission resulted; *why* it did not is a separate question
from *whether* it did, and the Top Bot Failure Reasons chart is where the why
lives. This was considered and explicitly ruled on — moving those four
categories out of the denominator would have lifted the rate by 2.17pp.

*Attempted, Not Submitted* is the one row above that is not its own tile — it is
the second half of the *Attempted* tile's sub-line, the orange segment on the
volume chart, and the failure side of the success rate.

**The three ratios are not interchangeable, and the middle one is the only
program-level number.** Coverage × conversion = percent of total submitted, so
quoting any two fixes the third. What makes the middle one different is that it
is the only one that can move on work RPA never touched: close out a slice the
bot does not reach and it falls while the other two sit still. That is the
property a target needs and the property a bot scorecard must not have — judge
the bot on *Success Rate*, judge the program on *Percent of Total*.

A second row of two tiles sits below the six KPI tiles, added in V1.10:

| Metric | Definition |
|---|---|
| **Pending Notes Checked** | **Sum** of the `# of times pending notes checked` column — not a count of rows. The only measure on the page that adds up a column rather than counting authorizations, and the only one that counts work on authorizations that have not completed yet. |
| **Time Saved Checking Pending Notes** | The same number shown as hours and minutes, at **one minute per check**. 145 checks reads `2h 25m`. |

Both are scoped by the filters exactly like the four above them, and both read
`N/A` when a dimension has nothing ticked. The whole row is **hidden** on an
export with no `# of times pending notes checked` column, which includes every
`data.json` built before V1.10 — two tiles reading 0 would be a claim the data
can't support, so the build publishes a `hasPending` flag rather than letting the
dashboard infer it (an all-zero column and a missing column are not the same
statement).

**This column is almost entirely empty, and that is what the data says.** In
V1.10, 22 rows of 115,020 carry a value — 99.98% blank. The tiles will read small
numbers for a while; they are not broken.

**It is the one metric on the page that deliberately ignores the future-date
guard** — see
[Why this one metric ignores the future-date guard](#why-this-one-metric-ignores-the-future-date-guard).

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

> **The default preset is `1m`, and one reason is this safeguard.** Under the
> previous `YTD` default the opening view had **no ▲▼ deltas at all** — the export
> starts Jan 1, 2026, so there was no 2025 to compare against and every tile said
> so in place of a number. A 31-day window has its prior period entirely inside
> the export, so the deltas work on first load. `YTD` and `All` still show no
> deltas when picked, which remains the safeguard working rather than a bug.

**Date presets anchor to the latest date in the export, not to today.** A
"last 7 days" window measured from today would come up nearly empty whenever the
export is a few days stale. Anchoring to the data keeps the default view
populated. The exact window is always printed at the right of the filter row.

### Why this one metric ignores the future-date guard

Every other number on the page counts authorizations **completed** in the window,
so a row whose completion date is *after the export was pulled* has no window to
land in. The build drops 33 such rows in V1.10. That guard is not optional: the
date presets anchor to the latest date in the file, so one stray 2027 row drags
the "14 Days" window into empty calendar and the dashboard opens blank.

Six of those dropped rows carry pending-notes checks, worth **18** of the
column's **145**. For this metric the guard is wrong, and the reason is that
`Pending Notes Checked` is not a claim about a completion at all — the checks
already happened, they are sitting in the export, and whether the authorization
has since closed is beside the point. So:

- The authorization stays **out of every count**, exactly as before. `Rows
  aggregated`, the Total tile, the trend charts and the Easy/Hard split are all
  byte-identical to a build that discarded those rows outright.
- Its **checks are kept**, and re-dated to the last day of data so they sit
  inside the window every preset covers. The published total therefore equals the
  raw column sum: **145**.

The carrier is a cell with a **count of zero** — a fact tuple holding pending
minutes and no authorizations. Three things make that safe, and all three are
asserted by the headless probe (`_shots\probe_pending.ps1`, stage 2b):

1. It lands on the **last existing date**, never past it, so the date axis is not
   stretched and the guard's original purpose is intact.
2. Its reason index is forced to `-1`. A failure reason describes a counted
   authorization, and this cell counts none — which also keeps it out of the
   failure-reason chart, whose category list is built from the facts it is given.
3. Type and Status charts sum counts, so a zero adds nothing. No `0` row appears
   in any chart or table.

The build re-sums the raw column independently and **asserts the two agree** on
every refresh:

```
Pending notes  : 145 checks over 22 row(s) = 145 min (2.4 h) saved
  18 of those sit on 6 row(s) not completed as of the pull; counted anyway,
    dated 2026-11-03
  5 zero-count cell(s) added to carry them - they move no other figure
  matches the raw column sum of 145 exactly
```

A red `WARNING: … unaccounted for` in place of that last line means something is
eating checks; don't publish until you know what.

**The one case where the tile still won't match Excel** is a date filter that
excludes the last day of data — the clamped 18 live on that day, so a custom
range ending earlier won't see them. On every built-in preset they are included.

---

## Reading the dashboard

- **Filters** sit in one row and scope everything below them. Presets are
  **14 Days · 1M · 1Q · YTD · All**; the default is **1M** with both scheduling
  types ticked (`Easy Schedules` + `Hard Schedules`, i.e. the old
  `Type = Appointment` slice). `1M` and `1Q` are inclusive calendar months (a `1M`
  window ending Oct 8 starts Sep 9), and `YTD` runs from Jan 1 of the latest
  year in the export. Anything past 70 days buckets by week, so the default view
  is **daily** at 31 days and `1Q`/`YTD`/`All` are weekly. **`YTD` and `All`
  currently show the same window**, because the earliest completion date in the
  export is Jan 1, 2026. That will diverge in January, when `YTD` resets and
  `All` keeps 2026.
- **Why the default is a month and not the year.** Every rate on the page is a
  period average, and this program ramped hard inside the period: percent of
  total submitted ran 5.4% in January and 40.2% in October. `YTD` therefore
  reported **23.7%** against a current run rate near 40%, understating the
  program by about 15pp on its own goal tile. `1M` was chosen over `MTD` and
  `14 Days` on measured grounds — it keeps daily bars, its prior period fits
  inside the export so the deltas are real, and because presets anchor to
  `dateMax` it is always the trailing ~31 days, where `MTD` would collapse to a
  single part-day on a first-of-month export. The trade is that the default
  window is the one still filling in; see *The date basis*.
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
- **Type, Status, Portal and Provider are multi-select.** Each one opens a panel
  of checkboxes — tick as many or as few values as you want, and every ticked
  value is included in every KPI and chart below. The button reports the
  selection (`Hard Schedules`, `Easy Schedules, Referral`, `3 of 6 types`). The
  **All …** box at the top of each panel is tri-state: ticked means everything,
  dashed means a subset, and **unticking it clears the whole dimension** — the
  same control selects and deselects. Hovering a row reveals an **only** button,
  which isolates that one value in a single click — the fast path for "just
  Aerial", which would otherwise mean unticking ten boxes. The selection is also
  printed in the small line under the KPI row, because the filter row is hidden
  when the page is printed.
- **A dimension with nothing ticked reads `N/A`, not `0`.** Unticking every value
  is allowed as of V1.09 (it used to be a no-op). When it happens, **every** KPI
  tile reads `N/A` — both rows — their deltas and sparklines come off, and every chart and
  table says which dimension is empty and how to fix it. `N/A` means "you asked
  about no authorizations"; a genuine `0` still reads `0`, and the two are
  deliberately different — so a provider with 15 referrals and no appointments
  shows `0` under the default Type filter, not `N/A`.
- **`Easy Schedules` / `Hard Schedules` are the two halves of the old
  `Appointment` type.** An Appointment is a **hard schedule** when its
  `Encounter Provider` matches the roster in `providers.hard.txt`, and an **easy
  schedule** otherwise. Both are ticked by default, and together they are exactly
  the old `Appointment` total — 49,377 + 15,465 = 64,842 in the current export,
  with every headline figure (31.66% touched, 73.70% success) unchanged by the
  split. Untick one and the hero KPI renames itself to match, so a screenshot
  can't misrepresent which half it's showing.

  The two halves behave genuinely differently, which is the point of separating
  them: **easy** schedules are 34.5% RPA-touched at a 75.4% success rate,
  **hard** ones only 22.6% touched at 65.6%. On an export with no
  `Encounter Provider` column, or with no roster file, the type stays a single
  undivided `Appointment` and nothing else changes.

  Since V1.09 the roster is no longer hidden in practice: pick a single value in
  the `Provider` filter and the Type breakdown tells you which half that provider
  is in. See *PPI is removed at ingest* above.
- **`Portal`** is `Insurance Portal`, the payer portal the auth was worked in.
  It appears only on exports from **V1.07** on; on an older one the whole
  control is hidden rather than shown empty. `(Blank)` is a real answer, not
  missing data — it's ~22% of rows, mostly `Auth Not Required` work that never
  needed a portal, and it holds **1** of the ~21.6k RPA-touched rows. Expect
  portal-level RPA coverage to look wildly different per portal: `Aerial` is
  about a third of all volume and carries **58%** of all bot work at a 34% touch
  rate, while `Medicare` is ~10% of volume and under **1%** touched. Nine labels
  plus `(Blank)` in the current export, and one of those nine is two export
  values merged — see [Two names, one portal](#two-names-one-portal).
- **`Provider`** is `Encounter Provider`, the ordering or seeing provider on the
  row — **real names**, 78 of them in the current export, added in **V1.09**.
  Each clinician appears **once**: the export spells the same person two ways
  (short on appointment rows, full legal name on order rows) and the build
  collapses 93 spellings onto those 78 labels — see
  [One clinician, two spellings](#one-clinician-two-spellings) for which variants
  merge and which the build refuses to merge.
  Like Portal, the control is hidden entirely on an export that lacks the column.
  The panel has a **search box** because the list is too long to scan; typing
  narrows what's visible and never changes what's ticked, so searching can't
  silently drop a provider from the figures. `(Blank)` is **25% of rows** and is
  a real answer — order types that carry no provider at all — so leaving it
  ticked is usually what you want. Note that a provider's rows are spread across
  *all* Types, not just the two schedule types: with the default Type filter a
  referral-only provider shows `0`.
- **The second KPI row counts pending-note checks, not authorizations.** It is
  the one place on the page that sums a column instead of counting rows, and it
  appears only on exports from **V1.10** on — older ones hide the row rather than
  show two zeros. The column is 99.98% blank today, so expect small numbers.
  `Time Saved Checking Pending Notes` is the same figure in hours and minutes at
  one minute per check, so the two tiles cannot disagree. It is also the **only**
  row on the page that counts work on authorizations which have not completed
  yet, so it will sum to more than the completed rows alone would suggest — see
  [Why this one metric ignores the future-date guard](#why-this-one-metric-ignores-the-future-date-guard).
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
2. **Hours saved on the authorizations themselves** — successfully-submitted × an
   agreed minutes-per-auth figure. Needs one number from you to be credible.
   V1.10 delivered the *pending-notes* half of this at one minute per check; the
   much larger per-authorization figure is still unagreed.
3. **Get an outcome recorded for the `Worked by Automation`-only rows.** ~~Use
   `Worked by Automation` as a cross-check against `touched`.~~ **Done, and the
   premise was wrong** — it is not a cross-check, it is a successor field (blank
   before October; see *The columns*), so a disagreement between the two is
   mostly just the column's fill rate. It is now the third clause of *Attempted
   to Submit by RPA*. What remains is the rows that carry the flag and nothing
   else: no bot assignment, no failure date, so no record of how the run ended.
   By business rule they count as *attempted, not submitted* — automation worked
   them and we are no longer assigned to them — so there is no longer a separate
   bucket or chart segment to watch. **The fix still belongs upstream in the bot,
   not here:** have it write either an `RC Assigned` value or a
   `Bot Run Fail Date` on every row it touches, because today a genuine success
   that forgot to set `RC Assigned` would be booked as a non-submission. The
   build prints the count of these rows on every run (*"Of those, no fail reason
   — Worked by Automation only"*); if it climbs faster than the column's fill
   rate, this moves up the list.
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
| `Build-Standalone.ps1` | Export → one self-contained plaintext HTML file in `standalone/`. No server, no passphrase. |
| `sanitized/` | Row-level copies of each export with identifiers and coordinator names removed — but **with** provider names, as the collapsed published label rather than the raw spelling. Go back to the export itself to see how a name was originally written. Local only, gitignored. |
| `standalone/` | Plaintext single-file builds. Local only, gitignored. |
| `.gitignore` | Blocks source exports, sanitized copies, standalone builds, the roster and the passphrase from being committed. |

No build toolchain, package manager, or CDN dependency — `docs/index.html` is
plain HTML, CSS, and vanilla JS, and the charts are hand-rolled SVG. It will
still work years from now.
