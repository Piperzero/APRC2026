# APRC 2026 Deposit Coverage Calculator

A deposit-insurance coverage calculator and behavioural survey built by PIDM
(Perbadanan Insurans Deposit Malaysia) for the APRC 2026 conference.
Users enter a deposit amount, compare coverage across 22 Asia-Pacific
jurisdictions, explore a family/joint-account scenario, and answer a
3-question behavioural survey.

**Status: mothballed (June 2026).** The event has concluded, survey data
collection is closed, and the webhook is decommissioned. The tool is kept
live for reference and will be revived for a future event.

## Current state

| Item | Value |
|---|---|
| **Live / canonical version** | [`V98-APRC.html`](https://piperzero.github.io/APRC2026/V98-APRC.html) |
| `V99-APRC.html` | Parked — exploratory features from APRC 2026 event feedback (dynamic regional-grid interaction). Not deployed; do not link to it. |
| `V96` / `V97` | Retired — now redirect to V98. |
| `index.html` | Redirects the site root to V98. |
| Survey webhook | **Disabled** (Apps Script deployment archived June 2026; URL returns 404). |
| Survey data | Collected ~May 5–6, 2026; exported from the Google Sheet and archived. |
| Korea (KDIC) limit | KRW 100,000,000 (raised Sept 2025 — already reflected in V98). |

## Repository contents

- `V98-APRC.html` — the entire app: single file, zero dependencies (HTML + CSS + vanilla JS). Country data and exchange rates are inline in the `<script>` block (search for `const countries` and `const exchangeRates`).
- `AppScript_APRC2026.gs` — source of the Google Apps Script webhook that received survey submissions. The deployed script lives in a personal Google account (see Ops notes); this file is the backup copy.
- `build_findings.py` — one-off analysis script (pandas/openpyxl) that turned the exported survey Sheet into a formatted findings workbook. Hardcoded to the May 2026 export path and dates; parameterize before reuse.
- `*.png` — logo and image assets.

## Ops notes

- **Hosting:** GitHub Pages from `main`. Deploying = pushing to `main`.
- **Apps Script ownership:** the webhook script and response Sheet are owned
  solely by a personal Google account (anaz.shazlan@gmail.com). The `.gs`
  file in this repo is the only copy outside Google. Before the next event,
  consider migrating to an organizational account.
- **Disclaimer date coupling:** the legal disclaimer on the landing page
  states the exchange-rate "as of" date. If you update `exchangeRates`,
  update the disclaimer text in the same commit (search the HTML for
  `as of`).

## Revival checklist (for the next event)

1. Branch from `main`; copy `V98-APRC.html` to the new version file.
2. Refresh all 22 entries in `exchangeRates` and verify every `limit` in
   `countries` against the respective deposit insurer / IADI; update the
   disclaimer "as of" date in the same commit.
3. Consolidate `countries` + `exchangeRates` + a single `RATES_AS_OF`
   constant into one clearly fenced block, and render the disclaimer date
   from it (removes the manual coupling above).
4. Remove known dead code: two `if (false)` Myanmar blocks, the empty
   `renderCountryGrid()`, the unreachable `case 5` in
   `validateCurrentPage()`, and the duplicated webhook URL constant.
5. Add a small math self-check (≈10 assertions: limit conversion both
   directions, fully/partially/exactly-covered boundaries, joint-split sums,
   every country has a rate) and run it after any data edit.
6. Re-enable data collection with a **hardened** Apps Script: shared token,
   field allow-lists, length caps, `'`-prefix for values starting with
   `= + - @` (Sheets formula injection), and `LockService` around the
   interaction counter. Redeploy and update the webhook URL constant in the
   HTML.
7. Decide the fate of `V99`'s parked features (dynamic regional-grid
   interaction) — merge or drop.
8. End-to-end test on desktop and mobile, then deploy and update this README.
