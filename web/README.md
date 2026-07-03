# LAPIS Mark — web version (skeleton)

A browser version of the LAPIS Mark grading app, so it can run on **iOS** (and any
device) via Safari/Chrome — no App Store. This folder is **fully independent** of the
Android app in `../app/`; they share no build files and their CI never overlaps.

## Status: marking loop (step 1 of the port)

Opens an `.xlsx` with [ExcelJS](https://github.com/exceljs/exceljs) (vendored in
`vendor/exceljs.min.js`, self-contained — no runtime CDN) and:

- recovers the full model — **students**, **criteria** (id, column, group, coefficient,
  contract, remarks), and **groups** — mirroring `XlsxParser.java`,
- shows **one criterion at a time** with the student name, group + `coeff. N`, the contract
  text, and a row of mark buttons (`× 0 ¼ ½ ¾ 1`),
- writes marks into the workbook **in memory**,
- navigates by criterion (`‹ critère / critère ›`) and student (`« étudiant / étudiant »`),
  both wrapping,
- **SAUVER** downloads a graded copy `<name>-graded.xlsx` (the original is never touched).

Against the reference workbook the parser recovers **159 students** and **65 criteria**.

### How the parsing works (validated against the reference workbook)

- Criterion metadata comes from the **`criteres_reviewed`** sheet (col B group with fill-down,
  C coefficient, D id, E contract, F remarks).
- On **`evaluation`**, the header row carries each criterion id per column (a formula result);
  `app.js` maps **column → id** by matching those against the known ids — robust and
  independent of fill detection.
- Mark cells sit at (student row, criterion column). They are a **theme fill** (theme 0,
  tint −0.15 → `#D9D9D9`), *not* a plain RGB — the earlier skeleton proved ExcelJS can read
  that via its `{theme, tint}` representation (the browser equivalent of POI's
  `getRGBWithTint`).

## Run it locally

It's a static site — no build step. Either open `index.html` directly, or serve the folder:

```
cd web
python3 -m http.server 8000   # then open http://localhost:8000
```

(A local server is closer to how it runs on Pages; opening the file directly also works.)

## Hosting

Deployed to **GitHub Pages** by `.github/workflows/web_build.yml` on every push to `main`
that touches `web/**`. Once live it will be at:

```
https://amsteingraphics.github.io/apk-excelparse/
```

The workflow asks GitHub to enable Pages automatically (`configure-pages` with
`enablement: true`). If the very first deploy fails with a Pages-not-enabled error, enable
it once by hand: repo **Settings → Pages → Build and deployment → Source = GitHub Actions**,
then re-run the workflow.

## Roadmap (port order)

1. ✅ Marking one student's criteria + Save = download the graded `.xlsx`.
2. ✅ Overview page — general + per-group averages (`CK`→`CL` math, **computed in JS**
   since ExcelJS has no formula engine) and canvas dot rows; GÉNÉRAL/CRITÈRES toggle.
3. ✅ Group filter (⋮ menu, per-workbook via `localStorage`) + Completion page with subpages.
4. ✅ Notes — editable box on each criterion page + a Notes carousel (⋮ menu), stored on a
   `notes` sheet at the same (row, col) as the mark cell (travels with the file, saved on SAUVER).
5. PWA offline (service worker) + home-screen icon.

