# LAPIS Mark — web version (skeleton)

A browser version of the LAPIS Mark grading app, so it can run on **iOS** (and any
device) via Safari/Chrome — no App Store. This folder is **fully independent** of the
Android app in `../app/`; they share no build files and their CI never overlaps.

## Status: skeleton / proof of concept

Right now this only proves the risky part works in a browser: **open an `.xlsx` and read
its structure, including cell fill colours.** It:

- opens a picked `.xlsx` with [ExcelJS](https://github.com/exceljs/exceljs) (vendored in
  `vendor/exceljs.min.js`, so the page is self-contained — no runtime CDN),
- finds the `evaluation` sheet,
- detects **mark columns** by the gray mark-cell fill — which is actually a **theme fill**
  (theme 0, tint −0.15) that resolves to `#D9D9D9`, not a plain RGB (see below),
- lists the **students** (column A numeric, column B name),
- **never modifies the file.**

Against the reference workbook this reports **159 students** and **65 mark columns**.

It mirrors a simplified slice of the Android `XlsxParser.java`. The key thing this skeleton
proves: **theme+tint fills can be read in the browser.** ExcelJS reports theme fills as
`{theme, tint}` (not a resolved RGB), so `app.js` matches the mark cell by that signature —
the browser equivalent of POI's `getRGBWithTint`. Criterion *label* cells (an Accent-1
theme+tint fill) are detected the same way in the next step.

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

1. Marking one student's criteria (mark buttons writing back into the workbook in memory).
2. Overview + group/general averages (`CK`→`CL` math).
3. Group filter + completion.
4. Notes (the `notes` sheet).
5. Save = download the graded `.xlsx`; optional PWA offline + home-screen icon.
