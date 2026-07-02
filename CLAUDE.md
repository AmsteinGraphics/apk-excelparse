# apk-excelparse — context for future Claude Code sessions

This file exists so a Claude Code session on any machine can pick up the project without re-deriving decisions or repeating past mistakes. Update it when you make a non-obvious change; delete stale entries.

## What the app does

Personal Android grading app for a Swiss teacher. Reads student rows and per-criterion columns from an `.xlsx` file, presents one criterion at a time on a phone-sized screen with a row of five mark buttons (0 / ¼ / ½ / ¾ / 1) that write the mark back into the workbook. The original picked file is never modified — a working copy lives in app-private storage, and each Save mirrors it to `Downloads/{basename}-graded.xlsx` via MediaStore.

## Non-obvious architecture

- **Cells are identified by fill color, not by fixed ranges.** Merged cells and group-average columns in the workbook mean row/column indices aren't stable. The parser scans for two colors on the `evaluation` sheet:
  - Criterion label cells (POI-resolved RGB `#E6E9EB`, ±4 tolerance). Excel *displays* these as pale blue `#DAE9F8`; POI's HSL-based tint math on the Accent-1 theme color resolves to that near-neutral. Trust POI's number, not Excel's swatch. Always call `getRGBWithTint()` first for fills — theme+tint fills return null from `getRGB()`/`getARGB()`.
  - Mark cells (light gray `#D9D9D9`, direct RGB, ±4 tolerance). Writes only ever land in these.
- **Student rows** live below the criterion header row. Filter requires column A to be **numeric** (student number) — otherwise the "coefficient du critère" label row above the students gets misidentified as a student.
- **Criterion metadata** lives on sheet `criteres_reviewed`: columns B (group), C (coefficient), D (criterion id — join key), E (contract text), F (remarks). Group name is written once per group in Excel — the parser must **fill-down** the last non-empty value across subsequent rows, otherwise criteria 17..N in the same group show blank group labels.
- **Formula cells storing criterion IDs** cache their numeric result. Reading via `getNumericCellValue()` and `String.valueOf` gives `"1.0"`, which breaks the join against the metadata sheet's numeric `"1"`. The parser respects `getCachedFormulaResultType()` and normalises integer results to `String.valueOf((long) d)`.
- **Student average** comes from column CL (index 89) on the evaluation sheet — an Excel formula. `FormulaEvaluator` recomputes it on demand, and `notifyUpdateCell()` is called on every mark write so dependent formula cache entries are invalidated.
- **Groups** are contiguous runs of criteria sharing the same (fill-down) group name. `XlsxParser.buildGroups` collapses the ordered criteria list into `Group` objects; each group's grade lives in the **"sur 6" formula column immediately to the right of its last criterion column** (`averageColumnIndex = maxCriterionColumn + 1`) — verified across every group in the reference workbook (généraux→H, 2.1→L, 2.3→R, …). That column is already on the /6 scale; read it via the same `readNumericAt`/`notifyUpdateCell` path as the CL average, don't remap.
- **Paging = one overview page per student + one page per criterion.** `MainActivity.pageIdx` runs `0..criteria.size()`; page 0 is the read-only **overview** (general CL average + dots for *all* criteria + per-group table, slider/criterion detail hidden), pages `1..N` are criterion pages (criterion index = `pageIdx - 1`). A criterion page scopes the dot row and the top-right average to the criterion's **group** (group's dots only; group's grade). `navigate()` (prev/next criterion) flattens over `students × (criteria + 1)`; `navigateStudent()` keeps the same page index across students; `navigateGroup()` jumps to a group's first criterion page. In-memory `pageIdx`/`studentIdx` survive rotation via the same `configChanges` mechanism (not persisted to `savedInstanceState`).
- **GENERAL/BACK button** (full-width, above the nav rows): from a criterion page it stashes `pageIdx` into `overviewReturnPage` and jumps to the overview so the teacher can see the marking's effect on the general grade; there it reads **BACK** and returns to the stash. `navigate()` and `navigateGroup()` flush the stash (revert to GENERAL); `navigateStudent()`/`navigateLetter()` keep it. Invariant: `overviewReturnPage >= 0` ⇒ currently on the overview page.
- **Nav bar is four two-button rows** stuck to the bottom, so labels don't hyphenate in portrait: `« lettre / lettre »`, `« étudiant / étudiant »`, `‹ groupe / groupe ›`, `‹ critère / critère ›`. Convention: double angle `«»` = across-students moves, single `‹›` = within-student. `navigateLetter()` jumps to the first student of the previous/next *starting letter that occurs in the roster* (`firstLetter()`, alphabetical, present letters only). Student-dimension moves (`navigateStudent`, `navigateLetter`) **wrap around** the roster; group/criterion moves clamp.
- **Averages are shown mark-only** (no "/6" suffix) — both the top-right average and the group grades in the overview table.
- **Top-right average / group grades are already on the /6 scale** — read via `readNumericAt`/`notifyUpdateCell`, don't remap.
- **Average text is red (`AVG_LOW_COLOR` `#D32F2F`) when the /6 value is below `AVG_LOW_THRESHOLD` (4.0)**, else the normal blue (`AVG_OK_COLOR` `#1976D2`) — applied via `avgColor()` to both the top-right average and every overview group grade.
- **Config-change survival:** MainActivity declares `android:configChanges="orientation|screenSize|screenLayout|smallestScreenSize|keyboardHidden|uiMode|locale|layoutDirection|fontScale|density"` because in-memory workbook state cannot survive an Activity recreate (rotation would silently lose unsaved marks). System-initiated process death still loses unsaved marks — accepted trade-off (see "Manual save only" below).
- **Save is manual only.** No autosave on nav or `onPause`. A visible Save button + "• unsaved" indicator in the top row make dirty state first-class. The user chose this explicitly — don't add "safety-net" autosaves.

## Apache POI on Android — hard constraints

- **`minSdk` ≥ 26.** Below that, D8 refuses to dex `poi/poifs/nio/CleanerUtil` and `log4j-api-2.21.1` because they call `MethodHandle.invokeExact` (needs Android O runtime). Don't lower it — if you need older-phone support, migrate off POI.
- **Core library desugaring** must be enabled (`coreLibraryDesugaringEnabled true` + `com.android.tools:desugar_jdk_libs`).
- **`multiDexEnabled true`** — POI blows past 64K methods.
- **META-INF packaging tuning** in `packagingOptions`: exclude `DEPENDENCIES`, `LICENSE*`, `NOTICE*`, `versions/**`, `maven/**`, `*.SF/*.DSA/*.RSA`; pickFirst `services/*` (POI uses SPI — don't exclude it).

## Build, signing, versioning

- **Debug keystore is checked in** at `app/keystore/debug.p12` (PKCS12, alias `androiddebugkey`, passwords `android`). Every CI build signs with this same key. **Do not** regenerate it — every already-installed build depends on that signature for OTA upgrades. If it must change, users will need to uninstall + reinstall once.
- **`versionCode`** = `GITHUB_RUN_NUMBER + 100` on CI, `100` on local builds. Each CI build gets a unique code so Android treats it as an upgrade.
- **`versionName`** = `0.3.0-{short SHA}`, stamped from `GITHUB_SHA`.
- **`BuildConfig.BUILD_SHA`** is the full 40-char `GITHUB_SHA`. The OTA update flow compares it against the SHA baked into the rolling release body.
- CI workflow at `.github/workflows/android_build.yml` builds a debug APK on every push, and on `main` publishes it to a rolling release tagged `latest` via `softprops/action-gh-release@v2`. Requires `permissions: contents: write`.

## Diagnosing CI failures without GitHub auth

Anonymous requests to the workflow-logs and artifact-download API endpoints get `401`/`403`. Workaround already in the workflow: on failure it posts the last 400 lines of the build log as a **commit comment** on the failing SHA. Read via:

```
GET https://api.github.com/repos/AmsteinGraphics/apk-excelparse/commits/{sha}/comments
```

Anonymous, unauthenticated, works fine.

## Downloading the CI-produced APK

`Actions -> Artifacts` requires auth. The **rolling release** does not. Predictable URL (no API rate limit):

```
https://github.com/AmsteinGraphics/apk-excelparse/releases/download/latest/app-debug.apk
```

The user has asked for the APK to land in `\\wsl$\Ubuntu\home\downloads\` (WSL path `/home/downloads/`) after every green CI. Note that on the user's Windows machine, the Bash tool is Git Bash whose `/home/` is `C:\Program Files\Git\home\` (NOT the WSL filesystem). Use PowerShell writing to the UNC path:

```powershell
Invoke-WebRequest -UseBasicParsing `
  -Uri "https://github.com/AmsteinGraphics/apk-excelparse/releases/download/latest/app-debug.apk" `
  -OutFile "\\wsl.localhost\Ubuntu\home\downloads\app-debug.apk" `
  -Headers @{'User-Agent'='apk-excelparse-tools'}
```

The WSL `/home/downloads` was chowned once to `emy:emy` so UNC writes have permission.

## In-app OTA

Menu (⋮ button top-right of grading screen, or "Check for updates" button on the picker screen) → hits `/releases/latest`, parses the 40-char SHA from the release body, compares against `BuildConfig.BUILD_SHA`. If different, offers a dialog that downloads the APK to `getCacheDir()/update.apk` and hands it to Android's package installer via `FileProvider` + `Intent.ACTION_VIEW`. First-time only, Android sends the user to Settings to grant `REQUEST_INSTALL_PACKAGES` for this app.

The FileProvider authority is `${applicationId}.fileprovider` with paths declared in `res/xml/file_paths.xml`.

## UI conventions

- Theme is `Theme.MaterialComponents.Light.NoActionBar`. There is no ActionBar, so `onCreateOptionsMenu` items are invisible. Use the ⋮ `MaterialButton` + `PopupMenu` pattern instead (already wired in `MainActivity.showOverflowMenu`).
- All user-facing terms use lowercase French: "étudiant N / M", "vue générale", nav buttons "« étudiant  ‹ critère  critère ›  étudiant »". Keep this style; do not capitalize.
- **Mark input is a row of buttons** in a `markButtonsContainer` just above the GENERAL button: a **`×` clear button** (`btnMarkClear`) then five value buttons (`btnMark0..4`, values 0…1, labels `0 ¼ ½ ¾ 1` left→right). Base style is OutlinedButton; `refreshMarkButtons(selectedIndex)` fills the selected value button with theme `colorPrimary` + white text and leaves the rest outlined (transparent fill, primary text). `-1` = no stored mark = nothing selected (the `×` is never "selected"). Tapping a value writes via `applyMarkToWorkbook`; tapping `×` calls `onClearMark()` → `XlsxParser.eraseMark` (`cell.setBlank()`) so the criterion reads ungraded again. Resolve `colorPrimary` via `androidx.appcompat.R.attr.colorPrimary` — it is NOT in `com.google.android.material.R.attr`. (Replaced the old `Slider`/`markLabel`.)
- Mark buttons + nav buttons sit at the bottom, criterion description scrolls above. Layout is intentionally structured as `LinearLayout` with the middle `ScrollView` weighted so the mark controls never shift.
- Progress row: `DotProgressView` under the student name, one dot per criterion. Unmarked = white circle with thin black stroke; marked = solid color (0.0 black, 0.25 red, 0.5 orange, 0.75 yellow, 1.0 green). Its function is progress visibility for the teacher — how much is still left to evaluate for this student.
- **Dots scale by coefficient** (`coefScale(coef, step)`): coef 2 = 100%, ±`step` per integer unit (`1 + (coef−2)·step`), clamped `[0.4, 2.0]`; radius still hard-capped at `h/2 − overshoot`. Bigger dot = heavier criterion. **The overview table uses a 0.25 step; criterion pages use a wider 0.33 step** so the coefficient weight reads more strongly on the marking screen. **The coef-2 base dot radius is a fixed `BASE_RADIUS_DP` (11.2dp), NOT derived from row height** (was `h·0.34`), so a taller row only adds headroom for heavy dots without enlarging the coef-2 dot. Still capped by `slot·0.42` so crowded rows shrink. **The header dot row is 60dp tall on criterion pages** (headroom for heavy-coef dots + the 16dp `NUMBER_BAND_DP` under them) **and 33dp on the overview** (tiny all-criteria dots, no number line — a tall row there is just whitespace); `refreshDots` sets the height per page. The overview per-group table rows stay 33dp.
- **On criterion pages every dot shows two digits** (`DotProgressView`, all at fixed `DIGIT_TEXT_DP`): the **coefficient inside the dot** (`setLabels`, `coefDigit()`) — white on graded, 50% gray (`#808080`) on blank (`drawLabel()`) — and the **criterion id under the dot** (`setBelowLabels`, `c.id`), gray, on a **fixed baseline** in the reserved bottom band. Dot centres share one line, so a smaller dot's edge sits higher and its id reads further away — the gap itself encodes coefficient weight. Both are set only on the criterion-page header row (`null` on the overview row and the table). (The old "critère N · coefficient X" detail line was removed — that info now lives in/under the dots.)
- **Highlight ring = dot radius + a fixed `HIGHLIGHT_OVERSHOOT_DP` (3dp)**, not a multiple of the radius — so it reads as a constant-weight stroke behind dots of any coefficient scale (a scaled multiple vanished behind large dots). The dot-radius cap reserves that overshoot so the ring never spills out of the row.
- **`DotProgressView.setFixedSlotPx`**: forces a fixed per-dot slot width so several instances render dots at identical physical size regardless of their own width. Its `onMeasure` honours an EXACTLY width spec (fixed table cell → dots left-pack in a uniform column) but otherwise sizes to the dot string. The overview table uses this with `headerSlotPx()` = `(screenWidth − 48dp) / criteria.size()` and a 33dp row height so each group's dots exactly match the full-width header string.
- **Overview page's per-group table** (`overviewTable`, built in `buildOverviewTable`): one borderless `TableRow` per group — group dots | group name | group grade (mark only, 11sp). The **group name is a link** (blue + underlined + ripple) that calls `jumpToGroup()` → that group's first criterion page; a **small red up-pointing triangle** (`redWarningTriangle()`, a `ShapeDrawable`/`PathShape` set as the name's `drawableStart`) precedes the link **only when the group still has an unmarked criterion**, flagging where evaluation is incomplete. Dots column is a **fixed width** (`maxDots × slot`) so names left-align; the name column is `setColumnShrinkable(1, true)` (note: `setShrinkColumns` is XML-only, no Java setter) so a long name wraps instead of pushing the average off-screen in portrait. Rebuilt on every overview render so live marks stay current.

## Working with the user

The user is a teacher, not a software developer. Explain build/CI/tooling concepts in plain language; don't assume familiarity with Gradle/Android jargon. They understand the app well from the domain side.
