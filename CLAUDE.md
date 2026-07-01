# apk-excelparse — context for future Claude Code sessions

This file exists so a Claude Code session on any machine can pick up the project without re-deriving decisions or repeating past mistakes. Update it when you make a non-obvious change; delete stale entries.

## What the app does

Personal Android grading app for a Swiss teacher. Reads student rows and per-criterion columns from an `.xlsx` file, presents one criterion at a time on a phone-sized screen with a slider that writes marks (0.0 / 0.25 / 0.5 / 0.75 / 1.0) back into the workbook. The original picked file is never modified — a working copy lives in app-private storage, and each Save mirrors it to `Downloads/{basename}-graded.xlsx` via MediaStore.

## Non-obvious architecture

- **Cells are identified by fill color, not by fixed ranges.** Merged cells and group-average columns in the workbook mean row/column indices aren't stable. The parser scans for two colors on the `evaluation` sheet:
  - Criterion label cells (POI-resolved RGB `#E6E9EB`, ±4 tolerance). Excel *displays* these as pale blue `#DAE9F8`; POI's HSL-based tint math on the Accent-1 theme color resolves to that near-neutral. Trust POI's number, not Excel's swatch. Always call `getRGBWithTint()` first for fills — theme+tint fills return null from `getRGB()`/`getARGB()`.
  - Mark cells (light gray `#D9D9D9`, direct RGB, ±4 tolerance). Writes only ever land in these.
- **Student rows** live below the criterion header row. Filter requires column A to be **numeric** (student number) — otherwise the "coefficient du critère" label row above the students gets misidentified as a student.
- **Criterion metadata** lives on sheet `criteres_reviewed`: columns B (group), C (coefficient), D (criterion id — join key), E (contract text), F (remarks). Group name is written once per group in Excel — the parser must **fill-down** the last non-empty value across subsequent rows, otherwise criteria 17..N in the same group show blank group labels.
- **Formula cells storing criterion IDs** cache their numeric result. Reading via `getNumericCellValue()` and `String.valueOf` gives `"1.0"`, which breaks the join against the metadata sheet's numeric `"1"`. The parser respects `getCachedFormulaResultType()` and normalises integer results to `String.valueOf((long) d)`.
- **Student average** comes from column CL (index 89) on the evaluation sheet — an Excel formula. `FormulaEvaluator` recomputes it on demand, and `notifyUpdateCell()` is called on every mark write so dependent formula cache entries are invalidated.
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
- All user-facing terms use lowercase French: "étudiant N / M", "critère 12 · coefficient 2.0", nav buttons "« étudiant  ‹ critère  critère ›  étudiant »". Keep this style; do not capitalize.
- Slider stays at the bottom, criterion description scrolls above. Layout is intentionally structured as `LinearLayout` with the middle `ScrollView` weighted so the mark controls never shift.
- Progress row: `DotProgressView` under the student name, one dot per criterion. Unmarked = white circle with thin black stroke; marked = solid color (0.0 black, 0.25 red, 0.5 orange, 0.75 yellow, 1.0 green). Its function is progress visibility for the teacher — how much is still left to evaluate for this student.

## Working with the user

The user is a teacher, not a software developer. Explain build/CI/tooling concepts in plain language; don't assume familiarity with Gradle/Android jargon. They understand the app well from the domain side.
