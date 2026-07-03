/*
 * LAPIS Mark — web skeleton.
 *
 * Goal of this skeleton: prove the risky part of a browser port works end to end —
 * open an .xlsx in the browser AND read cell fill colours, then recover the workbook
 * structure the Android app relies on. It mirrors (a simplified slice of) XlsxParser.java:
 *
 *   - Sheet "evaluation" holds students (rows) and criteria (columns).
 *   - MARK cells are a direct RGB light gray (#D9D9D9). Those columns are the criterion
 *     columns; every marked/markable student sits on a row that has them.
 *   - Student rows = column A is numeric (student number); column B is the name.
 *
 * IMPORTANT — the mark cells are NOT a plain RGB. In the reference workbook they are a
 * *theme* fill (theme 0 = the light background, tint -0.15), which Excel/POI resolve to
 * #D9D9D9. ExcelJS reports such cells as {theme, tint}, not {argb}. So we detect a mark cell
 * by EITHER a resolved gray argb OR the theme+tint signature. (This is the browser equivalent
 * of POI's getRGBWithTint, and it's the crux of the whole port: nearly every cell the app
 * cares about is theme-coloured.) Criterion *label* cells use an Accent-1 theme+tint fill —
 * detecting those (to read criterion ids/order) is the next step, same mechanism.
 *
 * Nothing here writes to the file.
 */

const EVALUATION_SHEET = 'evaluation';
const MARK_RGB = { r: 0xD9, g: 0xD9, b: 0xD9 };
const RGB_TOLERANCE = 6;
// Mark cells as a theme fill: theme 0 (light background) darkened by this tint → #D9D9D9.
const MARK_THEME = 0;
const MARK_TINT = -0.15;
const TINT_TOLERANCE = 0.03;

const fileInput = document.getElementById('file');
const statusEl = document.getElementById('status');
const summaryEl = document.getElementById('summary');
const studentsSection = document.getElementById('studentsSection');
const studentListEl = document.getElementById('studentList');
const statStudents = document.getElementById('statStudents');
const statCriteria = document.getElementById('statCriteria');

fileInput.addEventListener('change', onFilePicked);

async function onFilePicked(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setStatus('Lecture de « ' + file.name + ' »…');
    hideResults();
    try {
        const buffer = await file.arrayBuffer();
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        const model = parseWorkbook(wb);
        render(model);
        setStatus('« ' + file.name +' » chargé.');
    } catch (err) {
        console.error(err);
        setStatus('Impossible de lire le fichier : ' + (err && err.message ? err.message : err), true);
    }
}

/** The solid-fill foreground colour object ExcelJS gives, or null. May be {argb} or {theme,tint}. */
function fillFg(cell) {
    const f = cell.fill;
    if (!f || f.type !== 'pattern' || !f.fgColor) return null;
    return f.fgColor;
}

/** True if the fill is the gray mark-cell fill — matched as a resolved RGB OR the theme+tint form. */
function isMarkFill(cell) {
    const fg = fillFg(cell);
    if (!fg) return false;
    // Path A: a resolved/plain gray RGB (#D9D9D9 ± tolerance).
    if (typeof fg.argb === 'string' && fg.argb.length >= 8) {
        const r = parseInt(fg.argb.substr(2, 2), 16);
        const g = parseInt(fg.argb.substr(4, 2), 16);
        const b = parseInt(fg.argb.substr(6, 2), 16);
        if (Math.abs(r - MARK_RGB.r) <= RGB_TOLERANCE
            && Math.abs(g - MARK_RGB.g) <= RGB_TOLERANCE
            && Math.abs(b - MARK_RGB.b) <= RGB_TOLERANCE) return true;
    }
    // Path B: the theme+tint signature (theme 0 darkened by ~-0.15).
    if (fg.theme === MARK_THEME && Math.abs((fg.tint || 0) - MARK_TINT) <= TINT_TOLERANCE) {
        return true;
    }
    return false;
}

/** ExcelJS cell values can be strings, numbers, rich text, or formula objects. */
function cellText(v) {
    if (v == null) return '';
    if (typeof v === 'object') {
        if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join('');
        if ('result' in v) return v.result == null ? '' : String(v.result);
        if ('text' in v) return String(v.text);
        return '';
    }
    return String(v);
}

function isNumericValue(v) {
    if (typeof v === 'number') return true;
    if (typeof v === 'object' && v && 'result' in v) return typeof v.result === 'number';
    if (typeof v === 'string' && v.trim() !== '') return !isNaN(Number(v));
    return false;
}

function parseWorkbook(wb) {
    const ws = wb.getWorksheet(EVALUATION_SHEET);
    if (!ws) {
        throw new Error("Feuille « " + EVALUATION_SHEET + " » introuvable.");
    }
    const criterionCols = new Set();
    const students = [];
    ws.eachRow({ includeEmpty: false }, (row) => {
        // Collect mark columns from this row's gray mark cells.
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            if (isMarkFill(cell)) criterionCols.add(colNumber);
        });
        // Student row: column A numeric (student number), column B a non-empty name.
        const a = row.getCell(1).value;
        const name = cellText(row.getCell(2).value).trim();
        if (isNumericValue(a) && name !== '') {
            students.push({ number: cellText(a).trim(), name });
        }
    });
    return { students, criterionColumnCount: criterionCols.size };
}

function render(model) {
    statStudents.textContent = String(model.students.length);
    statCriteria.textContent = String(model.criterionColumnCount);
    summaryEl.hidden = false;

    studentListEl.innerHTML = '';
    for (const s of model.students) {
        const li = document.createElement('li');
        li.textContent = s.name;
        studentListEl.appendChild(li);
    }
    studentsSection.hidden = model.students.length === 0;
}

function setStatus(text, isError) {
    statusEl.textContent = text;
    statusEl.classList.toggle('error', !!isError);
}

function hideResults() {
    summaryEl.hidden = true;
    studentsSection.hidden = true;
}
