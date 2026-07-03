/*
 * LAPIS Mark — web version.
 *
 * Step 1 of the port: the marking loop. Opens an .xlsx, recovers the same model the Android
 * app uses (students, criteria with group/coefficient/contract, mark cells), lets the teacher
 * mark one criterion at a time, and saves a graded copy as a download. Mirrors XlsxParser.java
 * and MainActivity's marking flow. The original file is never touched (the browser can't anyway);
 * Save downloads a new "<name>-graded.xlsx".
 *
 * Parsing notes (validated against the reference workbook):
 *   - Sheet "criteres_reviewed": rows 2+, col B group (written once per group → fill down),
 *     C coefficient, D criterion id (join key), E contract, F remarks.
 *   - Sheet "evaluation": a header row (the "numéro du critère" row) carries each criterion id
 *     per column (as a formula result). We map column → id by matching those against the known
 *     ids — robust and independent of theme-fill detection. Mark cells sit at
 *     (student row, criterion column); they are a theme fill (theme 0, tint -0.15 → #D9D9D9).
 *   - Students: rows below the header with a numeric column A and a non-empty column B (name).
 *   - Group coefficient: row 2 ("coefficient du groupe") at each group's first criterion column.
 */

const EVALUATION_SHEET = 'evaluation';
const CRITERIA_SHEET = 'criteres_reviewed';
const MARK_VALUES = [0, 0.25, 0.5, 0.75, 1];

// ---- DOM ----
const el = (id) => document.getElementById(id);
const pickerScreen = el('picker');
const gradingScreen = el('grading');
const fileInput = el('file');
const statusEl = el('status');
const studentCounter = el('studentCounter');
const dirtyEl = el('dirty');
const studentName = el('studentName');
const progressEl = el('progress');
const groupName = el('groupName');
const groupCoef = el('groupCoef');
const contractEl = el('contract');
const remarksEl = el('remarks');
const marksEl = el('marks');

// ---- State ----
let workbook = null;      // ExcelJS.Workbook
let ws = null;            // evaluation worksheet
let model = null;         // { students, criteria, groups }
let studentIdx = 0;
let critIdx = 0;
let dirty = false;
let sourceName = 'workbook.xlsx';

// ---- Wiring ----
fileInput.addEventListener('change', onFilePicked);
el('saveBtn').addEventListener('click', onSave);
el('prevCrit').addEventListener('click', () => moveCrit(-1));
el('nextCrit').addEventListener('click', () => moveCrit(1));
el('prevStudent').addEventListener('click', () => moveStudent(-1));
el('nextStudent').addEventListener('click', () => moveStudent(1));
marksEl.addEventListener('click', (e) => {
    const btn = e.target.closest('.mark');
    if (!btn) return;
    const idx = parseInt(btn.getAttribute('data-idx'), 10);
    if (idx === -1) clearMark();
    else applyMark(idx);
});

// ---- Load ----
async function onFilePicked(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    sourceName = file.name;
    setStatus('Lecture de « ' + file.name + ' »…');
    try {
        const buffer = await file.arrayBuffer();
        workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        ws = workbook.getWorksheet(EVALUATION_SHEET);
        if (!ws) throw new Error("Feuille « " + EVALUATION_SHEET + " » introuvable.");
        model = parseModel(workbook, ws);
        if (!model.students.length || !model.criteria.length) {
            throw new Error('Aucun étudiant ou critère détecté.');
        }
        studentIdx = 0;
        critIdx = 0;
        dirty = false;
        pickerScreen.hidden = true;
        gradingScreen.hidden = false;
        render();
    } catch (err) {
        console.error(err);
        setStatus('Impossible de lire le fichier : ' + (err && err.message ? err.message : err), true);
    }
}

// ---- Parsing ----
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

function cellNumber(v) {
    if (typeof v === 'number') return v;
    if (v && typeof v === 'object' && typeof v.result === 'number') return v.result;
    if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) return Number(v);
    return null;
}

function parseModel(wb, evalWs) {
    // 1) Criterion metadata from criteres_reviewed (group filled down).
    const metaById = new Map();
    const metaWs = wb.getWorksheet(CRITERIA_SHEET);
    if (metaWs) {
        let lastGroup = null;
        metaWs.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return; // header
            const id = cellText(row.getCell(4).value).trim();
            if (!id) return;
            const g = cellText(row.getCell(2).value).trim();
            if (g) lastGroup = g;
            metaById.set(id, {
                group: lastGroup,
                coef: cellNumber(row.getCell(3).value),
                contract: cellText(row.getCell(5).value).trim(),
                remarks: cellText(row.getCell(6).value).trim(),
            });
        });
    }

    // 2) Header row on evaluation = the row with the most cells whose text is a known id.
    let headerRow = -1, headerMap = null, headerHits = 0;
    evalWs.eachRow((row, rowNumber) => {
        const map = new Map();
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            const t = cellText(cell.value).trim();
            if (metaById.has(t)) map.set(colNumber, t);
        });
        if (map.size > headerHits) { headerHits = map.size; headerMap = map; headerRow = rowNumber; }
    });
    if (!headerMap || headerMap.size === 0) throw new Error('Ligne des critères introuvable.');

    // 3) Criteria ordered by column.
    const criteria = [];
    [...headerMap.keys()].sort((a, b) => a - b).forEach((col) => {
        const id = headerMap.get(col);
        const m = metaById.get(id) || {};
        criteria.push({
            id, col,
            group: m.group || '',
            coef: m.coef,
            contract: m.contract || '',
            remarks: m.remarks || '',
        });
    });

    // 4) Students: rows after the header with numeric col A and a name in col B.
    const students = [];
    evalWs.eachRow((row, rowNumber) => {
        if (rowNumber <= headerRow) return;
        const num = cellNumber(row.getCell(1).value);
        const name = cellText(row.getCell(2).value).trim();
        if (num != null && name) students.push({ row: rowNumber, number: num, name });
    });

    // 5) Groups: contiguous runs sharing a group name; coef from row 2 at the run's first column.
    const groups = [];
    for (let i = 0; i < criteria.length;) {
        const name = criteria[i].group;
        let last = i;
        while (last + 1 < criteria.length && criteria[last + 1].group === name) last++;
        const firstCol = criteria[i].col;
        const coef = cellNumber(evalWs.getRow(2).getCell(firstCol).value);
        for (let k = i; k <= last; k++) criteria[k].groupCoef = coef;
        groups.push({ name, first: i, last, coef });
        i = last + 1;
    }

    return { students, criteria, groups };
}

// ---- Marks ----
function markCell(student, criterion) {
    return ws.getRow(student.row).getCell(criterion.col);
}

function readBucket(student, criterion) {
    const v = cellNumber(markCell(student, criterion).value);
    if (v == null) return -1;
    let b = Math.round(v * 4);
    return b < 0 ? 0 : b > 4 ? 4 : b;
}

function applyMark(idx) {
    const s = model.students[studentIdx];
    const c = model.criteria[critIdx];
    markCell(s, c).value = MARK_VALUES[idx];
    setDirty();
    render();
}

function clearMark() {
    const s = model.students[studentIdx];
    const c = model.criteria[critIdx];
    markCell(s, c).value = null;
    setDirty();
    render();
}

function markedCount(student) {
    let n = 0;
    for (const c of model.criteria) if (readBucket(student, c) >= 0) n++;
    return n;
}

// ---- Navigation ----
function moveCrit(delta) {
    const n = model.criteria.length;
    critIdx = ((critIdx + delta) % n + n) % n;
    render();
}
function moveStudent(delta) {
    const n = model.students.length;
    studentIdx = ((studentIdx + delta) % n + n) % n;
    render();
}

// ---- Render ----
function render() {
    const s = model.students[studentIdx];
    const c = model.criteria[critIdx];
    studentCounter.textContent = 'étudiant ' + (studentIdx + 1) + ' / ' + model.students.length;
    studentName.textContent = s.name;
    progressEl.textContent = markedCount(s) + ' / ' + model.criteria.length + ' critères notés';
    groupName.textContent = c.group || '';
    groupCoef.textContent = c.groupCoef != null ? 'coeff. ' + c.groupCoef : '';
    groupCoef.style.display = c.groupCoef != null ? '' : 'none';
    contractEl.textContent = c.contract || '';
    remarksEl.textContent = c.remarks || '';

    const bucket = readBucket(s, c);
    for (const btn of marksEl.querySelectorAll('.mark')) {
        const idx = parseInt(btn.getAttribute('data-idx'), 10);
        btn.classList.toggle('selected', idx === bucket && idx >= 0);
    }
    dirtyEl.hidden = !dirty;
}

function setDirty() {
    dirty = true;
    dirtyEl.hidden = false;
}

// ---- Save (download graded copy) ----
async function onSave() {
    if (!workbook) return;
    const btn = el('saveBtn');
    btn.disabled = true;
    dirtyEl.textContent = '• sauvegarde…';
    dirtyEl.hidden = false;
    try {
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = gradedFilename(sourceName);
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(a.href);
        dirty = false;
    } catch (err) {
        console.error(err);
        alert('Échec de la sauvegarde : ' + (err && err.message ? err.message : err));
    } finally {
        btn.disabled = false;
        dirtyEl.textContent = '• modifié';
        dirtyEl.hidden = !dirty;
    }
}

function gradedFilename(name) {
    const dot = name.lastIndexOf('.');
    const base = dot > 0 ? name.slice(0, dot) : name;
    return base + '-graded.xlsx';
}

function setStatus(text, isError) {
    statusEl.textContent = text;
    statusEl.classList.toggle('error', !!isError);
}
