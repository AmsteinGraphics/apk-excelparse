/*
 * LAPIS Mark — web version.
 *
 * Steps 1–2 of the port: the marking loop + the overview page (general/group averages and the
 * dot progress rows). Mirrors XlsxParser.java + MainActivity.
 *
 * IMPORTANT — averages are computed in JS, not read from the workbook. ExcelJS has no formula
 * engine, so the workbook's cached CL / "sur 6" values are stale the moment we write a mark. We
 * therefore replicate the formulas (validated against the reference workbook):
 *   - group grade /6 = round2( Σ(mark·critCoef)/Σ(critCoef) · 6 )   (blank marks = 0)
 *   - general CK     = round2( Σ(groupGrade·groupCoef)/Σ(groupCoef) )
 *   - general CL     = CK rounded to the nearest 0.25   (what we display)
 * The saved .xlsx still carries Excel's real formulas, so Excel recomputes them on open.
 */

const EVALUATION_SHEET = 'evaluation';
const CRITERIA_SHEET = 'criteres_reviewed';
const MARK_VALUES = [0, 0.25, 0.5, 0.75, 1];
const BUCKET_COLORS = ['#000000', '#EF5350', '#FB8C00', '#FDD835', '#66BB6A'];
const AVG_LOW = '#D32F2F', AVG_OK = '#1976D2', AVG_INCOMPLETE = '#CCCCCC', AVG_THRESHOLD = 4.0;

const el = (id) => document.getElementById(id);
const pickerScreen = el('picker'), gradingScreen = el('grading'), fileInput = el('file');
const statusEl = el('status'), studentCounter = el('studentCounter'), dirtyEl = el('dirty');
const studentName = el('studentName'), averageEl = el('average'), progressEl = el('progress');
const dotRow = el('dotRow'), critBlock = el('critBlock'), overviewBlock = el('overviewBlock');
const groupName = el('groupName'), groupCoef = el('groupCoef');
const contractEl = el('contract'), remarksEl = el('remarks');
const overviewTable = el('overviewTable'), marksEl = el('marks');
const overviewBtn = el('overviewBtn'), critNav = el('critNav');

let workbook = null, ws = null, model = null;
let studentIdx = 0;
let pageIdx = 0;              // 0 = overview, 1..N = criterion (critIdx = pageIdx - 1)
let overviewReturnCrit = 0;  // criterion to return to from the overview
let dirty = false, sourceName = 'workbook.xlsx';
let dotJump = null;          // when set on a criterion page: (slotIndex) => jump

// ---- Wiring ----
fileInput.addEventListener('change', onFilePicked);
el('saveBtn').addEventListener('click', onSave);
el('prevCrit').addEventListener('click', () => moveCrit(-1));
el('nextCrit').addEventListener('click', () => moveCrit(1));
el('prevStudent').addEventListener('click', () => moveStudent(-1));
el('nextStudent').addEventListener('click', () => moveStudent(1));
overviewBtn.addEventListener('click', toggleOverview);
marksEl.addEventListener('click', (e) => {
    const btn = e.target.closest('.mark');
    if (!btn) return;
    const idx = parseInt(btn.getAttribute('data-idx'), 10);
    if (idx === -1) clearMark(); else applyMark(idx);
});
dotRow.addEventListener('click', (e) => {
    if (!dotJump) return;
    const rect = dotRow.getBoundingClientRect();
    dotJump(e.clientX - rect.left, rect.width);
});
window.addEventListener('resize', () => { if (model) render(); });

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
        if (!model.students.length || !model.criteria.length) throw new Error('Aucun étudiant ou critère détecté.');
        studentIdx = 0; pageIdx = 0; overviewReturnCrit = 0; dirty = false;
        pickerScreen.hidden = true; gradingScreen.hidden = false;
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
    const metaById = new Map();
    const metaWs = wb.getWorksheet(CRITERIA_SHEET);
    if (metaWs) {
        let lastGroup = null;
        metaWs.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;
            const id = cellText(row.getCell(4).value).trim();
            if (!id) return;
            const g = cellText(row.getCell(2).value).trim();
            if (g) lastGroup = g;
            metaById.set(id, {
                group: lastGroup, coef: cellNumber(row.getCell(3).value),
                contract: cellText(row.getCell(5).value).trim(),
                remarks: cellText(row.getCell(6).value).trim(),
            });
        });
    }
    // Header row = the row with the most cells whose text is a known criterion id.
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

    const criteria = [];
    [...headerMap.keys()].sort((a, b) => a - b).forEach((col) => {
        const id = headerMap.get(col), m = metaById.get(id) || {};
        criteria.push({ id, col, group: m.group || '', coef: m.coef,
            contract: m.contract || '', remarks: m.remarks || '' });
    });

    const students = [];
    evalWs.eachRow((row, rowNumber) => {
        if (rowNumber <= headerRow) return;
        const num = cellNumber(row.getCell(1).value);
        const name = cellText(row.getCell(2).value).trim();
        if (num != null && name) students.push({ row: rowNumber, number: num, name });
    });

    const groups = [];
    for (let i = 0; i < criteria.length;) {
        const name = criteria[i].group;
        let last = i;
        while (last + 1 < criteria.length && criteria[last + 1].group === name) last++;
        const coef = cellNumber(evalWs.getRow(2).getCell(criteria[i].col).value);
        for (let k = i; k <= last; k++) criteria[k].groupCoef = coef;
        groups.push({ name, first: i, last, coef });
        i = last + 1;
    }
    return { students, criteria, groups };
}

// ---- Marks & averages ----
function markCell(student, criterion) { return ws.getRow(student.row).getCell(criterion.col); }
function markValue(student, criterion) { return cellNumber(markCell(student, criterion).value); }
function bucketOf(value) {
    if (value == null) return -1;
    let b = Math.round(value * 4);
    return b < 0 ? 0 : b > 4 ? 4 : b;
}
const round2 = (x) => Math.round(x * 100) / 100;

function groupGrade(student, group) {
    let sw = 0, w = 0, complete = true;
    for (let k = group.first; k <= group.last; k++) {
        const c = model.criteria[k], v = markValue(student, c), coef = c.coef != null ? c.coef : 1;
        if (v == null) complete = false;
        sw += (v == null ? 0 : v) * coef; w += coef;
    }
    return { grade: w > 0 ? round2((sw / w) * 6) : null, complete };
}

function generalAverage(student) {
    let sw = 0, w = 0;
    for (const g of model.groups) {
        const { grade } = groupGrade(student, g);
        if (grade == null || g.coef == null) continue;
        sw += grade * g.coef; w += g.coef;
    }
    if (w <= 0) return null;
    const ck = round2(sw / w);
    return round2(Math.floor(ck * 4 + 0.5) / 4);
}
const avgColor = (v) => (v != null && v < AVG_THRESHOLD) ? AVG_LOW : AVG_OK;

function applyMark(idx) {
    markCell(model.students[studentIdx], model.criteria[pageIdx - 1]).value = MARK_VALUES[idx];
    setDirty(); render();
}
function clearMark() {
    markCell(model.students[studentIdx], model.criteria[pageIdx - 1]).value = null;
    setDirty(); render();
}
function markedCount(student) {
    let n = 0;
    for (const c of model.criteria) if (markValue(student, c) != null) n++;
    return n;
}

// ---- Coefficient dot scaling (mirrors DotProgressView / coefScale) ----
function coefScale(coef, step) {
    const c = coef != null ? coef : 2;
    let s = 1 + (c - 2) * step;
    return s < 0.4 ? 0.4 : s > 2 ? 2 : s;
}
function coefDigit(coef) {
    if (coef == null) return '';
    return Number.isInteger(coef) ? String(coef) : String(coef).replace(/\.?0+$/, '');
}

function drawDots(canvas, dots, opts) {
    const cssW = opts.width || canvas.clientWidth || 300;
    const cssH = opts.height;
    const dpr = window.devicePixelRatio || 1;
    canvas.style.width = cssW + 'px';
    canvas.style.height = cssH + 'px';
    canvas.width = Math.round(cssW * dpr);
    canvas.height = Math.round(cssH * dpr);
    const ctx = canvas.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, cssW, cssH);
    const n = dots.length;
    if (!n) return;
    const slot = cssW / n;
    const belowBand = opts.showBelow ? 16 : 0;
    const areaH = cssH - belowBand;
    const overshoot = 3;
    const maxR = areaH / 2 - overshoot;
    const baseR = Math.min(slot * 0.42, (opts.baseRadius || 11.2) * (opts.sizeMult || 1));
    const cy = areaH / 2;
    for (let i = 0; i < n; i++) {
        const d = dots[i], cx = slot * (i + 0.5);
        let r = Math.min(baseR * coefScale(d.coef, opts.step || 0.25), maxR);
        if (r < 1) r = 1;
        if (i === opts.highlightIndex) {
            ctx.beginPath(); ctx.arc(cx, cy, r + overshoot, 0, 2 * Math.PI);
            ctx.fillStyle = '#000'; ctx.fill();
        }
        ctx.beginPath(); ctx.arc(cx, cy, r, 0, 2 * Math.PI);
        if (d.bucket < 0) {
            ctx.fillStyle = '#fff'; ctx.fill();
            ctx.lineWidth = 1; ctx.strokeStyle = '#000'; ctx.stroke();
        } else {
            ctx.fillStyle = BUCKET_COLORS[d.bucket]; ctx.fill();
        }
        if (opts.showLabels && d.coef != null) {
            ctx.fillStyle = d.bucket < 0 ? '#808080' : '#fff';
            ctx.font = 'bold ' + Math.round(12 * (opts.sizeMult || 1)) + 'px sans-serif';
            ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
            ctx.fillText(coefDigit(d.coef), cx, cy);
        }
        if (opts.showBelow && d.id != null) {
            ctx.fillStyle = '#808080'; ctx.font = 'bold 12px sans-serif';
            ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
            ctx.fillText(String(d.id), cx, areaH + belowBand / 2);
        }
    }
}

// ---- Navigation ----
function moveCrit(delta) {
    const n = model.criteria.length;
    let ci = ((pageIdx - 1 + delta) % n + n) % n;
    pageIdx = ci + 1;
    render();
}
function moveStudent(delta) {
    const n = model.students.length;
    studentIdx = ((studentIdx + delta) % n + n) % n;
    render();
}
function toggleOverview() {
    if (pageIdx === 0) { pageIdx = overviewReturnCrit + 1; }
    else { overviewReturnCrit = pageIdx - 1; pageIdx = 0; }
    render();
}
function jumpToCriterion(ci) { pageIdx = ci + 1; render(); }
function groupOfCriterion(ci) {
    for (const g of model.groups) if (ci >= g.first && ci <= g.last) return g;
    return null;
}

// ---- Render ----
function render() {
    const s = model.students[studentIdx];
    const overview = pageIdx === 0;
    studentCounter.textContent = 'étudiant ' + (studentIdx + 1) + ' / ' + model.students.length;
    studentName.textContent = s.name;
    progressEl.textContent = markedCount(s) + ' / ' + model.criteria.length + ' critères notés';

    critBlock.hidden = overview;
    overviewBlock.hidden = !overview;
    marksEl.hidden = overview;
    critNav.hidden = overview;
    overviewBtn.textContent = overview ? 'CRITÈRES' : 'GÉNÉRAL';

    if (overview) {
        renderOverview(s);
    } else {
        renderCriterion(s);
    }
}

function renderOverview(s) {
    const avg = generalAverage(s);
    averageEl.textContent = avg == null ? '' : avg.toFixed(2);
    averageEl.style.color = avgColor(avg);
    dotJump = null;
    // All criteria dots, small, non-interactive.
    const dots = model.criteria.map((c) => ({ bucket: bucketOf(markValue(s, c)), coef: c.coef, id: c.id }));
    requestAnimationFrame(() => drawDots(dotRow, dots, { height: 26, step: 0.25 }));

    // Per-group table.
    overviewTable.innerHTML = '';
    for (const g of model.groups) {
        const { grade, complete } = groupGrade(s, g);
        const row = document.createElement('div');
        row.className = 'grp-row';

        const canvas = document.createElement('canvas');
        canvas.className = 'grp-dots';
        const gd = [];
        for (let k = g.first; k <= g.last; k++) {
            const c = model.criteria[k];
            gd.push({ bucket: bucketOf(markValue(s, c)), coef: c.coef, id: c.id });
        }
        const slot = 13;
        drawDots(canvas, gd, { width: gd.length * slot, height: 20, step: 0.25, baseRadius: 5 });
        row.appendChild(canvas);

        const name = document.createElement('span');
        name.className = 'grp-name';
        name.textContent = g.name || '';
        name.addEventListener('click', () => jumpToCriterion(g.first));
        row.appendChild(name);

        const gradeEl = document.createElement('span');
        gradeEl.className = 'grp-grade';
        gradeEl.textContent = grade == null ? '' : grade.toFixed(2);
        gradeEl.style.color = !complete ? AVG_INCOMPLETE : avgColor(grade);
        row.appendChild(gradeEl);

        overviewTable.appendChild(row);
    }
}

function renderCriterion(s) {
    const ci = pageIdx - 1;
    const c = model.criteria[ci];
    const g = groupOfCriterion(ci);

    // Top-right average = this group's /6 grade (red/blue; gray-for-incomplete is overview-only).
    const gg = g ? groupGrade(s, g) : { grade: null, complete: true };
    averageEl.textContent = gg.grade == null ? '' : gg.grade.toFixed(2);
    averageEl.style.color = avgColor(gg.grade);

    groupName.textContent = c.group || '';
    groupCoef.textContent = c.groupCoef != null ? 'coeff. ' + coefDigit(c.groupCoef) : '';
    groupCoef.hidden = c.groupCoef == null;
    contractEl.textContent = c.contract || '';
    remarksEl.textContent = c.remarks || '';

    // Group's dots — larger, with coefficient digit inside + criterion id below; current highlighted.
    const first = g ? g.first : ci, last = g ? g.last : ci;
    const dots = [];
    for (let k = first; k <= last; k++) {
        const cc = model.criteria[k];
        dots.push({ bucket: bucketOf(markValue(s, cc)), coef: cc.coef, id: cc.id });
    }
    const highlight = ci - first;
    requestAnimationFrame(() => drawDots(dotRow, dots, {
        height: 56, step: 0.33, showLabels: true, showBelow: true, highlightIndex: highlight,
    }));
    // Tapping a dot jumps to that criterion within the group.
    dotJump = (x, width) => {
        const slot = width / dots.length;
        let idx = Math.floor(x / slot);
        if (idx < 0) idx = 0; if (idx >= dots.length) idx = dots.length - 1;
        jumpToCriterion(first + idx);
    };

    const bucket = bucketOf(markValue(s, c));
    for (const btn of marksEl.querySelectorAll('.mark')) {
        const idx = parseInt(btn.getAttribute('data-idx'), 10);
        btn.classList.toggle('selected', idx === bucket && idx >= 0);
    }
    dirtyEl.hidden = !dirty;
}

function setDirty() { dirty = true; dirtyEl.hidden = false; }

// ---- Save ----
/*
 * ExcelJS can't write a workbook that uses *shared* formulas (one master cell + clones) — it
 * throws "Shared Formula master must exist above and or left of clone". The reference workbook
 * stores the per-student "sur 6"/CL grades that way. Rewriting every formula cell as a standalone
 * (translated) formula removes the shared relationship while keeping the formula, so Excel still
 * recalculates on open. Idempotent; safe to run before every save.
 */
function deshareFormulas(wb) {
    wb.eachSheet((sheet) => {
        sheet.eachRow({ includeEmpty: false }, (row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                try {
                    if (cell.type !== ExcelJS.ValueType.Formula) return;
                    const cur = cell.value;
                    const result = (cur && typeof cur === 'object' && 'result' in cur) ? cur.result : undefined;
                    const formula = cell.formula; // ExcelJS translates shared clones to their own formula
                    if (formula) cell.value = { formula, result };
                } catch (e) { /* leave this cell as-is */ }
            });
        });
    });
}

async function onSave() {
    if (!workbook) return;
    const btn = el('saveBtn');
    btn.disabled = true;
    dirtyEl.textContent = '• sauvegarde…'; dirtyEl.hidden = false;
    try {
        deshareFormulas(workbook);
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = gradedFilename(sourceName);
        document.body.appendChild(a); a.click(); a.remove();
        URL.revokeObjectURL(a.href);
        dirty = false;
    } catch (err) {
        console.error(err);
        alert('Échec de la sauvegarde : ' + (err && err.message ? err.message : err));
    } finally {
        btn.disabled = false;
        dirtyEl.textContent = '• modifié';
        dirtyEl.hidden = !dirty;
    }
}
function gradedFilename(name) {
    const dot = name.lastIndexOf('.');
    return (dot > 0 ? name.slice(0, dot) : name) + '-graded.xlsx';
}
function setStatus(text, isError) {
    statusEl.textContent = text;
    statusEl.classList.toggle('error', !!isError);
}
