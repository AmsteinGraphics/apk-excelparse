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

// Build marker: liste — full-roster page (⋮ menu) linking each student to their GENERAL page.
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
const menuBtn = el('menuBtn'), menuEl = el('menu');
const filterScreen = el('filter'), completionScreen = el('completion'), subpageScreen = el('subpage');
const filterList = el('filterList'), completionContent = el('completionContent');
const subpageTitle = el('subpageTitle'), subpageContent = el('subpageContent');
const noteBox = el('noteBox'), backToNotesBtn = el('backToNotes');
const redFlagBlock = el('redFlagBlock'), redFlagCheck = el('redFlagCheck'), redFlagReason = el('redFlagReason');
const notesScreen = el('notes'), noteCounter = el('noteCounter'), noteEmpty = el('noteEmpty');
const noteBody = el('noteBody'), noteStudent = el('noteStudent'), noteGroupEl = el('noteGroup');
const noteCrit = el('noteCrit'), noteDotEl = el('noteDot'), noteEdit = el('noteEdit');

const NOTES_SHEET = 'notes';
// Red flag: stored on the notes sheet at the student-name column (col B = ExcelJS colNumber 2),
// marker-prefixed so an on-but-empty flag differs from no flag and reads as a flag in Excel.
const FLAG_MARK = '⚑';
const FLAG_COL = 2;

let workbook = null, ws = null, model = null;
let studentIdx = 0;
let pageIdx = 0;              // 0 = overview, 1..N = criterion (critIdx = pageIdx - 1)
let overviewReturnCrit = 0;  // criterion to return to from the overview
let dirty = false, sourceName = 'workbook.xlsx';
let dotJump = null;          // when set on a criterion page: (slotIndex) => jump
let groupHidden = [];        // groupHidden[i] hides model.groups[i] everywhere; per-workbook persisted
let notesList = [];          // [{si, ci}] pairs that have a note
let notePos = 0;             // carousel cursor
let noteReturnPos = -1;      // when a criterion page is reached via a note, the note to return to
// When a subpage link opens a student's GENERAL page, this remembers which subpage to return to so
// the GÉNÉRAL button reads RETOUR and goes back to that list. null = none. 'liste' = the roster.
let overviewReturnSubpage = null;
let subpageIsListe = false;  // routes the shared subpage BACK button (roster → grading, else → completion)

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

// Menu + overlays
menuBtn.addEventListener('click', (e) => { e.stopPropagation(); menuEl.hidden = !menuEl.hidden; });
document.addEventListener('click', () => { menuEl.hidden = true; });
menuEl.addEventListener('click', (e) => {
    const b = e.target.closest('button'); if (!b) return;
    const act = b.getAttribute('data-act');
    if (act === 'filter') showFilter();
    else if (act === 'completion') showCompletion();
    else if (act === 'notes') showNotes();
    else if (act === 'liste') showListe();
    else if (act === 'change') changeFile();
});
el('filterBack').addEventListener('click', closeFilter);
el('completionBack').addEventListener('click', () => showScreen('grading'));
el('subpageBack').addEventListener('click', () => subpageIsListe ? showScreen('grading') : showCompletion());
el('selectAll').addEventListener('click', () => setAllGroupsHidden(false));
el('clearAll').addEventListener('click', () => setAllGroupsHidden(true));
// Criterion-page note box: edit in place, saved into the workbook.
noteBox.addEventListener('input', () => {
    if (pageIdx === 0) return;
    setNote(model.students[studentIdx], model.criteria[pageIdx - 1], noteBox.value);
});
// Red flag (GENERAL page). Toggling re-renders (updates the name dot + reason box); editing the
// reason writes on each keystroke without re-render, so the textarea keeps focus (like the note box).
redFlagCheck.addEventListener('change', () => {
    const s = model.students[studentIdx];
    setFlag(s, redFlagCheck.checked, flagReason(s));
    render();
});
redFlagReason.addEventListener('input', () => {
    if (pageIdx !== 0) return;
    setFlag(model.students[studentIdx], true, redFlagReason.value);
});
// Notes carousel.
el('prevNote').addEventListener('click', () => navigateNote(-1));
el('nextNote').addEventListener('click', () => navigateNote(1));
el('notesBack').addEventListener('click', () => showScreen('grading'));
noteEdit.addEventListener('input', () => {
    if (!notesList.length) return;
    const { si, ci } = notesList[notePos];
    setNote(model.students[si], model.criteria[ci], noteEdit.value);
});
noteDotEl.addEventListener('click', gotoNoteCriterion);
backToNotesBtn.addEventListener('click', () => {
    if (noteReturnPos >= 0) notePos = noteReturnPos;
    noteReturnPos = -1;
    showNotes();
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
        if (!model.students.length || !model.criteria.length) throw new Error('Aucun étudiant ou critère détecté.');
        studentIdx = 0; pageIdx = 0; overviewReturnCrit = 0; overviewReturnSubpage = null; dirty = false;
        loadFilter();
        showScreen('grading');
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
    // Header row = the row matching the most DISTINCT criterion ids (one per column). Scoring by
    // distinct ids (not raw matches) is essential: the coefficient row and "sur 6" sums are small
    // integers that collide with the id space (1..61), so a raw count would wrongly pick row 4.
    let headerRow = -1, headerMap = null, headerHits = 0;
    evalWs.eachRow((row, rowNumber) => {
        const map = new Map();
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            const t = cellText(cell.value).trim();
            if (metaById.has(t)) map.set(colNumber, t);
        });
        const distinct = new Set(map.values()).size;
        if (distinct > headerHits) { headerHits = distinct; headerMap = map; headerRow = rowNumber; }
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
    for (const gi of visibleGroupIdx()) {
        const g = model.groups[gi];
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
        // Thin gray reference rings at every possible weight, behind the dot.
        if (opts.referenceScales) {
            ctx.lineWidth = 1; ctx.strokeStyle = '#CCCCCC';
            for (const rs of opts.referenceScales) {
                const rr = Math.min(baseR * rs, maxR);
                if (rr >= 1) { ctx.beginPath(); ctx.arc(cx, cy, rr, 0, 2 * Math.PI); ctx.stroke(); }
            }
        }
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
    const vis = visibleCriteria();
    if (!vis.length) return;
    let pos = vis.indexOf(pageIdx - 1);
    if (pos < 0) pos = delta >= 0 ? 0 : vis.length - 1;
    else pos = ((pos + delta) % vis.length + vis.length) % vis.length;
    pageIdx = vis[pos] + 1;
    noteReturnPos = -1;
    overviewReturnSubpage = null;
    render();
}
function moveStudent(delta) {
    const n = model.students.length;
    studentIdx = ((studentIdx + delta) % n + n) % n;
    noteReturnPos = -1;
    render();
}
function toggleOverview() {
    if (overviewReturnSubpage) {
        // Reached this overview from a subpage link: RETOUR goes back to that list.
        const kind = overviewReturnSubpage;
        overviewReturnSubpage = null;
        if (kind === 'liste') showListe(); else showSubpage(kind);
        return;
    }
    if (pageIdx === 0) {
        let ci = overviewReturnCrit;
        if (!isCriterionVisible(ci)) { const vis = visibleCriteria(); ci = vis.length ? vis[0] : 0; }
        pageIdx = ci + 1;
    } else {
        overviewReturnCrit = pageIdx - 1;
        pageIdx = 0;
    }
    render();
}
function jumpToCriterion(ci) { pageIdx = ci + 1; noteReturnPos = -1; overviewReturnSubpage = null; render(); }
function groupOfCriterion(ci) {
    for (const g of model.groups) if (ci >= g.first && ci <= g.last) return g;
    return null;
}

// ---- Render ----
function render() {
    const s = model.students[studentIdx];
    const overview = pageIdx === 0;
    studentCounter.textContent = 'étudiant ' + (studentIdx + 1) + ' / ' + model.students.length;
    setStudentLabel(studentName, studentIdx, null);
    progressEl.textContent = markedCount(s) + ' / ' + model.criteria.length + ' critères notés';

    // Red flag tick + reason box, under the name — GENERAL page only.
    redFlagBlock.hidden = !overview;
    if (overview) {
        const flagged = isFlagged(s);
        redFlagCheck.checked = flagged;
        redFlagReason.hidden = !flagged;
        if (document.activeElement !== redFlagReason) redFlagReason.value = flagReason(s);
    }

    critBlock.hidden = overview;
    overviewBlock.hidden = !overview;
    marksEl.hidden = overview;
    critNav.hidden = overview;
    overviewBtn.textContent = overview ? (overviewReturnSubpage ? 'RETOUR' : 'CRITÈRES') : 'GÉNÉRAL';
    // "‹ NOTES" appears only on a criterion page reached via a note's dot.
    backToNotesBtn.hidden = overview || noteReturnPos < 0;

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
    // Visible criteria dots, small, non-interactive.
    const dots = visibleCriteria().map((i) => {
        const c = model.criteria[i];
        return { bucket: bucketOf(markValue(s, c)), coef: c.coef, id: c.id };
    });
    requestAnimationFrame(() => drawDots(dotRow, dots, { height: 26, step: 0.25 }));

    // Per-group table (visible groups only).
    overviewTable.innerHTML = '';
    for (const gi of visibleGroupIdx()) {
        const g = model.groups[gi];
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
    noteBox.value = getNote(s, c);
    dirtyEl.hidden = !dirty;
}

function setDirty() { dirty = true; dirtyEl.hidden = false; }

// ---- Screens ----
function showScreen(name) {
    pickerScreen.hidden = name !== 'picker';
    gradingScreen.hidden = name !== 'grading';
    filterScreen.hidden = name !== 'filter';
    completionScreen.hidden = name !== 'completion';
    subpageScreen.hidden = name !== 'subpage';
    notesScreen.hidden = name !== 'notes';
    menuEl.hidden = true;
}
function changeFile() { fileInput.value = ''; fileInput.click(); }

// ---- Group filter ----
function groupIndexOfCriterion(ci) {
    for (let i = 0; i < model.groups.length; i++) {
        const g = model.groups[i];
        if (ci >= g.first && ci <= g.last) return i;
    }
    return -1;
}
function isCriterionVisible(ci) { const gi = groupIndexOfCriterion(ci); return gi < 0 || !groupHidden[gi]; }
function visibleGroupIdx() { const r = []; model.groups.forEach((g, i) => { if (!groupHidden[i]) r.push(i); }); return r; }
function visibleCriteria() { const r = []; model.criteria.forEach((c, i) => { if (isCriterionVisible(i)) r.push(i); }); return r; }

function filterKey() { return 'lapis-filter::' + sourceName; }
function loadFilter() {
    groupHidden = new Array(model.groups.length).fill(false);
    try {
        const raw = localStorage.getItem(filterKey());
        if (raw) {
            const names = JSON.parse(raw);
            model.groups.forEach((g, i) => { if (names.includes(g.name)) groupHidden[i] = true; });
        }
    } catch (e) { /* ignore */ }
    if (groupHidden.length && groupHidden.every(Boolean)) groupHidden.fill(false); // never all hidden
}
function saveFilter() {
    try {
        const names = model.groups.filter((g, i) => groupHidden[i]).map((g) => g.name);
        localStorage.setItem(filterKey(), JSON.stringify(names));
    } catch (e) { /* ignore */ }
}
function showFilter() { buildFilter(); showScreen('filter'); }
function buildFilter() {
    filterList.innerHTML = '';
    model.groups.forEach((g, i) => {
        const label = document.createElement('label');
        label.className = 'filter-item';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = !groupHidden[i];
        cb.addEventListener('change', () => { groupHidden[i] = !cb.checked; saveFilter(); });
        const span = document.createElement('span');
        span.textContent = g.name || '(groupe)';
        label.appendChild(cb);
        label.appendChild(span);
        filterList.appendChild(label);
    });
}
function setAllGroupsHidden(hidden) { groupHidden.fill(hidden); saveFilter(); buildFilter(); }
function closeFilter() {
    if (pageIdx !== 0 && !isCriterionVisible(pageIdx - 1)) pageIdx = 0; // fell into a hidden group
    showScreen('grading');
    render();
}

// ---- Completion ----
function computeCompletion() {
    const vg = visibleGroupIdx(), vc = visibleCriteria();
    const nStu = model.students.length, nVC = vc.length;
    const cc = {
        markedCriteria: 0, totalCriteria: nStu * nVC,
        markedGroups: 0, partialGroups: 0, unmarkedGroups: 0, totalGroups: nStu * vg.length,
        totalStudents: nStu, marked: [], partial: [], partialPct: [], unmarked: [],
    };
    for (let si = 0; si < nStu; si++) {
        const s = model.students[si];
        let sm = 0;
        for (const gi of vg) {
            const g = model.groups[gi], sz = g.last - g.first + 1;
            let gm = 0;
            for (let k = g.first; k <= g.last; k++) if (markValue(s, model.criteria[k]) != null) gm++;
            if (gm === sz) cc.markedGroups++; else if (gm > 0) cc.partialGroups++; else cc.unmarkedGroups++;
            sm += gm;
        }
        cc.markedCriteria += sm;
        if (nVC > 0 && sm === nVC) cc.marked.push(si);
        else if (sm > 0) { cc.partial.push(si); cc.partialPct.push(100 * sm / nVC); }
        else cc.unmarked.push(si);
    }
    return cc;
}
function showCompletion() { buildCompletion(); showScreen('completion'); }
function buildCompletion() {
    const cc = computeCompletion();
    completionContent.innerHTML = '';
    const add = (n) => completionContent.appendChild(n);
    add(compSection('totaux'));
    add(compRow('étudiants notés', cc.marked.length, cc.totalStudents, () => showSubpage('marked')));
    add(compRow('critères notés', cc.markedCriteria, cc.totalCriteria, null));
    add(compRow('groupes notés', cc.markedGroups, cc.totalGroups, null));
    add(compSection('notation partielle'));
    add(compRow('étudiants partiellement notés', cc.partial.length, cc.totalStudents, () => showSubpage('partial')));
    add(compRow('groupes partiellement notés', cc.partialGroups, cc.totalGroups, null));
    add(compSection('non noté'));
    add(compRow('étudiants non notés', cc.unmarked.length, cc.totalStudents, () => showSubpage('unmarked')));
    add(compRow('groupes non notés', cc.unmarkedGroups, cc.totalGroups, null));
}
function compSection(t) { const d = document.createElement('div'); d.className = 'comp-section'; d.textContent = t; return d; }
function compRow(label, count, total, onClick) {
    const row = document.createElement('div');
    row.className = 'comp-row';
    const l = document.createElement('span');
    l.className = 'comp-label' + (onClick ? ' link' : '');
    l.textContent = label;
    if (onClick) l.addEventListener('click', onClick);
    row.appendChild(l);
    const pct = total > 0 ? (100 * count / total) : 0;
    const v = document.createElement('span');
    v.className = 'comp-value';
    v.textContent = count + '/' + total + '   ' + pct.toFixed(1) + '%';
    row.appendChild(v);
    return row;
}
function showSubpage(kind) {
    subpageIsListe = false;
    const cc = computeCompletion();
    let list, title, pct = null;
    if (kind === 'marked') { list = cc.marked; title = 'étudiants notés'; }
    else if (kind === 'partial') { list = cc.partial; title = 'étudiants partiellement notés'; pct = cc.partialPct; }
    else { list = cc.unmarked; title = 'étudiants non notés'; }
    subpageTitle.textContent = title;
    subpageContent.innerHTML = '';
    if (!list.length) {
        const e = document.createElement('div'); e.className = 'sub-empty'; e.textContent = '(aucun)';
        subpageContent.appendChild(e);
    }
    list.forEach((si, k) => {
        const a = document.createElement('div');
        a.className = 'sub-item';
        setStudentLabel(a, si, pct ? ('   —   ' + pct[k].toFixed(1) + '%') : null);
        a.addEventListener('click', () => openStudentOverview(si, kind));
        subpageContent.appendChild(a);
    });
    showScreen('subpage');
}

// ---- Liste (full roster; each name links to that student's GENERAL page) ----
function showListe() {
    subpageIsListe = true;
    subpageTitle.textContent = 'liste des étudiants';
    subpageContent.innerHTML = '';
    if (!model.students.length) {
        const e = document.createElement('div'); e.className = 'sub-empty'; e.textContent = '(aucun)';
        subpageContent.appendChild(e);
    }
    model.students.forEach((s, si) => {
        const a = document.createElement('div');
        a.className = 'sub-item';
        setStudentLabel(a, si, null);
        a.addEventListener('click', () => openStudentOverview(si, 'liste'));
        subpageContent.appendChild(a);
    });
    showScreen('subpage');
}

// Open a student's GENERAL page from a subpage link, remembering which list to RETOUR to.
function openStudentOverview(si, kind) {
    studentIdx = si; pageIdx = 0;
    overviewReturnSubpage = kind || null;
    noteReturnPos = -1;
    showScreen('grading'); render();
}

// ---- Notes (stored on a "notes" sheet at the same (row, col) as the mark cell) ----
function getNote(student, criterion) {
    const s = workbook.getWorksheet(NOTES_SHEET);
    if (!s) return '';
    const v = s.getRow(student.row).getCell(criterion.col).value;
    return v == null ? '' : (typeof v === 'string' ? v : cellText(v));
}
function setNote(student, criterion, text) {
    const t = (text || '').trim();
    let s = workbook.getWorksheet(NOTES_SHEET);
    if (!s) { if (!t) return; s = workbook.addWorksheet(NOTES_SHEET); }
    s.getRow(student.row).getCell(criterion.col).value = t ? t : null;
    setDirty();
}
// ---- Red flag (stored on the "notes" sheet at the student-name column) ----
function readFlagRaw(student) {
    const s = workbook.getWorksheet(NOTES_SHEET);
    if (!s) return '';
    const v = s.getRow(student.row).getCell(FLAG_COL).value;
    return v == null ? '' : (typeof v === 'string' ? v : cellText(v));
}
// Flagged = the ⚑ marker is present; a cell holding only kept reason text reads as unflagged.
function isFlagged(student) { return readFlagRaw(student).startsWith(FLAG_MARK); }
function flagReason(student) {
    const raw = readFlagRaw(student);
    if (!raw) return '';
    return raw.startsWith(FLAG_MARK) ? raw.slice(FLAG_MARK.length).trim() : raw.trim();
}
// When off, keep the reason text but drop the marker (blank the cell only if there's no reason).
function setFlag(student, on, reason) {
    const r = (reason || '').trim();
    const val = on ? (r ? (FLAG_MARK + ' ' + r) : FLAG_MARK) : (r || null);
    let s = workbook.getWorksheet(NOTES_SHEET);
    if (val == null) {
        if (s) s.getRow(student.row).getCell(FLAG_COL).value = null;
        setDirty();
        return;
    }
    if (!s) s = workbook.addWorksheet(NOTES_SHEET);
    s.getRow(student.row).getCell(FLAG_COL).value = val;
    setDirty();
}
// Set an element's content to a student's name, with a trailing red dot when flagged, plus an
// optional plain suffix (e.g. the partial "— 42.0%").
function setStudentLabel(node, si, suffix) {
    const s = model.students[si];
    node.textContent = '';
    node.appendChild(document.createTextNode(s.name));
    if (isFlagged(s)) {
        const dot = document.createElement('span');
        dot.className = 'flag-dot';
        dot.textContent = ' ●';
        node.appendChild(dot);
    }
    if (suffix) node.appendChild(document.createTextNode(suffix));
}

function buildNotesList() {
    notesList = [];
    if (!workbook.getWorksheet(NOTES_SHEET)) return;
    for (let si = 0; si < model.students.length; si++) {
        for (let ci = 0; ci < model.criteria.length; ci++) {
            if (getNote(model.students[si], model.criteria[ci])) notesList.push({ si, ci });
        }
    }
}
function showNotes() {
    buildNotesList();
    if (notePos < 0 || notePos >= notesList.length) notePos = 0;
    renderNote();
    showScreen('notes');
}
function renderNote() {
    const has = notesList.length > 0;
    noteEmpty.hidden = has;
    noteBody.hidden = !has;
    el('prevNote').disabled = notesList.length < 2;
    el('nextNote').disabled = notesList.length < 2;
    if (!has) { noteCounter.textContent = ''; return; }
    if (notePos >= notesList.length) notePos = notesList.length - 1;
    const { si, ci } = notesList[notePos];
    const s = model.students[si], c = model.criteria[ci];
    noteCounter.textContent = '· note ' + (notePos + 1) + ' / ' + notesList.length;
    setStudentLabel(noteStudent, si, null);
    noteGroupEl.textContent = c.group || '';
    noteCrit.textContent = 'critère ' + c.id + (c.contract ? ' · ' + c.contract : '');
    noteEdit.value = getNote(s, c);
    // Enlarged dot (2×) over thin reference rings at every possible weight.
    const dot = [{ bucket: bucketOf(markValue(s, c)), coef: c.coef, id: c.id }];
    requestAnimationFrame(() => drawDots(noteDotEl, dot, {
        height: 120, step: 0.33, sizeMult: 2, showLabels: true, referenceScales: weightReferenceScales(),
    }));
}
function navigateNote(delta) {
    buildNotesList(); // a note may have been emptied via the editor
    if (!notesList.length) { renderNote(); return; }
    const n = notesList.length;
    notePos = ((notePos + delta) % n + n) % n;
    renderNote();
}
function gotoNoteCriterion() {
    if (!notesList.length) return;
    const { si, ci } = notesList[notePos];
    noteReturnPos = notePos;
    overviewReturnSubpage = null;
    studentIdx = si;
    pageIdx = ci + 1;
    showScreen('grading');
    render();
}
function weightReferenceScales() {
    const set = new Set();
    model.criteria.forEach((c) => { if (c.coef != null) set.add(c.coef); });
    return [...set].sort((a, b) => a - b).map((cf) => coefScale(cf, 0.33));
}

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

// Best-effort portrait lock. Works on Android installed/fullscreen PWAs; iOS ignores it, where
// the CSS #rotate-notice covers landscape instead. manifest.webmanifest also requests portrait.
if (screen.orientation && screen.orientation.lock) {
    try { screen.orientation.lock('portrait').catch(() => {}); } catch (e) { /* unsupported */ }
}
