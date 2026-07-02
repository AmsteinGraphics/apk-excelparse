package com.example.apkexcelparse;

import android.content.ContentResolver;
import android.content.ContentUris;
import android.content.ContentValues;
import android.content.Intent;
import android.content.SharedPreferences;
import android.database.Cursor;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.os.Handler;
import android.os.Looper;
import android.provider.MediaStore;
import android.provider.OpenableColumns;
import android.provider.Settings;
import android.view.Gravity;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.widget.PopupMenu;
import android.widget.ProgressBar;
import android.widget.TableLayout;
import android.widget.TableRow;
import android.widget.TextView;
import android.widget.Toast;

import androidx.core.content.FileProvider;

import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;

import com.example.apkexcelparse.model.Criterion;
import com.example.apkexcelparse.model.GradingModel;
import com.example.apkexcelparse.model.Group;
import com.example.apkexcelparse.model.Student;
import com.example.apkexcelparse.ui.DotProgressView;
import com.example.apkexcelparse.xlsx.XlsxParser;
import com.google.android.material.button.MaterialButton;
import com.google.android.material.color.MaterialColors;
import android.content.res.ColorStateList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Locale;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class MainActivity extends AppCompatActivity {

    private static final String PREFS = "apk_excelparse";
    private static final String KEY_HAS_WORKBOOK = "has_workbook";
    private static final String KEY_DISPLAY_NAME = "display_name";
    private static final String WORKING_COPY_NAME = "working_copy.xlsx";
    private static final String XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    private static final String RELEASES_API = "https://api.github.com/repos/AmsteinGraphics/apk-excelparse/releases/latest";
    private static final Pattern SHA_PATTERN = Pattern.compile("[0-9a-f]{40}");

    private final ExecutorService executor = Executors.newSingleThreadExecutor();
    private final Handler mainHandler = new Handler(Looper.getMainLooper());

    private View pickerContainer;
    private View gradingContainer;
    private ProgressBar loadingIndicator;
    private TextView studentCounter;
    private TextView studentName;
    private TextView criterionGroup;
    private TextView criterionIdCoef;
    private TextView criterionContract;
    private TextView criterionRemarks;
    private View markButtonsContainer;
    private final MaterialButton[] markButtons = new MaterialButton[5];
    private MaterialButton clearMarkButton;
    private MaterialButton prevButton;
    private MaterialButton nextButton;
    private MaterialButton prevStudentButton;
    private MaterialButton nextStudentButton;
    private MaterialButton prevLetterButton;
    private MaterialButton nextLetterButton;
    private MaterialButton prevGroupButton;
    private MaterialButton nextGroupButton;
    private MaterialButton overviewButton;
    private MaterialButton saveButton;
    private View dirtyIndicator;
    private TextView studentAverage;
    private DotProgressView dotProgressView;
    private TableLayout overviewTable;

    private XSSFWorkbook workbook;
    private FormulaEvaluator formulaEvaluator;
    private GradingModel model;
    private int studentIdx;
    // Per-student page index: 0 = overview (general average + all dots, no grading),
    // 1..criteria.size() = criterion pages (criterion index = pageIdx - 1).
    private int pageIdx;
    // When the OVERVIEW button jumps to the overview page it stashes the criterion page to
    // return to here (>= 1). -1 = no pending return, so the button shows "OVERVIEW".
    private int overviewReturnPage = -1;
    private boolean dirty;

    // Mark values for the 5 buttons, left→right (0 … 1).
    private static final float[] MARK_VALUES = {0f, 0.25f, 0.5f, 0.75f, 1f};

    private final ActivityResultLauncher<String[]> pickFileLauncher =
            registerForActivityResult(new ActivityResultContracts.OpenDocument(), this::onFilePicked);

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        pickerContainer = findViewById(R.id.pickerContainer);
        gradingContainer = findViewById(R.id.gradingContainer);
        loadingIndicator = findViewById(R.id.loadingIndicator);
        studentCounter = findViewById(R.id.studentCounter);
        studentName = findViewById(R.id.studentName);
        criterionGroup = findViewById(R.id.criterionGroup);
        criterionIdCoef = findViewById(R.id.criterionIdCoef);
        criterionContract = findViewById(R.id.criterionContract);
        criterionRemarks = findViewById(R.id.criterionRemarks);
        markButtonsContainer = findViewById(R.id.markButtonsContainer);
        markButtons[0] = findViewById(R.id.btnMark0);
        markButtons[1] = findViewById(R.id.btnMark1);
        markButtons[2] = findViewById(R.id.btnMark2);
        markButtons[3] = findViewById(R.id.btnMark3);
        markButtons[4] = findViewById(R.id.btnMark4);
        clearMarkButton = findViewById(R.id.btnMarkClear);
        prevButton = findViewById(R.id.prevButton);
        nextButton = findViewById(R.id.nextButton);
        prevStudentButton = findViewById(R.id.prevStudentButton);
        nextStudentButton = findViewById(R.id.nextStudentButton);
        prevLetterButton = findViewById(R.id.prevLetterButton);
        nextLetterButton = findViewById(R.id.nextLetterButton);
        prevGroupButton = findViewById(R.id.prevGroupButton);
        nextGroupButton = findViewById(R.id.nextGroupButton);
        overviewButton = findViewById(R.id.overviewButton);
        saveButton = findViewById(R.id.saveButton);
        dirtyIndicator = findViewById(R.id.dirtyIndicator);
        studentAverage = findViewById(R.id.studentAverage);
        dotProgressView = findViewById(R.id.dotProgressView);
        overviewTable = findViewById(R.id.overviewTable);

        findViewById(R.id.pickButton).setOnClickListener(v -> launchPicker());
        findViewById(R.id.pickerCheckUpdatesButton).setOnClickListener(v -> checkForUpdates());
        findViewById(R.id.menuButton).setOnClickListener(this::showOverflowMenu);
        prevButton.setOnClickListener(v -> navigate(-1));
        nextButton.setOnClickListener(v -> navigate(1));
        prevStudentButton.setOnClickListener(v -> navigateStudent(-1));
        nextStudentButton.setOnClickListener(v -> navigateStudent(1));
        prevLetterButton.setOnClickListener(v -> navigateLetter(-1));
        nextLetterButton.setOnClickListener(v -> navigateLetter(1));
        prevGroupButton.setOnClickListener(v -> navigateGroup(-1));
        nextGroupButton.setOnClickListener(v -> navigateGroup(1));
        overviewButton.setOnClickListener(v -> toggleOverview());
        saveButton.setOnClickListener(v -> {
            if (!dirty) {
                Toast.makeText(this, "Nothing to save", Toast.LENGTH_SHORT).show();
                return;
            }
            saveAsync(null);
        });

        for (int i = 0; i < markButtons.length; i++) {
            final int idx = i;
            markButtons[i].setOnClickListener(v -> onMarkButtonClicked(idx));
        }
        clearMarkButton.setOnClickListener(v -> onClearMark());

        SharedPreferences prefs = getSharedPreferences(PREFS, MODE_PRIVATE);
        if (prefs.getBoolean(KEY_HAS_WORKBOOK, false) && workingCopyFile().exists()) {
            loadWorkbookAsync();
        } else {
            showPicker();
        }
    }

    // onPause intentionally does NOT autosave. The workbook stays in memory
    // until the user hits Save. Trade-off: if Android kills the process, in-memory
    // marks are lost. The dirty indicator makes pending state visible.

    private void showOverflowMenu(View anchor) {
        PopupMenu menu = new PopupMenu(this, anchor);
        menu.getMenu().add(0, 1, 0, R.string.change_file);
        menu.getMenu().add(0, 2, 1, R.string.check_updates);
        menu.setOnMenuItemClickListener(item -> {
            if (item.getItemId() == 1) { launchPicker(); return true; }
            if (item.getItemId() == 2) { checkForUpdates(); return true; }
            return false;
        });
        menu.show();
    }

    private void launchPicker() {
        if (dirty) {
            new androidx.appcompat.app.AlertDialog.Builder(this)
                    .setTitle("Discard unsaved marks?")
                    .setMessage("You have unsaved marks in the current workbook. Picking a new file will discard them.")
                    .setPositiveButton("Discard", (d, w) -> doLaunchPicker())
                    .setNegativeButton(android.R.string.cancel, null)
                    .show();
        } else {
            doLaunchPicker();
        }
    }

    private void doLaunchPicker() {
        pickFileLauncher.launch(new String[]{
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/octet-stream"
        });
    }

    private void onFilePicked(@Nullable Uri uri) {
        if (uri == null) return;
        setLoading(true);
        executor.execute(() -> {
            try {
                String displayName = queryDisplayName(uri);
                File dest = workingCopyFile();
                try (InputStream in = getContentResolver().openInputStream(uri);
                     OutputStream out = new FileOutputStream(dest)) {
                    if (in == null) throw new IllegalStateException("Could not open picked file");
                    byte[] buf = new byte[16 * 1024];
                    int n;
                    while ((n = in.read(buf)) > 0) out.write(buf, 0, n);
                }
                getSharedPreferences(PREFS, MODE_PRIVATE).edit()
                        .putBoolean(KEY_HAS_WORKBOOK, true)
                        .putString(KEY_DISPLAY_NAME, displayName)
                        .apply();
                mainHandler.post(this::loadWorkbookAsync);
            } catch (Exception e) {
                mainHandler.post(() -> {
                    setLoading(false);
                    Toast.makeText(this, "Could not read file: " + e.getMessage(), Toast.LENGTH_LONG).show();
                });
            }
        });
    }

    private String queryDisplayName(Uri uri) {
        try (Cursor c = getContentResolver().query(uri,
                new String[]{OpenableColumns.DISPLAY_NAME}, null, null, null)) {
            if (c != null && c.moveToFirst()) {
                int idx = c.getColumnIndex(OpenableColumns.DISPLAY_NAME);
                if (idx >= 0) return c.getString(idx);
            }
        } catch (Exception ignored) {
        }
        return null;
    }

    private void loadWorkbookAsync() {
        setLoading(true);
        executor.execute(() -> {
            try {
                closeWorkbookQuietly();
                XSSFWorkbook wb;
                try (FileInputStream in = new FileInputStream(workingCopyFile())) {
                    wb = XlsxParser.open(in);
                }
                GradingModel m = XlsxParser.parse(wb);
                if (m.students.isEmpty() || m.criteria.isEmpty()) {
                    throw new IllegalStateException("No students or criteria found. Check sheet layout and cell colors.");
                }
                mainHandler.post(() -> {
                    workbook = wb;
                    formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
                    model = m;
                    studentIdx = 0;
                    pageIdx = 0;
                    overviewReturnPage = -1;
                    dirty = false;
                    setLoading(false);
                    showGrading();
                    render();
                    updateDirtyUi();
                });
            } catch (Exception e) {
                mainHandler.post(() -> {
                    setLoading(false);
                    showPicker();
                    showErrorDialog("Could not load workbook", e.getMessage());
                });
            }
        });
    }

    private void showErrorDialog(String title, String message) {
        new androidx.appcompat.app.AlertDialog.Builder(this)
                .setTitle(title)
                .setMessage(message != null ? message : "(no details)")
                .setPositiveButton(android.R.string.ok, null)
                .show();
    }

    // Pages per student: one overview page (index 0) plus one page per criterion.
    private int pagesPerStudent() {
        return model.criteria.size() + 1;
    }

    private void navigate(int delta) {
        if (model == null) return;
        // Moving between criteria abandons any stashed "return to criterion" from OVERVIEW.
        overviewReturnPage = -1;
        int perStudent = pagesPerStudent();
        int total = model.students.size() * perStudent;
        int flat = studentIdx * perStudent + pageIdx + delta;
        if (flat < 0) flat = 0;
        if (flat >= total) flat = total - 1;
        studentIdx = flat / perStudent;
        pageIdx = flat % perStudent;
        render();
    }

    private void navigateStudent(int delta) {
        if (model == null) return;
        int size = model.students.size();
        // Wrap around the roster; keeps the same page (overview stays overview) across students.
        int newIdx = ((studentIdx + delta) % size + size) % size;
        if (newIdx == studentIdx) return;
        studentIdx = newIdx;
        render();
    }

    /**
     * Jump to the first student whose name starts with the previous/next starting letter that
     * actually occurs in the roster, wrapping around the alphabet. Keeps the current page index.
     */
    private void navigateLetter(int delta) {
        if (model == null || model.students.isEmpty()) return;
        java.util.List<Character> letters = new java.util.ArrayList<>();
        for (Student st : model.students) {
            char c = firstLetter(st.name);
            if (!letters.contains(c)) letters.add(c);
        }
        java.util.Collections.sort(letters);
        int cur = letters.indexOf(firstLetter(model.students.get(studentIdx).name));
        if (cur < 0) return;
        int n = letters.size();
        char target = letters.get(((cur + delta) % n + n) % n);
        for (int i = 0; i < model.students.size(); i++) {
            if (firstLetter(model.students.get(i).name) == target) {
                if (i == studentIdx) return;
                studentIdx = i;
                render();
                return;
            }
        }
    }

    /** First alphabetic character of a name, upper-cased; '#' if there is none. */
    private static char firstLetter(String name) {
        if (name != null) {
            for (int i = 0; i < name.length(); i++) {
                char c = name.charAt(i);
                if (Character.isLetter(c)) return Character.toUpperCase(c);
            }
        }
        return '#';
    }

    /** Jump to the first criterion page of the previous/next group. */
    private void navigateGroup(int delta) {
        if (model == null || model.groups.isEmpty()) return;
        // Landing on a criterion page abandons any stashed "return to criterion" from OVERVIEW.
        overviewReturnPage = -1;
        int curGroupIdx = isOverviewPage() ? -1 : indexOfGroupForCriterion(pageIdx - 1);
        int target = curGroupIdx + delta;
        if (target < 0) target = 0;
        if (target >= model.groups.size()) target = model.groups.size() - 1;
        pageIdx = model.groups.get(target).firstCriterionIndex + 1;
        render();
    }

    /** Jump straight to a group's first criterion page (used by the overview table name links). */
    private void jumpToGroup(int firstCriterionIndex) {
        if (model == null) return;
        overviewReturnPage = -1; // deliberate jump to a criterion; abandon any GENERAL stash
        pageIdx = firstCriterionIndex + 1;
        render();
    }

    private int indexOfGroupForCriterion(int criterionIndex) {
        for (int i = 0; i < model.groups.size(); i++) {
            Group g = model.groups.get(i);
            if (criterionIndex >= g.firstCriterionIndex && criterionIndex <= g.lastCriterionIndex) {
                return i;
            }
        }
        return -1;
    }

    /**
     * OVERVIEW button. From a criterion page: stash the page and jump to the overview so the
     * teacher can see the effect on the general grade. From that stashed state: return ("BACK").
     */
    private void toggleOverview() {
        if (model == null) return;
        if (overviewReturnPage >= 0) {
            pageIdx = overviewReturnPage;
            overviewReturnPage = -1;
            render();
        } else if (!isOverviewPage()) {
            overviewReturnPage = pageIdx;
            pageIdx = 0;
            render();
        }
        // Already on the overview with nothing stashed: no-op.
    }

    private void updateOverviewButton() {
        if (overviewButton == null) return;
        overviewButton.setText(overviewReturnPage >= 0
                ? R.string.overview_back : R.string.overview_button);
    }

    private boolean isOverviewPage() {
        return pageIdx == 0;
    }

    private void render() {
        if (model == null) return;
        Student s = model.students.get(studentIdx);

        studentCounter.setText(String.format(Locale.getDefault(),
                "étudiant %d / %d", studentIdx + 1, model.students.size()));
        studentName.setText(s.name);

        if (isOverviewPage()) {
            renderOverview(s);
        } else {
            renderCriterion(model.criteria.get(pageIdx - 1), s);
        }

        updateOverviewButton();
        refreshDots();
        refreshAverage();
    }

    /** Overview page: general average + all dots + per-group table, no grading control. */
    private void renderOverview(Student s) {
        criterionGroup.setText("vue générale");
        criterionGroup.setVisibility(View.VISIBLE);
        criterionIdCoef.setVisibility(View.GONE);
        criterionContract.setText("moyenne générale et progression de l'étudiant");
        criterionContract.setVisibility(View.VISIBLE);
        criterionRemarks.setVisibility(View.GONE);
        markButtonsContainer.setVisibility(View.GONE);
        buildOverviewTable(s);
        overviewTable.setVisibility(View.VISIBLE);
    }

    /**
     * Build the per-group overview table: one borderless row per group, each with the group's
     * dots (same physical size and coefficient scaling as the header), the group name, and the
     * group's grade (mark only). Rebuilt on every overview render so marks stay current.
     */
    private void buildOverviewTable(Student s) {
        overviewTable.removeAllViews();
        if (model.criteria.isEmpty()) return;
        // Let the name column (index 1) shrink/wrap so a long name can't push the average column
        // off the right edge in portrait. No stretch — columns stay compact and left-packed.
        overviewTable.setColumnShrinkable(1, true);
        float slotPx = headerSlotPx();
        int rowH = Math.round(dpToPx(33f));
        // Fixed-width dots column = widest group's dot string, so every group name starts at the
        // same x (dots stay left-packed inside it).
        int maxDots = 1;
        for (Group g : model.groups) {
            maxDots = Math.max(maxDots, g.lastCriterionIndex - g.firstCriterionIndex + 1);
        }
        int dotsCellW = Math.round(slotPx * maxDots);
        for (Group g : model.groups) {
            TableRow row = new TableRow(this);

            int n = g.lastCriterionIndex - g.firstCriterionIndex + 1;
            int[] buckets = new int[n];
            float[] scales = new float[n];
            for (int i = 0; i < n; i++) {
                Criterion c = model.criteria.get(g.firstCriterionIndex + i);
                buckets[i] = markToBucket(XlsxParser.readMark(workbook, s, c));
                scales[i] = coefScale(c.coefficient);
            }
            DotProgressView dots = new DotProgressView(this);
            dots.setFixedSlotPx(slotPx);
            dots.setValues(buckets);
            dots.setScales(scales);
            TableRow.LayoutParams dotsLp = new TableRow.LayoutParams(dotsCellW, rowH);
            dotsLp.gravity = Gravity.CENTER_VERTICAL;
            row.addView(dots, dotsLp);

            // Group name is a link that jumps straight to that group's first criterion page.
            TextView nameCell = overviewCell(g.name != null ? g.name : "", false);
            final int firstCriterionIndex = g.firstCriterionIndex;
            nameCell.setTextColor(0xFF1976D2);
            nameCell.setPaintFlags(nameCell.getPaintFlags() | android.graphics.Paint.UNDERLINE_TEXT_FLAG);
            nameCell.setClickable(true);
            nameCell.setFocusable(true);
            android.util.TypedValue ripple = new android.util.TypedValue();
            getTheme().resolveAttribute(android.R.attr.selectableItemBackground, ripple, true);
            nameCell.setBackgroundResource(ripple.resourceId);
            int pad = Math.round(dpToPx(8f)); // setBackgroundResource can reset padding; re-apply.
            nameCell.setPadding(pad, 0, pad, 0);
            nameCell.setOnClickListener(v -> jumpToGroup(firstCriterionIndex));
            row.addView(nameCell);

            Double avg = XlsxParser.readNumericAt(workbook, s, g.averageColumnIndex, formulaEvaluator);
            String avgText = (avg == null || avg.isNaN())
                    ? "" : String.format(Locale.getDefault(), "%.2f", avg);
            row.addView(overviewCell(avgText, true));

            overviewTable.addView(row);
        }
    }

    private TextView overviewCell(String text, boolean bold) {
        TextView tv = new TextView(this);
        tv.setText(text);
        tv.setTextSize(11f);
        tv.setGravity(Gravity.CENTER_VERTICAL);
        int pad = Math.round(dpToPx(8f));
        tv.setPadding(pad, 0, pad, 0);
        if (bold) {
            tv.setTypeface(tv.getTypeface(), android.graphics.Typeface.BOLD);
            tv.setTextColor(0xFF1976D2);
        }
        TableRow.LayoutParams lp = new TableRow.LayoutParams(
                TableRow.LayoutParams.WRAP_CONTENT, TableRow.LayoutParams.MATCH_PARENT);
        lp.gravity = Gravity.CENTER_VERTICAL;
        tv.setLayoutParams(lp);
        return tv;
    }

    /** Slot width (px) of one dot in the full-width header string of all criteria. */
    private float headerSlotPx() {
        float contentPx = getResources().getDisplayMetrics().widthPixels - dpToPx(48f);
        int n = Math.max(1, model.criteria.size());
        return contentPx / n;
    }

    /** Dot radius multiplier from a criterion coefficient: coef 2 → 1.0, ±0.25 per unit (overview). */
    private static float coefScale(Double coef) {
        return coefScale(coef, 0.25f);
    }

    /**
     * Dot radius multiplier: coef 2 → 1.0, ±{@code step} per unit, clamped [0.4, 2.0].
     * Criterion pages use a wider 0.33 step so the coefficient weight reads more strongly there.
     */
    private static float coefScale(Double coef, float step) {
        float c = coef != null ? coef.floatValue() : 2f;
        float scale = 1f + (c - 2f) * step;
        if (scale < 0.4f) scale = 0.4f;
        if (scale > 2f) scale = 2f;
        return scale;
    }

    private float dpToPx(float dp) {
        return dp * getResources().getDisplayMetrics().density;
    }

    private void renderCriterion(Criterion c, Student s) {
        overviewTable.setVisibility(View.GONE);
        criterionGroup.setVisibility(View.VISIBLE);
        criterionIdCoef.setVisibility(View.VISIBLE);
        criterionContract.setVisibility(View.VISIBLE);
        criterionRemarks.setVisibility(View.VISIBLE);
        markButtonsContainer.setVisibility(View.VISIBLE);

        criterionGroup.setText(c.groupName != null ? c.groupName : "");
        String coefText = c.coefficient != null
                ? String.format(Locale.getDefault(), "critère %s  ·  coefficient %s", c.id, formatCoef(c.coefficient))
                : String.format(Locale.getDefault(), "critère %s", c.id);
        criterionIdCoef.setText(coefText);
        criterionContract.setText(c.contract != null ? c.contract : "");
        criterionRemarks.setText(c.remarks != null ? c.remarks : "");

        Double existing = XlsxParser.readMark(workbook, s, c);
        // No stored mark → nothing selected (all buttons shaded).
        refreshMarkButtons(existing != null ? markToBucket(existing) : -1);
    }

    /** A mark button was tapped: write its value and reflect the new selection. */
    private void onMarkButtonClicked(int index) {
        if (isOverviewPage()) return;
        applyMarkToWorkbook(MARK_VALUES[index]);
        refreshMarkButtons(index);
    }

    /** The "×" button: clear the mark for this criterion so it reads as ungraded (all blank). */
    private void onClearMark() {
        if (model == null || workbook == null || isOverviewPage()) return;
        Student s = model.students.get(studentIdx);
        Criterion c = model.criteria.get(pageIdx - 1);
        if (XlsxParser.eraseMark(workbook, s, c)) {
            dirty = true;
            updateDirtyUi();
            if (formulaEvaluator != null) {
                Cell cell = XlsxParser.getEvalCell(workbook, s, c.columnIndex);
                if (cell != null) formulaEvaluator.notifyUpdateCell(cell);
            }
            refreshDots();
            refreshAverage();
        }
        refreshMarkButtons(-1);
    }

    /**
     * Selected mark button is filled blue (theme primary) with white text; the rest stay outlined
     * (transparent fill, primary-coloured text) like the other buttons. -1 = nothing selected.
     */
    private void refreshMarkButtons(int selectedIndex) {
        int primary = MaterialColors.getColor(this,
                androidx.appcompat.R.attr.colorPrimary, 0xFF1976D2);
        for (int i = 0; i < markButtons.length; i++) {
            boolean selected = i == selectedIndex;
            markButtons[i].setBackgroundTintList(
                    ColorStateList.valueOf(selected ? primary : 0x00000000));
            markButtons[i].setTextColor(selected ? 0xFFFFFFFF : primary);
        }
    }

    /**
     * Refresh the dot row for the current page: all criteria on the overview page,
     * only the current criterion's group on a criterion page.
     */
    private void refreshDots() {
        if (model == null || workbook == null || dotProgressView == null) return;
        Student s = model.students.get(studentIdx);
        int first;
        int last; // inclusive criterion indices to display
        if (isOverviewPage() || model.criteria.isEmpty()) {
            first = 0;
            last = model.criteria.size() - 1;
        } else {
            Group g = model.groupForCriterion(pageIdx - 1);
            if (g != null) {
                first = g.firstCriterionIndex;
                last = g.lastCriterionIndex;
            } else {
                first = last = pageIdx - 1;
            }
        }
        int n = last - first + 1;
        int[] buckets = new int[n];
        float[] scales = new float[n];
        String[] labels = new String[n];
        for (int i = 0; i < n; i++) {
            Criterion c = model.criteria.get(first + i);
            buckets[i] = markToBucket(XlsxParser.readMark(workbook, s, c));
            scales[i] = coefScale(c.coefficient, 0.33f);
            labels[i] = coefDigit(c.coefficient);
        }
        // Criterion pages need a taller row so the heaviest dots (coef 4 → 2× the coef-2 size)
        // render full-size instead of being clamped; the overview's tiny all-criteria dots stay 33dp.
        int rowDp = isOverviewPage() ? 33 : 52;
        android.view.ViewGroup.LayoutParams lp = dotProgressView.getLayoutParams();
        int rowPx = Math.round(dpToPx(rowDp));
        if (lp.height != rowPx) {
            lp.height = rowPx;
            dotProgressView.setLayoutParams(lp);
        }
        dotProgressView.setValues(buckets);
        dotProgressView.setScales(scales);
        // Coefficient digit inside each dot on criterion pages only; the overview stays plain.
        dotProgressView.setLabels(isOverviewPage() ? null : labels);
        // Highlight the current criterion's dot on criterion pages; none on the overview.
        dotProgressView.setHighlightIndex(isOverviewPage() ? -1 : (pageIdx - 1) - first);
    }

    /** Compact coefficient string for a dot: "2" for integers, trailing zeros trimmed otherwise. */
    private static String coefDigit(Double coef) {
        if (coef == null) return null;
        if (coef == Math.floor(coef)) return String.valueOf(coef.intValue());
        String s = String.format(Locale.getDefault(), "%.2f", coef);
        while (s.endsWith("0")) s = s.substring(0, s.length() - 1);
        if (s.endsWith(".") || s.endsWith(",")) s = s.substring(0, s.length() - 1);
        return s;
    }

    /**
     * Refresh the average shown top-right: the overall student average (column CL) on the
     * overview page, or the current group's "sur 6" grade on a criterion page. Both are /6.
     */
    private void refreshAverage() {
        if (model == null || workbook == null || studentAverage == null) return;
        Student s = model.students.get(studentIdx);
        int col = XlsxParser.COL_STUDENT_AVERAGE;
        if (!isOverviewPage()) {
            Group g = model.groupForCriterion(pageIdx - 1);
            if (g != null) col = g.averageColumnIndex;
        }
        Double avg = XlsxParser.readNumericAt(workbook, s, col, formulaEvaluator);
        if (avg == null || avg.isNaN()) {
            studentAverage.setText("");
        } else {
            studentAverage.setText(String.format(Locale.getDefault(), "%.2f", avg));
        }
    }

    private static int markToBucket(Double value) {
        if (value == null) return -1;
        int b = (int) Math.round(value * 4);
        if (b < 0) b = 0;
        if (b > 4) b = 4;
        return b;
    }

    private void applyMarkToWorkbook(float value) {
        if (model == null || workbook == null || isOverviewPage()) return;
        Student s = model.students.get(studentIdx);
        Criterion c = model.criteria.get(pageIdx - 1);
        boolean written = XlsxParser.writeMark(workbook, s, c, value);
        if (written) {
            dirty = true;
            updateDirtyUi();
            if (formulaEvaluator != null) {
                Cell cell = XlsxParser.getEvalCell(workbook, s, c.columnIndex);
                if (cell != null) formulaEvaluator.notifyUpdateCell(cell);
            }
            refreshDots();
            refreshAverage();
        } else {
            Toast.makeText(this, "Cell not writable (not a mark cell)", Toast.LENGTH_SHORT).show();
        }
    }

    private void updateDirtyUi() {
        if (dirtyIndicator != null) {
            dirtyIndicator.setVisibility(dirty ? View.VISIBLE : View.GONE);
        }
    }

    private final java.util.concurrent.atomic.AtomicBoolean saveEnqueued =
            new java.util.concurrent.atomic.AtomicBoolean(false);

    private void saveAsync(@Nullable Runnable then) {
        if (workbook == null) {
            if (then != null) then.run();
            return;
        }
        if (then == null) {
            // Fire-and-forget: at most one pending save at a time; later changes ride the next save.
            if (!saveEnqueued.compareAndSet(false, true)) return;
            executor.execute(() -> {
                saveEnqueued.set(false);
                performSave(null);
            });
        } else {
            executor.execute(() -> performSave(then));
        }
    }

    private void performSave(@Nullable Runnable then) {
        try {
            try (FileOutputStream out = new FileOutputStream(workingCopyFile())) {
                workbook.write(out);
            }
            String exportedName = null;
            String exportError = null;
            try {
                exportedName = exportToDownloads(workingCopyFile());
            } catch (Exception e) {
                exportError = e.getMessage();
            }
            final String finalExportedName = exportedName;
            final String finalExportError = exportError;
            mainHandler.post(() -> {
                dirty = false;
                updateDirtyUi();
                if (finalExportedName != null) {
                    Toast.makeText(this,
                            "Saved to Downloads/" + finalExportedName,
                            Toast.LENGTH_LONG).show();
                } else if (finalExportError != null) {
                    Toast.makeText(this,
                            "Saved locally; export to Downloads failed: " + finalExportError,
                            Toast.LENGTH_LONG).show();
                } else {
                    Toast.makeText(this, "Saved", Toast.LENGTH_SHORT).show();
                }
                if (then != null) then.run();
            });
        } catch (Exception e) {
            mainHandler.post(() -> {
                Toast.makeText(this, "Save failed: " + e.getMessage(), Toast.LENGTH_LONG).show();
                if (then != null) then.run();
            });
        }
    }

    /**
     * Copy the current workbook file to public Downloads under a "-graded" filename,
     * overwriting any existing copy with the same name. Returns the filename used, or
     * null if the platform lacks scoped MediaStore Downloads.
     */
    private String exportToDownloads(File src) throws Exception {
        if (Build.VERSION.SDK_INT < Build.VERSION_CODES.Q) {
            return null;
        }
        String filename = buildExportFilename();
        ContentResolver cr = getContentResolver();
        Uri existing = findInDownloads(filename);
        Uri target = existing;
        if (target == null) {
            ContentValues values = new ContentValues();
            values.put(MediaStore.MediaColumns.DISPLAY_NAME, filename);
            values.put(MediaStore.MediaColumns.MIME_TYPE, XLSX_MIME);
            values.put(MediaStore.MediaColumns.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS);
            target = cr.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, values);
            if (target == null) throw new Exception("insert returned null");
        }
        try (InputStream in = new FileInputStream(src);
             OutputStream out = cr.openOutputStream(target, "wt")) {
            if (out == null) throw new Exception("Could not open Downloads output");
            byte[] buf = new byte[16 * 1024];
            int n;
            while ((n = in.read(buf)) > 0) out.write(buf, 0, n);
        }
        return filename;
    }

    private Uri findInDownloads(String filename) {
        if (Build.VERSION.SDK_INT < Build.VERSION_CODES.Q) return null;
        Uri collection = MediaStore.Downloads.EXTERNAL_CONTENT_URI;
        String[] projection = {MediaStore.MediaColumns._ID};
        String selection = MediaStore.MediaColumns.DISPLAY_NAME + " = ?";
        String[] args = {filename};
        try (Cursor c = getContentResolver().query(collection, projection, selection, args, null)) {
            if (c != null && c.moveToFirst()) {
                return ContentUris.withAppendedId(collection, c.getLong(0));
            }
        } catch (Exception ignored) {
        }
        return null;
    }

    private String buildExportFilename() {
        String display = getSharedPreferences(PREFS, MODE_PRIVATE).getString(KEY_DISPLAY_NAME, null);
        if (display == null || display.isEmpty()) {
            return "graded.xlsx";
        }
        int dot = display.lastIndexOf('.');
        String base = dot > 0 ? display.substring(0, dot) : display;
        return base + "-graded.xlsx";
    }

    private void showPicker() {
        pickerContainer.setVisibility(View.VISIBLE);
        gradingContainer.setVisibility(View.GONE);
    }

    private void showGrading() {
        pickerContainer.setVisibility(View.GONE);
        gradingContainer.setVisibility(View.VISIBLE);
    }

    private void setLoading(boolean loading) {
        loadingIndicator.setVisibility(loading ? View.VISIBLE : View.GONE);
    }

    private File workingCopyFile() {
        return new File(getFilesDir(), WORKING_COPY_NAME);
    }

    private void closeWorkbookQuietly() {
        if (workbook != null) {
            try {
                workbook.close();
            } catch (Exception ignored) {
            }
            workbook = null;
        }
    }

    private static String formatCoef(Double coef) {
        if (coef == null) return "";
        if (coef == Math.floor(coef)) return String.valueOf(coef.intValue());
        return String.format(Locale.getDefault(), "%.2f", coef);
    }

    private void checkForUpdates() {
        Toast.makeText(this, "Checking for updates…", Toast.LENGTH_SHORT).show();
        executor.execute(() -> {
            try {
                URL url = new URL(RELEASES_API);
                HttpURLConnection connection = (HttpURLConnection) url.openConnection();
                connection.setRequestMethod("GET");
                connection.setRequestProperty("Accept", "application/vnd.github+json");
                connection.setRequestProperty("User-Agent", "apk-excelparse-app");
                int code = connection.getResponseCode();
                if (code != HttpURLConnection.HTTP_OK) throw new RuntimeException("HTTP " + code);
                StringBuilder builder = new StringBuilder();
                try (BufferedReader reader = new BufferedReader(new InputStreamReader(connection.getInputStream()))) {
                    String line;
                    while ((line = reader.readLine()) != null) builder.append(line);
                }
                JSONObject release = new JSONObject(builder.toString());
                Matcher m = SHA_PATTERN.matcher(release.optString("body", ""));
                String latestSha = m.find() ? m.group() : null;
                JSONArray assets = release.optJSONArray("assets");
                String apkUrl = null;
                if (assets != null) {
                    for (int i = 0; i < assets.length(); i++) {
                        JSONObject asset = assets.getJSONObject(i);
                        if (asset.optString("name", "").endsWith(".apk")) {
                            apkUrl = asset.optString("browser_download_url", null);
                            break;
                        }
                    }
                }
                if (latestSha == null || apkUrl == null) {
                    throw new RuntimeException("Release has no APK or SHA");
                }
                final String current = BuildConfig.BUILD_SHA;
                final String latest = latestSha;
                final String downloadUrl = apkUrl;
                mainHandler.post(() -> {
                    if (latest.equals(current)) {
                        Toast.makeText(this,
                                "Already up to date (" + shortSha(current) + ")",
                                Toast.LENGTH_SHORT).show();
                        return;
                    }
                    promptUpdate(current, latest, downloadUrl);
                });
            } catch (Exception e) {
                mainHandler.post(() -> showErrorDialog("Update check failed", e.getMessage()));
            }
        });
    }

    private void promptUpdate(String current, String latest, String apkUrl) {
        new androidx.appcompat.app.AlertDialog.Builder(this)
                .setTitle("Update available")
                .setMessage("Current: " + shortSha(current)
                        + "\nLatest:  " + shortSha(latest)
                        + "\n\nDownload and install now?"
                        + (dirty ? "\n\nWarning: unsaved marks will be preserved but should be saved first." : ""))
                .setPositiveButton("Update", (d, w) -> downloadAndInstall(apkUrl))
                .setNegativeButton(android.R.string.cancel, null)
                .show();
    }

    private void downloadAndInstall(String apkUrl) {
        setLoading(true);
        Toast.makeText(this, "Downloading update…", Toast.LENGTH_SHORT).show();
        executor.execute(() -> {
            try {
                File apkFile = new File(getCacheDir(), "update.apk");
                downloadTo(apkUrl, apkFile);
                mainHandler.post(() -> {
                    setLoading(false);
                    requestInstall(apkFile);
                });
            } catch (Exception e) {
                mainHandler.post(() -> {
                    setLoading(false);
                    showErrorDialog("Download failed", e.getMessage());
                });
            }
        });
    }

    private void downloadTo(String urlStr, File dest) throws Exception {
        URL url = new URL(urlStr);
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setInstanceFollowRedirects(true);
        conn.setRequestProperty("User-Agent", "apk-excelparse-app");
        int code = conn.getResponseCode();
        if (code < 200 || code >= 300) throw new RuntimeException("HTTP " + code);
        try (InputStream in = conn.getInputStream();
             OutputStream out = new FileOutputStream(dest)) {
            byte[] buf = new byte[64 * 1024];
            int n;
            while ((n = in.read(buf)) > 0) out.write(buf, 0, n);
        }
    }

    private void requestInstall(File apkFile) {
        if (!getPackageManager().canRequestPackageInstalls()) {
            new androidx.appcompat.app.AlertDialog.Builder(this)
                    .setTitle("Permission needed")
                    .setMessage("Grant this app permission to install updates, then tap Update again.")
                    .setPositiveButton("Open settings", (d, w) -> {
                        Intent i = new Intent(Settings.ACTION_MANAGE_UNKNOWN_APP_SOURCES,
                                Uri.parse("package:" + getPackageName()));
                        startActivity(i);
                    })
                    .setNegativeButton(android.R.string.cancel, null)
                    .show();
            return;
        }
        Uri apkUri = FileProvider.getUriForFile(this,
                getPackageName() + ".fileprovider", apkFile);
        Intent intent = new Intent(Intent.ACTION_VIEW);
        intent.setDataAndType(apkUri, "application/vnd.android.package-archive");
        intent.setFlags(Intent.FLAG_ACTIVITY_NEW_TASK | Intent.FLAG_GRANT_READ_URI_PERMISSION);
        startActivity(intent);
    }

    private static String shortSha(String sha) {
        if (sha == null || sha.isEmpty()) return "unknown";
        return sha.length() >= 7 ? sha.substring(0, 7) : sha;
    }
}
