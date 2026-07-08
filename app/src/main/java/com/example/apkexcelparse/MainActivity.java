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
import android.text.SpannableStringBuilder;
import android.text.Spanned;
import android.text.style.ForegroundColorSpan;
import android.view.Gravity;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.widget.CheckBox;
import android.widget.EditText;
import android.widget.LinearLayout;
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
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
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
    private MaterialButton prevStudentButton;
    private MaterialButton nextStudentButton;
    private MaterialButton prevLetterButton;
    private MaterialButton nextLetterButton;
    private MaterialButton prevGroupButton;
    private MaterialButton nextGroupButton;
    private MaterialButton overviewButton;
    private MaterialButton saveButton;
    private TextView dirtyIndicator;
    private TextView studentAverage;
    private DotProgressView dotProgressView;
    private TableLayout overviewTable;

    // Overlay screens (filter / completion / subpage) reachable from the burger menu.
    private View filterContainer;
    private View completionContainer;
    private View subpageContainer;
    private LinearLayout filterList;
    private LinearLayout completionContent;
    private LinearLayout subpageContent;
    private TextView subpageTitle;
    private TextView criterionNote;
    // Red flag (overview page only): tick + reason box shown under the student name.
    private CheckBox redFlagCheck;
    private TextView redFlagReason;
    private MaterialButton backToNotesButton;
    private View notesContainer;
    private TextView noteCounter;
    private TextView noteStudentName;
    private TextView noteGroupName;
    private TextView noteCriterionName;
    private TextView noteText;
    private TextView noteEmpty;
    private DotProgressView noteDot;

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
    // When a Complétude subpage opens a student's overview, this remembers which subpage to return
    // to so the GENERAL button reads RETOUR and goes back to that list. 0 = none, 1 = partial, 2 = unmarked.
    private int overviewReturnSubpage = 0;
    private static final int SUBPAGE_PARTIAL = 1;
    private static final int SUBPAGE_UNMARKED = 2;
    private static final int SUBPAGE_MARKED = 3;
    private static final int SUBPAGE_LISTE = 4;
    // The generic subpage list is shared between Complétude (back → completion) and the Liste
    // roster (back → grading). This flag routes the subpage BACK button to the right place.
    private boolean subpageIsListe;
    // Liste subpage: when true, the roster shows only red-flagged students (top tick box).
    private boolean listeRedFlagOnly;
    private boolean dirty;
    // True from pressing SAUVER until the write finishes; drives the "• sauvegarde…" indicator.
    private boolean saving;

    // Per-(student, criterion) notes live on a "notes" sheet inside the workbook (see XlsxParser),
    // at the same (row, column) as the mark cell. They are dirty-tracked and saved like marks.
    // Notes carousel: list of {studentIndex, criterionIndex} pairs that have a note, + cursor.
    private List<int[]> notesList = new ArrayList<>();
    private int notePos;
    // When a criterion page is reached via the Notes "ALLER À…" button, this is the notesList
    // position to return to (the criterion page then shows an extra "‹ NOTES" button). -1 = inactive.
    private int noteReturnPos = -1;

    // Per-group visibility filter. groupHidden[i] hides model.groups.get(i) everywhere in the app;
    // remembered per workbook (keyed by display name). Never touches the workbook's formulas.
    private boolean[] groupHidden;
    private static final String FILTER_PREFIX = "filter_hidden::";

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
        filterContainer = findViewById(R.id.filterContainer);
        completionContainer = findViewById(R.id.completionContainer);
        subpageContainer = findViewById(R.id.subpageContainer);
        filterList = findViewById(R.id.filterList);
        completionContent = findViewById(R.id.completionContent);
        subpageContent = findViewById(R.id.subpageContent);
        subpageTitle = findViewById(R.id.subpageTitle);
        criterionNote = findViewById(R.id.criterionNote);
        redFlagCheck = findViewById(R.id.redFlagCheck);
        redFlagReason = findViewById(R.id.redFlagReason);
        backToNotesButton = findViewById(R.id.backToNotesButton);
        notesContainer = findViewById(R.id.notesContainer);
        noteCounter = findViewById(R.id.noteCounter);
        noteStudentName = findViewById(R.id.noteStudentName);
        noteGroupName = findViewById(R.id.noteGroupName);
        noteCriterionName = findViewById(R.id.noteCriterionName);
        noteText = findViewById(R.id.noteText);
        noteEmpty = findViewById(R.id.noteEmpty);
        noteDot = findViewById(R.id.noteDot);

        findViewById(R.id.filterBackButton).setOnClickListener(v -> closeFilterScreen());
        findViewById(R.id.completionBackButton).setOnClickListener(v -> showGrading());
        findViewById(R.id.subpageBackButton).setOnClickListener(v -> onSubpageBack());
        findViewById(R.id.selectAllGroupsButton).setOnClickListener(v -> setAllGroupsHidden(false));
        findViewById(R.id.clearAllGroupsButton).setOnClickListener(v -> setAllGroupsHidden(true));
        criterionNote.setOnClickListener(v -> editCurrentCriterionNote());
        // setOnClickListener fires only on user taps (post-toggle state), so programmatic
        // setChecked() during render never re-triggers a write.
        redFlagCheck.setOnClickListener(v -> onRedFlagToggled(redFlagCheck.isChecked()));
        redFlagReason.setOnClickListener(v -> editRedFlagReason());
        backToNotesButton.setOnClickListener(v -> returnToNotes());
        findViewById(R.id.prevNoteButton).setOnClickListener(v -> navigateNote(-1));
        findViewById(R.id.nextNoteButton).setOnClickListener(v -> navigateNote(1));
        findViewById(R.id.notesBackButton).setOnClickListener(v -> showGrading());
        noteText.setOnClickListener(v -> editCarouselNote());
        // Tapping the dot jumps to its criterion page (replaces the old "ALLER À…" button).
        noteDot.setOnDotTapListener(idx -> gotoNoteCriterion());

        findViewById(R.id.pickButton).setOnClickListener(v -> launchPicker());
        findViewById(R.id.pickerCheckUpdatesButton).setOnClickListener(v -> checkForUpdates());
        findViewById(R.id.menuButton).setOnClickListener(this::showOverflowMenu);
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
            saving = true;
            updateDirtyUi();          // "• sauvegarde…" until the write finishes
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
        menu.getMenu().add(0, 1, 0, R.string.filter_groups);
        menu.getMenu().add(0, 2, 1, R.string.completion);
        menu.getMenu().add(0, 3, 2, R.string.notes);
        menu.getMenu().add(0, 6, 3, R.string.liste);
        menu.getMenu().add(0, 4, 4, R.string.change_file);
        menu.getMenu().add(0, 5, 5, R.string.check_updates);
        menu.setOnMenuItemClickListener(item -> {
            if (item.getItemId() == 1) { showFilterScreen(); return true; }
            if (item.getItemId() == 2) { showCompletionScreen(); return true; }
            if (item.getItemId() == 3) { showNotesScreen(); return true; }
            if (item.getItemId() == 6) { showListeSubpage(); return true; }
            if (item.getItemId() == 4) { launchPicker(); return true; }
            if (item.getItemId() == 5) { checkForUpdates(); return true; }
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
                    noteReturnPos = -1;
                    loadFilterState();
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

    private void navigateStudent(int delta) {
        if (model == null) return;
        int size = model.students.size();
        // Wrap around the roster; keeps the same page (overview stays overview) across students.
        int newIdx = ((studentIdx + delta) % size + size) % size;
        if (newIdx == studentIdx) return;
        studentIdx = newIdx;
        noteReturnPos = -1;
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
                noteReturnPos = -1;
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

    /**
     * Jump to the previous/next *visible* group. At the ends of the group sequence this rolls over
     * into the neighbouring student (roster wraps): ‹ groupe on the first group → previous student's
     * last group; groupe › on the last group → next student's first group.
     */
    private void navigateGroup(int delta) {
        if (model == null || model.groups.isEmpty()) return;
        List<Integer> visible = visibleGroupIndices();
        if (visible.isEmpty()) return;
        // Landing on a criterion page abandons any stashed OVERVIEW / subpage / notes return.
        overviewReturnPage = -1;
        overviewReturnSubpage = 0;
        noteReturnPos = -1;
        int curGroupIdx = isOverviewPage() ? -1 : indexOfGroupForCriterion(pageIdx - 1);
        int pos = visible.indexOf(curGroupIdx);
        if (pos < 0) {
            // From the overview: first (›) or last (‹) visible group of the same student.
            int gi = delta >= 0 ? visible.get(0) : visible.get(visible.size() - 1);
            pageIdx = model.groups.get(gi).firstCriterionIndex + 1;
            render();
            return;
        }
        int target = pos + delta;
        if (target < 0) {
            goToNeighbourStudentGroup(-1, true);        // previous student, its last group
        } else if (target >= visible.size()) {
            goToNeighbourStudentGroup(1, false);        // next student, its first group
        } else {
            pageIdx = model.groups.get(visible.get(target)).firstCriterionIndex + 1;
            render();
        }
    }

    /** Move to an adjacent student (roster wraps) and land on their first/last visible group. */
    private void goToNeighbourStudentGroup(int studentDelta, boolean landOnLastGroup) {
        int size = model.students.size();
        studentIdx = ((studentIdx + studentDelta) % size + size) % size;
        List<Integer> visible = visibleGroupIndices();
        if (visible.isEmpty()) {
            pageIdx = 0;
        } else {
            int gi = landOnLastGroup ? visible.get(visible.size() - 1) : visible.get(0);
            pageIdx = model.groups.get(gi).firstCriterionIndex + 1;
        }
        render();
    }

    /** Jump straight to a group's first criterion page (used by the overview table name links). */
    private void jumpToGroup(int firstCriterionIndex) {
        if (model == null) return;
        overviewReturnPage = -1; // deliberate jump to a criterion; abandon any GENERAL stash
        overviewReturnSubpage = 0;
        noteReturnPos = -1;
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
        if (overviewReturnSubpage > 0) {
            // Reached this overview from a Complétude subpage: RETOUR goes back to that list.
            int kind = overviewReturnSubpage;
            overviewReturnSubpage = 0;
            if (kind == SUBPAGE_MARKED) showMarkedSubpage();
            else if (kind == SUBPAGE_PARTIAL) showPartialSubpage();
            else if (kind == SUBPAGE_LISTE) showListeSubpage();
            else showUnmarkedSubpage();
        } else if (overviewReturnPage >= 0) {
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
        boolean back = overviewReturnPage >= 0 || overviewReturnSubpage > 0;
        overviewButton.setText(back ? R.string.overview_back : R.string.overview_button);
    }

    private boolean isOverviewPage() {
        return pageIdx == 0;
    }

    private void render() {
        if (model == null) return;
        Student s = model.students.get(studentIdx);

        studentCounter.setText(String.format(Locale.getDefault(),
                "étudiant %d / %d", studentIdx + 1, model.students.size()));
        studentName.setText(nameWithFlag(studentIdx));

        if (isOverviewPage()) {
            renderOverview(s);
        } else {
            renderCriterion(model.criteria.get(pageIdx - 1), s);
        }

        updateOverviewButton();
        // The extra "‹ NOTES" button shows only on a criterion page reached via ALLER À.
        backToNotesButton.setVisibility(
                (!isOverviewPage() && noteReturnPos >= 0) ? View.VISIBLE : View.GONE);
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
        criterionNote.setVisibility(View.GONE);
        markButtonsContainer.setVisibility(View.GONE);

        // Red flag tick + (when on) reason box, under the student name — GENERAL page only.
        boolean flagged = XlsxParser.readRedFlag(workbook, s);
        redFlagCheck.setVisibility(View.VISIBLE);
        redFlagCheck.setChecked(flagged);
        redFlagReason.setText(XlsxParser.readRedFlagReason(workbook, s));
        redFlagReason.setVisibility(flagged ? View.VISIBLE : View.GONE);

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
        List<Integer> visibleGroups = visibleGroupIndices();
        int maxDots = 1;
        for (int gi : visibleGroups) {
            Group g = model.groups.get(gi);
            maxDots = Math.max(maxDots, g.lastCriterionIndex - g.firstCriterionIndex + 1);
        }
        int dotsCellW = Math.round(slotPx * maxDots);
        for (int gi : visibleGroups) {
            Group g = model.groups.get(gi);
            TableRow row = new TableRow(this);

            int n = g.lastCriterionIndex - g.firstCriterionIndex + 1;
            int[] buckets = new int[n];
            float[] scales = new float[n];
            boolean hasUnmarked = false;
            for (int i = 0; i < n; i++) {
                Criterion c = model.criteria.get(g.firstCriterionIndex + i);
                buckets[i] = markToBucket(XlsxParser.readMark(workbook, s, c));
                scales[i] = coefScale(c.coefficient);
                if (buckets[i] < 0) hasUnmarked = true;
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
            TextView avgCell = overviewCell(avgText, true);
            // A group with an unmarked criterion has an incomplete (not yet meaningful) average, so
            // it is shown light gray; otherwise the usual red-below-4 / blue colouring applies.
            avgCell.setTextColor(hasUnmarked ? AVG_INCOMPLETE_COLOR : avgColor(avg));
            row.addView(avgCell);

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

    /** Slot width (px) of one dot in the full-width header string of the visible criteria. */
    private float headerSlotPx() {
        float contentPx = getResources().getDisplayMetrics().widthPixels - dpToPx(48f);
        int n = Math.max(1, visibleCriterionIndices().size());
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
        criterionContract.setVisibility(View.VISIBLE);
        criterionRemarks.setVisibility(View.VISIBLE);
        markButtonsContainer.setVisibility(View.VISIBLE);
        redFlagCheck.setVisibility(View.GONE);
        redFlagReason.setVisibility(View.GONE);

        criterionGroup.setText(c.groupName != null ? c.groupName : "");
        criterionContract.setText(c.contract != null ? c.contract : "");
        criterionRemarks.setText(c.remarks != null ? c.remarks : "");

        // Under the group name: the group's weight ("coeff. N"), same size/weight as the name but gray.
        Group g = model.groupForCriterion(pageIdx - 1);
        String groupCoef = g != null ? coefDigit(g.coefficient) : null;
        if (groupCoef != null) {
            criterionIdCoef.setText("coeff. " + groupCoef);
            criterionIdCoef.setTextSize(16f);
            criterionIdCoef.setTypeface(android.graphics.Typeface.DEFAULT_BOLD);
            criterionIdCoef.setTextColor(0xFF888888);
            criterionIdCoef.setVisibility(View.VISIBLE);
        } else {
            criterionIdCoef.setVisibility(View.GONE);
        }

        // Per-(student, criterion) note box, right under the description. Tap to edit.
        criterionNote.setVisibility(View.VISIBLE);
        criterionNote.setText(getNote(studentIdx, pageIdx - 1));

        Double existing = XlsxParser.readMark(workbook, s, c);
        // No stored mark → nothing selected (all buttons shaded).
        refreshMarkButtons(existing != null ? markToBucket(existing) : -1);
    }

    /** A mark button was tapped: write its value and reflect the new selection (no auto-advance). */
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
        boolean overview = isOverviewPage();
        // Which criteria to show: all visible criteria on the overview, or the current group's
        // criteria on a criterion page (the group is always visible when a criterion page is shown).
        List<Integer> idxs;
        if (overview || model.criteria.isEmpty()) {
            idxs = visibleCriterionIndices();
        } else {
            idxs = new ArrayList<>();
            Group g = model.groupForCriterion(pageIdx - 1);
            int first = g != null ? g.firstCriterionIndex : pageIdx - 1;
            int last = g != null ? g.lastCriterionIndex : pageIdx - 1;
            for (int i = first; i <= last; i++) idxs.add(i);
        }
        int n = idxs.size();
        int[] buckets = new int[n];
        float[] scales = new float[n];
        String[] labels = new String[n];
        String[] critIds = new String[n];
        for (int i = 0; i < n; i++) {
            Criterion c = model.criteria.get(idxs.get(i));
            buckets[i] = markToBucket(XlsxParser.readMark(workbook, s, c));
            scales[i] = coefScale(c.coefficient, 0.33f);
            labels[i] = coefDigit(c.coefficient);
            critIds[i] = c.id;
        }
        // Criterion pages need a taller row: headroom for heavy-coef dots plus the number band under
        // them; the overview's tiny all-criteria dots (no number line) stay 33dp.
        int rowDp = overview ? 33 : 60;
        android.view.ViewGroup.LayoutParams lp = dotProgressView.getLayoutParams();
        int rowPx = Math.round(dpToPx(rowDp));
        if (lp.height != rowPx) {
            lp.height = rowPx;
            dotProgressView.setLayoutParams(lp);
        }
        dotProgressView.setValues(buckets);
        dotProgressView.setScales(scales);
        // Coefficient digit inside each dot + criterion id under each dot, criterion pages only.
        dotProgressView.setLabels(overview ? null : labels);
        dotProgressView.setBelowLabels(overview ? null : critIds);
        // Highlight the current criterion's dot on criterion pages; none on the overview.
        dotProgressView.setHighlightIndex(overview ? -1 : idxs.indexOf(pageIdx - 1));
        // On criterion pages the dots are tappable to jump to that criterion within the group;
        // the overview's all-criteria dots stay non-interactive.
        final List<Integer> tapTargets = idxs;
        dotProgressView.setOnDotTapListener(
                overview ? null : idx -> goToCriterion(tapTargets.get(idx)));
    }

    /** Jump to a specific criterion of the current student (used by tapping a group dot). */
    private void goToCriterion(int criterionIndex) {
        if (model == null || criterionIndex < 0 || criterionIndex >= model.criteria.size()) return;
        overviewReturnPage = -1; // a within-student criterion move, like the old « critère » buttons
        overviewReturnSubpage = 0;
        noteReturnPos = -1;
        pageIdx = criterionIndex + 1;
        render();
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
        Double avg;
        if (isOverviewPage()) {
            // Overview: Excel's real CL when nothing is filtered, else our weighted mean of the
            // visible groups (computed on the fly; the workbook's formulas are never changed).
            avg = anyGroupHidden()
                    ? computeFilteredGeneralAverage(s)
                    : XlsxParser.readNumericAt(workbook, s, XlsxParser.COL_STUDENT_AVERAGE, formulaEvaluator);
        } else {
            Group g = model.groupForCriterion(pageIdx - 1);
            int col = g != null ? g.averageColumnIndex : XlsxParser.COL_STUDENT_AVERAGE;
            avg = XlsxParser.readNumericAt(workbook, s, col, formulaEvaluator);
        }
        if (avg == null || avg.isNaN()) {
            studentAverage.setText("");
        } else {
            studentAverage.setText(String.format(Locale.getDefault(), "%.2f", avg));
        }
        studentAverage.setTextColor(avgColor(avg));
    }

    // Averages are /6; a failing grade (< 4) is shown red, otherwise the normal blue.
    private static final int AVG_LOW_COLOR = 0xFFD32F2F;
    private static final int AVG_OK_COLOR = 0xFF1976D2;
    // Overview group average shown light gray when the group still has an unmarked criterion.
    private static final int AVG_INCOMPLETE_COLOR = 0xFFCCCCCC;
    private static final double AVG_LOW_THRESHOLD = 4.0;

    /** Red when the /6 average is below the pass threshold, blue (or absent) otherwise. */
    private static int avgColor(Double avg) {
        return (avg != null && !avg.isNaN() && avg < AVG_LOW_THRESHOLD) ? AVG_LOW_COLOR : AVG_OK_COLOR;
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
        if (dirtyIndicator == null) return;
        if (saving) {
            dirtyIndicator.setText("•  sauvegarde…");
            dirtyIndicator.setVisibility(View.VISIBLE);
        } else if (dirty) {
            dirtyIndicator.setText("•  modifié");
            dirtyIndicator.setVisibility(View.VISIBLE);
        } else {
            dirtyIndicator.setVisibility(View.GONE);
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
                saving = false;
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
                saving = false;
                updateDirtyUi();      // save failed → back to "• modifié"
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

    /** Show exactly one top-level screen; hide the picker, grading, and all overlays. */
    private void setScreen(View screen) {
        pickerContainer.setVisibility(screen == pickerContainer ? View.VISIBLE : View.GONE);
        gradingContainer.setVisibility(screen == gradingContainer ? View.VISIBLE : View.GONE);
        filterContainer.setVisibility(screen == filterContainer ? View.VISIBLE : View.GONE);
        completionContainer.setVisibility(screen == completionContainer ? View.VISIBLE : View.GONE);
        subpageContainer.setVisibility(screen == subpageContainer ? View.VISIBLE : View.GONE);
        notesContainer.setVisibility(screen == notesContainer ? View.VISIBLE : View.GONE);
    }

    private void showPicker() {
        setScreen(pickerContainer);
    }

    private void showGrading() {
        setScreen(gradingContainer);
    }

    // ---- Group filter -------------------------------------------------------------------------

    private boolean isGroupHiddenAt(int groupIndex) {
        return groupHidden != null && groupIndex >= 0 && groupIndex < groupHidden.length
                && groupHidden[groupIndex];
    }

    private boolean anyGroupHidden() {
        if (groupHidden == null) return false;
        for (boolean b : groupHidden) if (b) return true;
        return false;
    }

    private boolean allGroupsHidden() {
        if (groupHidden == null || groupHidden.length == 0) return false;
        for (boolean b : groupHidden) if (!b) return false;
        return true;
    }

    /** Indices into model.groups of the currently visible (ticked) groups, in order. */
    private List<Integer> visibleGroupIndices() {
        List<Integer> r = new ArrayList<>();
        if (model == null) return r;
        for (int i = 0; i < model.groups.size(); i++) if (!isGroupHiddenAt(i)) r.add(i);
        return r;
    }

    private boolean isCriterionVisible(int criterionIndex) {
        return !isGroupHiddenAt(indexOfGroupForCriterion(criterionIndex));
    }

    /** Indices into model.criteria of criteria belonging to visible groups, in order. */
    private List<Integer> visibleCriterionIndices() {
        List<Integer> r = new ArrayList<>();
        if (model == null) return r;
        for (int i = 0; i < model.criteria.size(); i++) if (isCriterionVisible(i)) r.add(i);
        return r;
    }

    private String filterPrefsKey() {
        String d = getSharedPreferences(PREFS, MODE_PRIVATE).getString(KEY_DISPLAY_NAME, "");
        return FILTER_PREFIX + (d == null ? "" : d);
    }

    /** Load the per-workbook filter (hidden group names) into groupHidden. Defaults to all visible. */
    private void loadFilterState() {
        groupHidden = new boolean[model.groups.size()];
        Set<String> hidden = getSharedPreferences(PREFS, MODE_PRIVATE)
                .getStringSet(filterPrefsKey(), null);
        if (hidden == null || hidden.isEmpty()) return;
        for (int i = 0; i < model.groups.size(); i++) {
            String name = model.groups.get(i).name;
            if (name != null && hidden.contains(name)) groupHidden[i] = true;
        }
        if (allGroupsHidden()) java.util.Arrays.fill(groupHidden, false); // never hide everything
    }

    private void saveFilterState() {
        if (groupHidden == null) return;
        Set<String> hidden = new HashSet<>();
        for (int i = 0; i < groupHidden.length; i++) {
            if (groupHidden[i]) {
                String n = model.groups.get(i).name;
                if (n != null) hidden.add(n);
            }
        }
        getSharedPreferences(PREFS, MODE_PRIVATE).edit()
                .putStringSet(filterPrefsKey(), hidden).apply();
    }

    private void showFilterScreen() {
        if (model == null) return;
        buildFilterList();
        setScreen(filterContainer);
    }

    /** One checkbox per group; ticked = visible. All groups may be unticked (transient empty state). */
    private void buildFilterList() {
        filterList.removeAllViews();
        int pad = Math.round(dpToPx(8f));
        for (int i = 0; i < model.groups.size(); i++) {
            final int gi = i;
            Group g = model.groups.get(i);
            CheckBox cb = new CheckBox(this);
            cb.setText(g.name != null ? g.name : "(groupe)");
            cb.setTextSize(16f);
            cb.setPadding(pad, pad, pad, pad);
            cb.setChecked(!isGroupHiddenAt(gi));
            cb.setOnCheckedChangeListener((btn, checked) -> {
                groupHidden[gi] = !checked;
                saveFilterState();
            });
            filterList.addView(cb);
        }
    }

    /** "SÉLECTIONNER TOUT" / "EFFACER SÉLECTION": tick or untick every group at once. */
    private void setAllGroupsHidden(boolean hidden) {
        if (groupHidden == null) return;
        java.util.Arrays.fill(groupHidden, hidden);
        saveFilterState();
        buildFilterList(); // refresh checkbox states
    }

    private void closeFilterScreen() {
        // If the criterion page we were on now sits in a hidden group, fall back to the overview.
        if (model != null && !isOverviewPage()
                && isGroupHiddenAt(indexOfGroupForCriterion(pageIdx - 1))) {
            pageIdx = 0;
            overviewReturnPage = -1;
        }
        showGrading();
        render();
    }

    /**
     * General average for the overview when a filter is active: the visible groups' /6 grades
     * weighted by their group coefficients, quarter-rounded — the same math as Excel's CK→CL,
     * restricted to ticked groups. Computed for display only; the workbook is never changed.
     */
    private Double computeFilteredGeneralAverage(Student s) {
        double wsum = 0, w = 0;
        for (int gi : visibleGroupIndices()) {
            Group g = model.groups.get(gi);
            Double grade = XlsxParser.readNumericAt(workbook, s, g.averageColumnIndex, formulaEvaluator);
            Double coef = g.coefficient;
            if (grade == null || grade.isNaN() || coef == null || coef <= 0) continue;
            wsum += grade * coef;
            w += coef;
        }
        if (w <= 0) return null;
        double ck = Math.round((wsum / w) * 100.0) / 100.0;                 // CK: round to 2 decimals
        return Math.round(Math.floor(ck * 4 + 0.5) / 4 * 100.0) / 100.0;    // CL: round to nearest 0.25
    }

    // ---- Completion page ----------------------------------------------------------------------

    /** Counts of marked/partial/unmarked criteria, groups and students over the visible groups. */
    private static class Completion {
        int markedCriteria, totalCriteria;
        int markedGroups, totalGroups;
        int partialGroups, unmarkedGroups;
        int totalStudents;
        final List<Integer> markedStudents = new ArrayList<>();    // fully-marked, indices into students
        final List<Integer> partialStudents = new ArrayList<>();  // indices into model.students
        final List<Double> partialPercents = new ArrayList<>();   // parallel to partialStudents, 0..100
        final List<Integer> unmarkedStudents = new ArrayList<>();
    }

    private Completion computeCompletion() {
        Completion cc = new Completion();
        List<Integer> vgroups = visibleGroupIndices();
        int nStudents = model.students.size();
        int nVisibleCriteria = visibleCriterionIndices().size();
        cc.totalStudents = nStudents;
        cc.totalCriteria = nStudents * nVisibleCriteria;
        cc.totalGroups = nStudents * vgroups.size();
        for (int si = 0; si < nStudents; si++) {
            Student s = model.students.get(si);
            int studentMarked = 0;
            for (int gi : vgroups) {
                Group g = model.groups.get(gi);
                int gSize = g.lastCriterionIndex - g.firstCriterionIndex + 1;
                int gMarked = 0;
                for (int ci = g.firstCriterionIndex; ci <= g.lastCriterionIndex; ci++) {
                    if (XlsxParser.readMark(workbook, s, model.criteria.get(ci)) != null) gMarked++;
                }
                if (gMarked == gSize) cc.markedGroups++;
                else if (gMarked > 0) cc.partialGroups++;
                else cc.unmarkedGroups++;
                studentMarked += gMarked;
            }
            cc.markedCriteria += studentMarked;
            if (nVisibleCriteria > 0 && studentMarked == nVisibleCriteria) {
                cc.markedStudents.add(si);
            } else if (studentMarked > 0) {
                cc.partialStudents.add(si);
                cc.partialPercents.add(100.0 * studentMarked / nVisibleCriteria);
            } else {
                cc.unmarkedStudents.add(si);
            }
        }
        return cc;
    }

    private void showCompletionScreen() {
        if (model == null) return;
        buildCompletion();
        setScreen(completionContainer);
    }

    private void buildCompletion() {
        completionContent.removeAllViews();
        Completion cc = computeCompletion();
        completionContent.addView(sectionHeader("totaux"));
        completionContent.addView(statRow("étudiants notés",
                cc.markedStudents.size(), cc.totalStudents, true, this::showMarkedSubpage));
        completionContent.addView(statRow("critères notés", cc.markedCriteria, cc.totalCriteria, false, null));
        completionContent.addView(statRow("groupes notés", cc.markedGroups, cc.totalGroups, false, null));
        completionContent.addView(sectionHeader("notation partielle"));
        completionContent.addView(statRow("étudiants partiellement notés",
                cc.partialStudents.size(), cc.totalStudents, true, this::showPartialSubpage));
        completionContent.addView(statRow("groupes partiellement notés", cc.partialGroups, cc.totalGroups, false, null));
        completionContent.addView(sectionHeader("non noté"));
        completionContent.addView(statRow("étudiants non notés",
                cc.unmarkedStudents.size(), cc.totalStudents, true, this::showUnmarkedSubpage));
        completionContent.addView(statRow("groupes non notés", cc.unmarkedGroups, cc.totalGroups, false, null));
    }

    private void showMarkedSubpage() {
        subpageIsListe = false;
        Completion cc = computeCompletion();
        subpageTitle.setText("étudiants notés");
        subpageContent.removeAllViews();
        if (cc.markedStudents.isEmpty()) subpageContent.addView(emptyNote("(aucun)"));
        for (int si : cc.markedStudents) {
            subpageContent.addView(studentLinkRow(si, null, SUBPAGE_MARKED));
        }
        setScreen(subpageContainer);
    }

    private void showPartialSubpage() {
        subpageIsListe = false;
        Completion cc = computeCompletion();
        subpageTitle.setText("étudiants partiellement notés");
        subpageContent.removeAllViews();
        if (cc.partialStudents.isEmpty()) subpageContent.addView(emptyNote("(aucun)"));
        for (int k = 0; k < cc.partialStudents.size(); k++) {
            int si = cc.partialStudents.get(k);
            String suffix = "   —   "
                    + String.format(Locale.getDefault(), "%.1f", cc.partialPercents.get(k)) + "%";
            subpageContent.addView(studentLinkRow(si, suffix, SUBPAGE_PARTIAL));
        }
        setScreen(subpageContainer);
    }

    private void showUnmarkedSubpage() {
        subpageIsListe = false;
        Completion cc = computeCompletion();
        subpageTitle.setText("étudiants non notés");
        subpageContent.removeAllViews();
        if (cc.unmarkedStudents.isEmpty()) subpageContent.addView(emptyNote("(aucun)"));
        for (int si : cc.unmarkedStudents) {
            subpageContent.addView(studentLinkRow(si, null, SUBPAGE_UNMARKED));
        }
        setScreen(subpageContainer);
    }

    /**
     * Full roster (burger → "Liste"): every student, each a link to their GENERAL page. Reuses the
     * generic subpage list; {@code subpageIsListe} routes its BACK button back to the grading screen.
     */
    private void showListeSubpage() {
        if (model == null) return;
        subpageIsListe = true;
        subpageTitle.setText(R.string.liste_title);
        subpageContent.removeAllViews();

        // Top tick box: keep only red-flagged students. Rebuilds the list on toggle.
        CheckBox flagFilter = new CheckBox(this);
        flagFilter.setText(R.string.liste_flag_filter);
        flagFilter.setTextSize(15f);
        flagFilter.setTextColor(RED_FLAG_COLOR);
        flagFilter.setButtonTintList(ColorStateList.valueOf(RED_FLAG_COLOR));
        flagFilter.setChecked(listeRedFlagOnly);
        int pad = Math.round(dpToPx(4f));
        flagFilter.setPadding(0, pad, 0, pad);
        flagFilter.setOnClickListener(v -> { listeRedFlagOnly = flagFilter.isChecked(); showListeSubpage(); });
        subpageContent.addView(flagFilter);

        int shown = 0;
        for (int si = 0; si < model.students.size(); si++) {
            if (listeRedFlagOnly && !XlsxParser.readRedFlag(workbook, model.students.get(si))) continue;
            subpageContent.addView(studentLinkRow(si, null, SUBPAGE_LISTE));
            shown++;
        }
        if (shown == 0) subpageContent.addView(emptyNote("(aucun)"));

        setScreen(subpageContainer);
    }

    /** The generic subpage BACK button: to the roster's grading screen, or back to Complétude. */
    private void onSubpageBack() {
        if (subpageIsListe) showGrading();
        else showCompletionScreen();
    }

    /**
     * Jump to a student's GENERAL (overview) page from a completion subpage link. Remembers the
     * subpage so the GENERAL button reads RETOUR and returns to that list (like arriving from a
     * criterion page).
     */
    private void openStudentOverview(int studentIndex, int subpageKind) {
        if (model == null || studentIndex < 0 || studentIndex >= model.students.size()) return;
        studentIdx = studentIndex;
        pageIdx = 0;
        overviewReturnPage = -1;
        overviewReturnSubpage = subpageKind;
        showGrading();
        render();
    }

    private TextView sectionHeader(String text) {
        TextView tv = new TextView(this);
        tv.setText(text);
        tv.setTextSize(13f);
        tv.setAllCaps(true);
        tv.setTypeface(android.graphics.Typeface.DEFAULT_BOLD);
        tv.setTextColor(0xFF888888);
        tv.setPadding(0, Math.round(dpToPx(16f)), 0, Math.round(dpToPx(4f)));
        return tv;
    }

    private View statRow(String label, int count, int total, boolean asLink, Runnable onClick) {
        LinearLayout row = new LinearLayout(this);
        row.setOrientation(LinearLayout.HORIZONTAL);
        row.setGravity(Gravity.CENTER_VERTICAL);
        int vpad = Math.round(dpToPx(8f));
        row.setPadding(0, vpad, 0, vpad);

        TextView l = new TextView(this);
        l.setText(label);
        l.setTextSize(16f);
        l.setLayoutParams(new LinearLayout.LayoutParams(0,
                LinearLayout.LayoutParams.WRAP_CONTENT, 1f));
        if (asLink) {
            l.setTextColor(0xFF1976D2);
            l.setPaintFlags(l.getPaintFlags() | android.graphics.Paint.UNDERLINE_TEXT_FLAG);
            l.setClickable(true);
            if (onClick != null) l.setOnClickListener(v -> onClick.run());
        }
        row.addView(l);

        double pct = total > 0 ? 100.0 * count / total : 0.0;
        TextView val = new TextView(this);
        val.setText(count + "/" + total + "   "
                + String.format(Locale.getDefault(), "%.1f", pct) + "%");
        val.setTextSize(16f);
        val.setTypeface(android.graphics.Typeface.DEFAULT_BOLD);
        val.setTextColor(0xFF333333);
        row.addView(val);
        return row;
    }

    private TextView studentLinkRow(int studentIndex, String suffix, int subpageKind) {
        TextView tv = new TextView(this);
        CharSequence base = nameWithFlag(studentIndex); // trailing red dot when flagged
        tv.setText(suffix == null ? base : new SpannableStringBuilder(base).append(suffix));
        tv.setTextSize(16f);
        tv.setTextColor(0xFF1976D2);
        tv.setPaintFlags(tv.getPaintFlags() | android.graphics.Paint.UNDERLINE_TEXT_FLAG);
        int pad = Math.round(dpToPx(10f));
        tv.setPadding(0, pad, 0, pad);
        tv.setClickable(true);
        tv.setOnClickListener(v -> openStudentOverview(studentIndex, subpageKind));
        return tv;
    }

    private TextView emptyNote(String text) {
        TextView tv = new TextView(this);
        tv.setText(text);
        tv.setTextSize(15f);
        tv.setTextColor(0xFF888888);
        int pad = Math.round(dpToPx(10f));
        tv.setPadding(0, pad, 0, pad);
        return tv;
    }

    // ---- Red flag -----------------------------------------------------------------------------
    // A student's red flag lives on the notes sheet at the student-name column (see XlsxParser).
    // It is dirty-tracked and saved on SAUVER like marks and notes.

    private static final int RED_FLAG_COLOR = 0xFFD32F2F;

    /** Student name as shown everywhere: a trailing red dot appended when the student is flagged. */
    private CharSequence nameWithFlag(int studentIndex) {
        Student s = model.students.get(studentIndex);
        String name = s.name != null ? s.name : "";
        if (workbook == null || !XlsxParser.readRedFlag(workbook, s)) return name;
        SpannableStringBuilder sb = new SpannableStringBuilder(name);
        sb.append("  ");
        int start = sb.length();
        sb.append("●"); // ● red dot
        sb.setSpan(new ForegroundColorSpan(RED_FLAG_COLOR), start, sb.length(),
                Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);
        return sb;
    }

    /** Overview red-flag tick toggled: persist, then re-render (updates the name dot + reason box). */
    private void onRedFlagToggled(boolean on) {
        if (model == null || workbook == null || !isOverviewPage()) return;
        Student s = model.students.get(studentIdx);
        // Keep any existing reason when toggling on; toggling off blanks the cell.
        XlsxParser.writeRedFlag(workbook, s, on, XlsxParser.readRedFlagReason(workbook, s));
        dirty = true;
        updateDirtyUi();
        render();
    }

    /** Edit dialog for a student's red-flag reason. Saving a reason implies the flag is on. */
    private void editRedFlagReason() {
        if (model == null || workbook == null || !isOverviewPage()) return;
        Student s = model.students.get(studentIdx);
        final EditText input = new EditText(this);
        input.setInputType(android.text.InputType.TYPE_CLASS_TEXT
                | android.text.InputType.TYPE_TEXT_FLAG_MULTI_LINE
                | android.text.InputType.TYPE_TEXT_FLAG_CAP_SENTENCES);
        input.setMinLines(3);
        input.setGravity(Gravity.TOP);
        input.setText(XlsxParser.readRedFlagReason(workbook, s));
        input.setSelection(input.getText().length());
        int pad = Math.round(dpToPx(16f));
        android.widget.FrameLayout wrap = new android.widget.FrameLayout(this);
        wrap.setPadding(pad, pad / 2, pad, 0);
        wrap.addView(input);
        androidx.appcompat.app.AlertDialog dialog = new androidx.appcompat.app.AlertDialog.Builder(this)
                .setTitle(s.name + " · signalement")
                .setView(wrap)
                .setPositiveButton(android.R.string.ok, (d, w) -> {
                    XlsxParser.writeRedFlag(workbook, s, true, input.getText().toString());
                    dirty = true;
                    updateDirtyUi();
                    render();
                })
                .setNegativeButton(android.R.string.cancel, null)
                .create();
        input.requestFocus();
        if (dialog.getWindow() != null) {
            dialog.getWindow().setSoftInputMode(
                    android.view.WindowManager.LayoutParams.SOFT_INPUT_STATE_ALWAYS_VISIBLE);
        }
        dialog.show();
    }

    // ---- Notes --------------------------------------------------------------------------------
    // Notes live on a "notes" sheet inside the workbook, at the same (row, column) as each mark cell
    // (see XlsxParser.readNote/writeNote). They travel with the .xlsx, appear in the -graded export,
    // and are dirty-tracked + saved on SAUVER exactly like marks -- no separate local store.

    private String getNote(int studentIndex, int criterionIndex) {
        if (model == null || workbook == null) return "";
        return XlsxParser.readNote(workbook,
                model.students.get(studentIndex), model.criteria.get(criterionIndex));
    }

    private void setNote(int studentIndex, int criterionIndex, String text) {
        if (model == null || workbook == null) return;
        XlsxParser.writeNote(workbook,
                model.students.get(studentIndex), model.criteria.get(criterionIndex), text);
        dirty = true;
        updateDirtyUi();
    }

    /** Edit dialog for a (student, criterion) note; onSaved runs after the note is stored. */
    private void editNote(int studentIndex, int criterionIndex, Runnable onSaved) {
        if (model == null) return;
        final EditText input = new EditText(this);
        input.setInputType(android.text.InputType.TYPE_CLASS_TEXT
                | android.text.InputType.TYPE_TEXT_FLAG_MULTI_LINE
                | android.text.InputType.TYPE_TEXT_FLAG_CAP_SENTENCES);
        input.setMinLines(4);
        input.setGravity(Gravity.TOP);
        input.setText(getNote(studentIndex, criterionIndex));
        input.setSelection(input.getText().length());
        int pad = Math.round(dpToPx(16f));
        android.widget.FrameLayout wrap = new android.widget.FrameLayout(this);
        wrap.setPadding(pad, pad / 2, pad, 0);
        wrap.addView(input);
        Criterion c = model.criteria.get(criterionIndex);
        String title = model.students.get(studentIndex).name
                + " · critère " + (c.id != null ? c.id : "");
        androidx.appcompat.app.AlertDialog dialog = new androidx.appcompat.app.AlertDialog.Builder(this)
                .setTitle(title)
                .setView(wrap)
                .setPositiveButton(android.R.string.ok, (d, w) -> {
                    setNote(studentIndex, criterionIndex, input.getText().toString());
                    if (onSaved != null) onSaved.run();
                })
                .setNegativeButton(android.R.string.cancel, null)
                .create();
        // Focus the field and raise the keyboard right away, so no extra tap is needed to type.
        input.requestFocus();
        if (dialog.getWindow() != null) {
            dialog.getWindow().setSoftInputMode(
                    android.view.WindowManager.LayoutParams.SOFT_INPUT_STATE_ALWAYS_VISIBLE);
        }
        dialog.show();
    }

    private void editCurrentCriterionNote() {
        if (model == null || isOverviewPage()) return;
        int ci = pageIdx - 1;
        editNote(studentIdx, ci, () -> criterionNote.setText(getNote(studentIdx, ci)));
    }

    // ---- Notes carousel screen ----------------------------------------------------------------

    /** Collect every (student, criterion) pair that has a note, in student-then-criterion order. */
    private void buildNotesList() {
        notesList = new ArrayList<>();
        if (model == null || workbook == null) return;
        for (int si = 0; si < model.students.size(); si++) {
            Student s = model.students.get(si);
            for (int ci = 0; ci < model.criteria.size(); ci++) {
                if (!XlsxParser.readNote(workbook, s, model.criteria.get(ci)).isEmpty()) {
                    notesList.add(new int[]{si, ci});
                }
            }
        }
    }

    private void showNotesScreen() {
        if (model == null) return;
        buildNotesList();
        if (notePos < 0 || notePos >= notesList.size()) notePos = 0;
        renderNote();
        setScreen(notesContainer);
    }

    private void renderNote() {
        boolean has = !notesList.isEmpty();
        noteEmpty.setVisibility(has ? View.GONE : View.VISIBLE);
        noteStudentName.setVisibility(has ? View.VISIBLE : View.GONE);
        noteGroupName.setVisibility(has ? View.VISIBLE : View.GONE);
        noteCriterionName.setVisibility(has ? View.VISIBLE : View.GONE);
        noteText.setVisibility(has ? View.VISIBLE : View.GONE);
        noteDot.setVisibility(has ? View.VISIBLE : View.GONE);
        findViewById(R.id.prevNoteButton).setEnabled(notesList.size() > 1);
        findViewById(R.id.nextNoteButton).setEnabled(notesList.size() > 1);
        if (!has) {
            noteCounter.setText("");
            return;
        }
        if (notePos >= notesList.size()) notePos = notesList.size() - 1;
        int[] pair = notesList.get(notePos);
        Student s = model.students.get(pair[0]);
        Criterion c = model.criteria.get(pair[1]);
        noteCounter.setText(String.format(Locale.getDefault(),
                "note %d / %d", notePos + 1, notesList.size()));
        noteStudentName.setText(nameWithFlag(pair[0]));
        noteGroupName.setText(c.groupName != null ? c.groupName : "");
        noteCriterionName.setText("critère " + (c.id != null ? c.id : "")
                + (c.contract != null && !c.contract.isEmpty() ? " · " + c.contract : ""));
        noteText.setText(getNote(pair[0], pair[1]));
        // The criterion's dot, enlarged to 2× the criterion-page size (dot + inner coefficient digit),
        // over thin gray rings at every possible weight so its size reads as a comparison.
        noteDot.setValues(new int[]{markToBucket(XlsxParser.readMark(workbook, s, c))});
        noteDot.setScales(new float[]{coefScale(c.coefficient, 0.33f)});
        noteDot.setLabels(new String[]{coefDigit(c.coefficient)});
        noteDot.setBelowLabels(null);
        noteDot.setHighlightIndex(-1);
        noteDot.setReferenceScales(weightReferenceScales());
        noteDot.setSizeMultiplier(2f);
    }

    /** Radius multipliers for every distinct criterion coefficient present, ascending (weight rings). */
    private float[] weightReferenceScales() {
        java.util.TreeSet<Double> coefs = new java.util.TreeSet<>();
        for (Criterion c : model.criteria) {
            if (c.coefficient != null) coefs.add(c.coefficient);
        }
        float[] out = new float[coefs.size()];
        int i = 0;
        for (Double cf : coefs) out[i++] = coefScale(cf, 0.33f);
        return out;
    }

    private void navigateNote(int delta) {
        if (notesList.isEmpty()) return;
        int n = notesList.size();
        notePos = ((notePos + delta) % n + n) % n;
        renderNote();
    }

    /** Tap the note on the carousel to edit it here; emptying it drops it from the list. */
    private void editCarouselNote() {
        if (notesList.isEmpty()) return;
        int[] pair = notesList.get(notePos);
        editNote(pair[0], pair[1], () -> {
            buildNotesList();
            if (notePos >= notesList.size()) notePos = Math.max(0, notesList.size() - 1);
            renderNote();
        });
    }

    /** "ALLER À…": jump to the current note's criterion page, arming the extra "‹ NOTES" button. */
    private void gotoNoteCriterion() {
        if (notesList.isEmpty()) return;
        int[] pair = notesList.get(notePos);
        noteReturnPos = notePos;
        studentIdx = pair[0];
        pageIdx = pair[1] + 1;
        overviewReturnPage = -1;
        overviewReturnSubpage = 0;
        showGrading();
        render();
    }

    /** The extra "‹ NOTES" button on a jumped-to criterion page: back to the carousel. */
    private void returnToNotes() {
        if (noteReturnPos >= 0) notePos = noteReturnPos;
        noteReturnPos = -1;
        showNotesScreen();
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
