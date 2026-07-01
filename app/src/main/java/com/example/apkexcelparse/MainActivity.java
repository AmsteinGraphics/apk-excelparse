package com.example.apkexcelparse;

import android.content.Intent;
import android.content.SharedPreferences;
import android.net.Uri;
import android.os.Bundle;
import android.os.Handler;
import android.os.Looper;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.widget.ProgressBar;
import android.widget.TextView;
import android.widget.Toast;

import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;

import com.example.apkexcelparse.model.Criterion;
import com.example.apkexcelparse.model.GradingModel;
import com.example.apkexcelparse.model.Student;
import com.example.apkexcelparse.ui.DotProgressView;
import com.example.apkexcelparse.xlsx.XlsxParser;
import com.google.android.material.button.MaterialButton;
import com.google.android.material.slider.Slider;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

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
    private static final String WORKING_COPY_NAME = "working_copy.xlsx";
    private static final String RELEASES_API = "https://api.github.com/repos/AmsteinGraphics/apk-excelparse/releases/latest";
    private static final String RELEASES_FALLBACK = "https://github.com/AmsteinGraphics/apk-excelparse/releases";

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
    private TextView markLabel;
    private Slider markSlider;
    private MaterialButton prevButton;
    private MaterialButton nextButton;
    private MaterialButton prevStudentButton;
    private MaterialButton nextStudentButton;
    private MaterialButton saveButton;
    private View dirtyIndicator;
    private TextView studentAverage;
    private DotProgressView dotProgressView;

    private XSSFWorkbook workbook;
    private FormulaEvaluator formulaEvaluator;
    private GradingModel model;
    private int studentIdx;
    private int criterionIdx;
    private boolean dirty;
    private boolean settingSliderProgrammatically;

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
        markLabel = findViewById(R.id.markLabel);
        markSlider = findViewById(R.id.markSlider);
        prevButton = findViewById(R.id.prevButton);
        nextButton = findViewById(R.id.nextButton);
        prevStudentButton = findViewById(R.id.prevStudentButton);
        nextStudentButton = findViewById(R.id.nextStudentButton);
        saveButton = findViewById(R.id.saveButton);
        dirtyIndicator = findViewById(R.id.dirtyIndicator);
        studentAverage = findViewById(R.id.studentAverage);
        dotProgressView = findViewById(R.id.dotProgressView);

        findViewById(R.id.pickButton).setOnClickListener(v -> launchPicker());
        prevButton.setOnClickListener(v -> navigate(-1));
        nextButton.setOnClickListener(v -> navigate(1));
        prevStudentButton.setOnClickListener(v -> navigateStudent(-1));
        nextStudentButton.setOnClickListener(v -> navigateStudent(1));
        saveButton.setOnClickListener(v -> {
            if (!dirty) {
                Toast.makeText(this, "Nothing to save", Toast.LENGTH_SHORT).show();
                return;
            }
            saveAsync(() -> {
                updateDirtyUi();
                Toast.makeText(this, "Saved", Toast.LENGTH_SHORT).show();
            });
        });

        markSlider.addOnChangeListener((slider, value, fromUser) -> {
            if (settingSliderProgrammatically) return;
            markLabel.setText(formatMark(value));
            applyMarkToWorkbook(value);
        });

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

    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        menu.add(0, 1, 0, R.string.change_file);
        menu.add(0, 2, 1, R.string.check_updates);
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(@androidx.annotation.NonNull MenuItem item) {
        if (item.getItemId() == 1) {
            launchPicker();
            return true;
        }
        if (item.getItemId() == 2) {
            checkForUpdates();
            return true;
        }
        return super.onOptionsItemSelected(item);
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
                    criterionIdx = 0;
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

    private void navigate(int delta) {
        if (model == null) return;
        int total = model.students.size() * model.criteria.size();
        int flat = studentIdx * model.criteria.size() + criterionIdx + delta;
        if (flat < 0) flat = 0;
        if (flat >= total) flat = total - 1;
        studentIdx = flat / model.criteria.size();
        criterionIdx = flat % model.criteria.size();
        render();
    }

    private void navigateStudent(int delta) {
        if (model == null) return;
        int newIdx = studentIdx + delta;
        if (newIdx < 0) newIdx = 0;
        if (newIdx >= model.students.size()) newIdx = model.students.size() - 1;
        if (newIdx == studentIdx) return;
        studentIdx = newIdx;
        render();
    }

    private void render() {
        if (model == null) return;
        Student s = model.students.get(studentIdx);
        Criterion c = model.criteria.get(criterionIdx);

        studentCounter.setText(String.format(Locale.getDefault(),
                "étudiant %d / %d", studentIdx + 1, model.students.size()));
        studentName.setText(s.name);

        criterionGroup.setText(c.groupName != null ? c.groupName : "");
        String coefText = c.coefficient != null
                ? String.format(Locale.getDefault(), "critère %s  ·  coefficient %s", c.id, formatCoef(c.coefficient))
                : String.format(Locale.getDefault(), "critère %s", c.id);
        criterionIdCoef.setText(coefText);
        criterionContract.setText(c.contract != null ? c.contract : "");
        criterionRemarks.setText(c.remarks != null ? c.remarks : "");

        Double existing = XlsxParser.readMark(workbook, s, c);
        float value = existing != null ? existing.floatValue() : 0f;
        if (value < 0f) value = 0f;
        if (value > 1f) value = 1f;
        settingSliderProgrammatically = true;
        markSlider.setValue(Math.round(value * 4f) / 4f);
        settingSliderProgrammatically = false;
        markLabel.setText(existing != null ? formatMark(value) : "—");

        refreshDotsForStudent();
        refreshStudentAverage();
    }

    private void refreshDotsForStudent() {
        if (model == null || workbook == null || dotProgressView == null) return;
        Student s = model.students.get(studentIdx);
        int n = model.criteria.size();
        int[] buckets = new int[n];
        for (int i = 0; i < n; i++) {
            Double mark = XlsxParser.readMark(workbook, s, model.criteria.get(i));
            buckets[i] = markToBucket(mark);
        }
        dotProgressView.setValues(buckets);
    }

    private void refreshStudentAverage() {
        if (model == null || workbook == null || studentAverage == null) return;
        Student s = model.students.get(studentIdx);
        Double avg = XlsxParser.readNumericAt(workbook, s, XlsxParser.COL_STUDENT_AVERAGE, formulaEvaluator);
        if (avg == null || avg.isNaN()) {
            studentAverage.setText("");
        } else {
            studentAverage.setText(String.format(Locale.getDefault(), "%.2f / 6", avg));
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
        if (model == null || workbook == null) return;
        Student s = model.students.get(studentIdx);
        Criterion c = model.criteria.get(criterionIdx);
        boolean written = XlsxParser.writeMark(workbook, s, c, value);
        if (written) {
            dirty = true;
            updateDirtyUi();
            if (formulaEvaluator != null) {
                Cell cell = XlsxParser.getEvalCell(workbook, s, c.columnIndex);
                if (cell != null) formulaEvaluator.notifyUpdateCell(cell);
            }
            dotProgressView.setValueAt(criterionIdx, markToBucket((double) value));
            refreshStudentAverage();
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
            mainHandler.post(() -> {
                dirty = false;
                updateDirtyUi();
                if (then != null) then.run();
            });
        } catch (Exception e) {
            mainHandler.post(() -> {
                Toast.makeText(this, "Save failed: " + e.getMessage(), Toast.LENGTH_LONG).show();
                if (then != null) then.run();
            });
        }
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

    private static String formatMark(float value) {
        return String.format(Locale.getDefault(), "%.2f", value);
    }

    private static String formatCoef(Double coef) {
        if (coef == null) return "";
        if (coef == Math.floor(coef)) return String.valueOf(coef.intValue());
        return String.format(Locale.getDefault(), "%.2f", coef);
    }

    private void checkForUpdates() {
        executor.execute(() -> {
            try {
                URL url = new URL(RELEASES_API);
                HttpURLConnection connection = (HttpURLConnection) url.openConnection();
                connection.setRequestMethod("GET");
                connection.setRequestProperty("Accept", "application/vnd.github+json");
                int code = connection.getResponseCode();
                if (code != HttpURLConnection.HTTP_OK) throw new RuntimeException("HTTP " + code);
                StringBuilder builder = new StringBuilder();
                try (BufferedReader reader = new BufferedReader(new InputStreamReader(connection.getInputStream()))) {
                    String line;
                    while ((line = reader.readLine()) != null) builder.append(line);
                }
                JSONObject release = new JSONObject(builder.toString());
                String htmlUrl = release.optString("html_url", RELEASES_FALLBACK);
                String tagName = release.optString("tag_name", "unknown");
                mainHandler.post(() -> {
                    startActivity(new Intent(Intent.ACTION_VIEW, Uri.parse(htmlUrl)));
                    Toast.makeText(this, "Latest release: " + tagName, Toast.LENGTH_SHORT).show();
                });
            } catch (Exception e) {
                mainHandler.post(() -> Toast.makeText(this,
                        "Update check failed: " + e.getMessage(), Toast.LENGTH_LONG).show());
            }
        });
    }
}
