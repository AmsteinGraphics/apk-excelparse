package com.example.apkexcelparse.xlsx;

import com.example.apkexcelparse.model.Criterion;
import com.example.apkexcelparse.model.GradingModel;
import com.example.apkexcelparse.model.Group;
import com.example.apkexcelparse.model.Student;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;

public class XlsxParser {

    private static final String EVALUATION_SHEET = "evaluation";
    private static final String CRITERIA_SHEET = "criteres_reviewed";

    // Excel fill colors that tag meaningful cells in the workbook.
    // Note: criterion labels display as pale blue in Excel but come from a themed fill
    // (Accent 1 + tint). POI's tint resolution differs from Excel's, so we match the
    // RGB POI reports (#E6E9EB), not the pale-blue Excel shows.
    private static final int[] MARK_CELL_RGB = {0xD9, 0xD9, 0xD9};        // light gray
    private static final int[] CRITERION_LABEL_RGB = {0xE6, 0xE9, 0xEB};  // themed blue as POI resolves it

    // Columns on the 'evaluation' sheet (0-indexed).
    private static final int COL_STUDENT_NUMBER = 0;
    private static final int COL_STUDENT_NAME = 1;
    public static final int COL_STUDENT_AVERAGE = 89; // Excel column CL

    // Columns on the 'criteres_reviewed' sheet (0-indexed).
    private static final int META_COL_GROUP = 1;
    private static final int META_COL_COEF = 2;
    private static final int META_COL_ID = 3;
    private static final int META_COL_CONTRACT = 4;
    private static final int META_COL_REMARKS = 5;

    public static GradingModel parse(XSSFWorkbook workbook) {
        Sheet eval = workbook.getSheet(EVALUATION_SHEET);
        if (eval == null) {
            throw new IllegalStateException("Sheet '" + EVALUATION_SHEET + "' not found");
        }
        Sheet meta = workbook.getSheet(CRITERIA_SHEET);

        Map<String, CriterionMeta> metaById = meta == null ? Collections.emptyMap() : readCriterionMeta(meta);

        // Find the criterion header row + column-per-criterion mapping by scanning pale-blue cells.
        Map<Integer, String> criterionColumns = new HashMap<>(); // column -> criterion id
        Integer criterionHeaderRow = null;
        Map<String, Integer> observedFills = new HashMap<>();
        for (Row row : eval) {
            for (Cell cell : row) {
                int[] rgb = extractCellRgb(cell);
                if (rgb == null) continue;
                if (!isNeutralWhite(rgb)) {
                    String key = String.format(Locale.ROOT, "#%02X%02X%02X", rgb[0], rgb[1], rgb[2]);
                    observedFills.merge(key, 1, Integer::sum);
                }
                if (!matchesRgb(rgb, CRITERION_LABEL_RGB)) continue;
                String id = readCellString(cell);
                if (id == null || id.isEmpty()) continue;
                criterionColumns.put(cell.getColumnIndex(), id);
                if (criterionHeaderRow == null || row.getRowNum() < criterionHeaderRow) {
                    criterionHeaderRow = row.getRowNum();
                }
            }
        }
        if (criterionHeaderRow == null || criterionColumns.isEmpty()) {
            throw new IllegalStateException(
                    "No criterion cells (pale blue " + hex(CRITERION_LABEL_RGB) + ") found. "
                            + "Colors observed on filled cells: " + topFills(observedFills, 8));
        }

        List<Integer> sortedColumns = new ArrayList<>(criterionColumns.keySet());
        Collections.sort(sortedColumns);

        List<Criterion> criteria = new ArrayList<>();
        for (int col : sortedColumns) {
            String id = criterionColumns.get(col);
            CriterionMeta m = metaById.get(id);
            criteria.add(new Criterion(
                    id,
                    col,
                    m != null ? m.groupName : null,
                    m != null ? m.coefficient : null,
                    m != null ? m.contract : null,
                    m != null ? m.remarks : null
            ));
        }

        List<Student> students = new ArrayList<>();
        for (Row row : eval) {
            if (row.getRowNum() <= criterionHeaderRow) continue;
            Cell numberCell = row.getCell(COL_STUDENT_NUMBER);
            if (!isNumeric(numberCell)) continue; // skip label rows like "coefficient du critère"
            Cell nameCell = row.getCell(COL_STUDENT_NAME);
            String name = readCellString(nameCell);
            if (name == null || name.isEmpty()) continue;
            String number = readCellString(numberCell);
            students.add(new Student(row.getRowNum(), number, name));
        }

        return new GradingModel(students, criteria, buildGroups(criteria));
    }

    /**
     * Collapse the ordered criteria list into groups: each maximal run of consecutive
     * criteria sharing the same group name is one group. The group's grade column is the
     * "sur 6" formula column immediately to the right of the group's last criterion column.
     */
    private static List<Group> buildGroups(List<Criterion> criteria) {
        List<Group> groups = new ArrayList<>();
        int i = 0;
        while (i < criteria.size()) {
            String name = criteria.get(i).groupName;
            int last = i;
            int maxCol = criteria.get(i).columnIndex;
            while (last + 1 < criteria.size()
                    && Objects.equals(criteria.get(last + 1).groupName, name)) {
                last++;
                maxCol = Math.max(maxCol, criteria.get(last).columnIndex);
            }
            groups.add(new Group(name, i, last, maxCol + 1));
            i = last + 1;
        }
        return groups;
    }

    /**
     * Read the current mark value stored in the cell for (student, criterion), or null if empty.
     */
    public static Double readMark(XSSFWorkbook workbook, Student student, Criterion criterion) {
        Sheet eval = workbook.getSheet(EVALUATION_SHEET);
        if (eval == null) return null;
        Row row = eval.getRow(student.rowIndex);
        if (row == null) return null;
        Cell cell = row.getCell(criterion.columnIndex);
        if (cell == null) return null;
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        return null;
    }

    /**
     * Write a mark value into the (student, criterion) cell. Only writes into cells tagged as mark cells.
     * Returns true if written, false if the target cell is not a valid mark cell.
     */
    public static boolean writeMark(XSSFWorkbook workbook, Student student, Criterion criterion, double value) {
        Sheet eval = workbook.getSheet(EVALUATION_SHEET);
        if (eval == null) return false;
        Row row = eval.getRow(student.rowIndex);
        if (row == null) row = eval.createRow(student.rowIndex);
        Cell cell = row.getCell(criterion.columnIndex);
        if (cell == null) return false; // don't create cells; only fill declared mark cells
        int[] rgb = extractCellRgb(cell);
        if (rgb == null || !matchesRgb(rgb, MARK_CELL_RGB)) return false;
        cell.setCellValue(value);
        return true;
    }

    /**
     * Clear the mark in the (student, criterion) cell so it reads as ungraded again. Only touches
     * cells tagged as mark cells. Returns true if cleared, false if the target is not a mark cell.
     */
    public static boolean eraseMark(XSSFWorkbook workbook, Student student, Criterion criterion) {
        Sheet eval = workbook.getSheet(EVALUATION_SHEET);
        if (eval == null) return false;
        Row row = eval.getRow(student.rowIndex);
        if (row == null) return false;
        Cell cell = row.getCell(criterion.columnIndex);
        if (cell == null) return false;
        int[] rgb = extractCellRgb(cell);
        if (rgb == null || !matchesRgb(rgb, MARK_CELL_RGB)) return false;
        cell.setBlank();
        return true;
    }

    /**
     * Return the Cell at (student row, columnIndex) on the evaluation sheet, or null if missing.
     * Useful for notifying a FormulaEvaluator after writing a mark.
     */
    public static Cell getEvalCell(XSSFWorkbook workbook, Student student, int columnIndex) {
        Sheet eval = workbook.getSheet(EVALUATION_SHEET);
        if (eval == null) return null;
        Row row = eval.getRow(student.rowIndex);
        if (row == null) return null;
        return row.getCell(columnIndex);
    }

    /**
     * Read a numeric cell value from the evaluation sheet, evaluating formulas if needed.
     * Returns null if the cell is missing/blank/non-numeric.
     */
    public static Double readNumericAt(XSSFWorkbook workbook, Student student, int columnIndex,
                                       FormulaEvaluator evaluator) {
        Cell cell = getEvalCell(workbook, student, columnIndex);
        if (cell == null) return null;
        try {
            if (cell.getCellType() == CellType.FORMULA) {
                if (evaluator != null) {
                    evaluator.evaluateFormulaCell(cell);
                }
                if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                    return cell.getNumericCellValue();
                }
                return null;
            }
            if (cell.getCellType() == CellType.NUMERIC) {
                return cell.getNumericCellValue();
            }
        } catch (Exception ignored) {
        }
        return null;
    }

    private static Map<String, CriterionMeta> readCriterionMeta(Sheet meta) {
        Map<String, CriterionMeta> byId = new HashMap<>();
        String lastGroup = null; // fill-down: group name is written once per group, blank on subsequent rows
        for (Row row : meta) {
            if (row.getRowNum() == 0) continue; // header row
            String id = readCellString(row.getCell(META_COL_ID));
            if (id == null || id.isEmpty()) continue;
            String rawGroup = readCellString(row.getCell(META_COL_GROUP));
            if (rawGroup != null && !rawGroup.isEmpty()) {
                lastGroup = rawGroup;
            }
            CriterionMeta m = new CriterionMeta();
            m.groupName = lastGroup;
            m.coefficient = readCellDouble(row.getCell(META_COL_COEF));
            m.contract = readCellString(row.getCell(META_COL_CONTRACT));
            m.remarks = readCellString(row.getCell(META_COL_REMARKS));
            byId.put(id, m);
        }
        return byId;
    }

    /**
     * Extract the fill foreground color of a cell as [R, G, B] (0..255), or null if unfilled.
     * Prefers the tinted RGB (what Excel actually displays) — for direct RGB fills with no
     * tint, POI returns the raw RGB unchanged.
     */
    private static int[] extractCellRgb(Cell cell) {
        if (!(cell instanceof XSSFCell)) return null;
        XSSFCellStyle style = ((XSSFCell) cell).getCellStyle();
        if (style == null) return null;
        XSSFColor color = style.getFillForegroundColorColor();
        if (color == null) return null;

        try {
            int[] tinted = fromBytes(color.getRGBWithTint());
            if (tinted != null) return tinted;
        } catch (Exception ignored) {
        }
        int[] rgb = fromBytes(color.getARGB());
        if (rgb == null) rgb = fromBytes(color.getRGB());
        return rgb;
    }

    private static int[] fromBytes(byte[] bytes) {
        if (bytes == null) return null;
        if (bytes.length == 4) {
            return new int[]{bytes[1] & 0xFF, bytes[2] & 0xFF, bytes[3] & 0xFF};
        }
        if (bytes.length == 3) {
            return new int[]{bytes[0] & 0xFF, bytes[1] & 0xFF, bytes[2] & 0xFF};
        }
        return null;
    }

    private static boolean matchesRgb(int[] rgb, int[] target) {
        int tol = 4;
        return Math.abs(rgb[0] - target[0]) <= tol
                && Math.abs(rgb[1] - target[1]) <= tol
                && Math.abs(rgb[2] - target[2]) <= tol;
    }

    private static boolean isNeutralWhite(int[] rgb) {
        return rgb[0] >= 250 && rgb[1] >= 250 && rgb[2] >= 250;
    }

    private static String hex(int[] rgb) {
        return String.format(Locale.ROOT, "#%02X%02X%02X", rgb[0], rgb[1], rgb[2]);
    }

    private static String topFills(Map<String, Integer> fills, int n) {
        if (fills.isEmpty()) return "(none)";
        return fills.entrySet().stream()
                .sorted((a, b) -> Integer.compare(b.getValue(), a.getValue()))
                .limit(n)
                .map(e -> e.getKey() + "×" + e.getValue())
                .reduce((a, b) -> a + ", " + b)
                .orElse("(none)");
    }

    private static String readCellString(Cell cell) {
        if (cell == null) return null;
        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) {
            type = cell.getCachedFormulaResultType();
        }
        switch (type) {
            case STRING:
                try {
                    return cell.getStringCellValue().trim();
                } catch (Exception ignored) {
                    return null;
                }
            case NUMERIC:
                try {
                    return numericToString(cell.getNumericCellValue());
                } catch (Exception ignored) {
                    return null;
                }
            default:
                return null;
        }
    }

    private static String numericToString(double d) {
        if (d == Math.floor(d) && !Double.isInfinite(d)) {
            return String.valueOf((long) d);
        }
        return String.valueOf(d);
    }

    private static boolean isNumeric(Cell cell) {
        if (cell == null) return false;
        CellType type = cell.getCellType();
        if (type == CellType.NUMERIC) return true;
        if (type == CellType.FORMULA) return cell.getCachedFormulaResultType() == CellType.NUMERIC;
        return false;
    }

    private static Double readCellDouble(Cell cell) {
        if (cell == null) return null;
        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                return cell.getNumericCellValue();
            }
            if (cell.getCellType() == CellType.STRING) {
                String s = cell.getStringCellValue().trim();
                if (s.isEmpty()) return null;
                return Double.parseDouble(s.replace(',', '.'));
            }
        } catch (Exception ignored) {
        }
        return null;
    }

    public static XSSFWorkbook open(InputStream in) throws IOException {
        return new XSSFWorkbook(in);
    }

    private static class CriterionMeta {
        String groupName;
        Double coefficient;
        String contract;
        String remarks;
    }
}
