package com.example.apkexcelparse.xlsx;

import com.example.apkexcelparse.model.Criterion;
import com.example.apkexcelparse.model.GradingModel;
import com.example.apkexcelparse.model.Student;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
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
import java.util.Map;

public class XlsxParser {

    private static final String EVALUATION_SHEET = "evaluation";
    private static final String CRITERIA_SHEET = "criteres_reviewed";

    // Excel fill colors that tag meaningful cells in the workbook.
    private static final int[] MARK_CELL_RGB = {0xD9, 0xD9, 0xD9};        // light gray
    private static final int[] CRITERION_LABEL_RGB = {0xDA, 0xE9, 0xF8};  // pale blue

    // Columns on the 'evaluation' sheet (0-indexed).
    private static final int COL_STUDENT_NUMBER = 0;
    private static final int COL_STUDENT_NAME = 1;

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
        for (Row row : eval) {
            for (Cell cell : row) {
                if (!isBackgroundColor(cell, CRITERION_LABEL_RGB)) continue;
                String id = readCellString(cell);
                if (id == null || id.isEmpty()) continue;
                criterionColumns.put(cell.getColumnIndex(), id);
                if (criterionHeaderRow == null || row.getRowNum() < criterionHeaderRow) {
                    criterionHeaderRow = row.getRowNum();
                }
            }
        }
        if (criterionHeaderRow == null || criterionColumns.isEmpty()) {
            throw new IllegalStateException("No criterion cells (pale blue) found in evaluation sheet");
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
            Cell nameCell = row.getCell(COL_STUDENT_NAME);
            String name = readCellString(nameCell);
            if (name == null || name.isEmpty()) continue;
            String number = readCellString(row.getCell(COL_STUDENT_NUMBER));
            students.add(new Student(row.getRowNum(), number, name));
        }

        return new GradingModel(students, criteria);
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
        if (!isBackgroundColor(cell, MARK_CELL_RGB)) return false;
        cell.setCellValue(value);
        return true;
    }

    private static Map<String, CriterionMeta> readCriterionMeta(Sheet meta) {
        Map<String, CriterionMeta> byId = new HashMap<>();
        for (Row row : meta) {
            if (row.getRowNum() == 0) continue; // header row
            String id = readCellString(row.getCell(META_COL_ID));
            if (id == null || id.isEmpty()) continue;
            CriterionMeta m = new CriterionMeta();
            m.groupName = readCellString(row.getCell(META_COL_GROUP));
            m.coefficient = readCellDouble(row.getCell(META_COL_COEF));
            m.contract = readCellString(row.getCell(META_COL_CONTRACT));
            m.remarks = readCellString(row.getCell(META_COL_REMARKS));
            byId.put(id, m);
        }
        return byId;
    }

    private static boolean isBackgroundColor(Cell cell, int[] rgb) {
        if (!(cell instanceof XSSFCell)) return false;
        XSSFCellStyle style = ((XSSFCell) cell).getCellStyle();
        if (style == null) return false;
        XSSFColor color = style.getFillForegroundColorColor();
        if (color == null) return false;
        byte[] argb = color.getRGB();
        if (argb == null || argb.length < 3) return false;
        int offset = argb.length == 4 ? 1 : 0;
        return (argb[offset] & 0xFF) == rgb[0]
                && (argb[offset + 1] & 0xFF) == rgb[1]
                && (argb[offset + 2] & 0xFF) == rgb[2];
    }

    private static String readCellString(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                double d = cell.getNumericCellValue();
                if (d == Math.floor(d) && !Double.isInfinite(d)) {
                    return String.valueOf((long) d);
                }
                return String.valueOf(d);
            case FORMULA:
                try {
                    return cell.getStringCellValue().trim();
                } catch (IllegalStateException ex) {
                    try {
                        return String.valueOf(cell.getNumericCellValue());
                    } catch (IllegalStateException ignored) {
                        return null;
                    }
                }
            default:
                return null;
        }
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
