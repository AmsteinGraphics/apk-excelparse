package com.example.apkexcelparse.model;

/**
 * A contiguous run of criteria that share a group name on the evaluation sheet.
 * Groups are visual blocks like "2.3 projections planes de l'objet" (criteria 4..8).
 *
 * The group's grade lives in the "sur 6" formula column immediately to the right of
 * the group's last criterion column — already expressed on the /6 scale, used as-is.
 */
public class Group {
    public final String name;
    public final int firstCriterionIndex; // inclusive index into GradingModel.criteria
    public final int lastCriterionIndex;  // inclusive
    public final int averageColumnIndex;  // evaluation-sheet column of the "sur 6" group grade
    // Group weight from the "coefficient du groupe" row (evaluation row 2), read at the group's
    // first-criterion column. Used to weight the group's /6 grade in the general average, exactly
    // as Excel's CK formula does. May be null if the workbook has no such value.
    public final Double coefficient;

    public Group(String name, int firstCriterionIndex, int lastCriterionIndex,
                 int averageColumnIndex, Double coefficient) {
        this.name = name;
        this.firstCriterionIndex = firstCriterionIndex;
        this.lastCriterionIndex = lastCriterionIndex;
        this.averageColumnIndex = averageColumnIndex;
        this.coefficient = coefficient;
    }
}
