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

    public Group(String name, int firstCriterionIndex, int lastCriterionIndex, int averageColumnIndex) {
        this.name = name;
        this.firstCriterionIndex = firstCriterionIndex;
        this.lastCriterionIndex = lastCriterionIndex;
        this.averageColumnIndex = averageColumnIndex;
    }
}
