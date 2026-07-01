package com.example.apkexcelparse.model;

import java.util.List;

public class GradingModel {
    public final List<Student> students;
    public final List<Criterion> criteria;
    public final List<Group> groups;

    public GradingModel(List<Student> students, List<Criterion> criteria, List<Group> groups) {
        this.students = students;
        this.criteria = criteria;
        this.groups = groups;
    }

    /** Return the group a criterion (by its index in {@link #criteria}) belongs to, or null. */
    public Group groupForCriterion(int criterionIndex) {
        for (Group g : groups) {
            if (criterionIndex >= g.firstCriterionIndex && criterionIndex <= g.lastCriterionIndex) {
                return g;
            }
        }
        return null;
    }
}
