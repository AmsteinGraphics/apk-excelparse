package com.example.apkexcelparse.model;

import java.util.List;

public class GradingModel {
    public final List<Student> students;
    public final List<Criterion> criteria;

    public GradingModel(List<Student> students, List<Criterion> criteria) {
        this.students = students;
        this.criteria = criteria;
    }
}
