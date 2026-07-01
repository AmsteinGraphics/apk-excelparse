package com.example.apkexcelparse.model;

public class Student {
    public final int rowIndex;
    public final String number;
    public final String name;

    public Student(int rowIndex, String number, String name) {
        this.rowIndex = rowIndex;
        this.number = number;
        this.name = name;
    }
}
