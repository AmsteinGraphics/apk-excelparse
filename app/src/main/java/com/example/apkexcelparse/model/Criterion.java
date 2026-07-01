package com.example.apkexcelparse.model;

public class Criterion {
    public final String id;
    public final int columnIndex;
    public final String groupName;
    public final Double coefficient;
    public final String contract;
    public final String remarks;

    public Criterion(String id, int columnIndex, String groupName, Double coefficient, String contract, String remarks) {
        this.id = id;
        this.columnIndex = columnIndex;
        this.groupName = groupName;
        this.coefficient = coefficient;
        this.contract = contract;
        this.remarks = remarks;
    }
}
