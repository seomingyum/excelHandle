package com.example.demo.dto;

public class ExcelChangeDto {

    private int rowIndex;
    private int colIndex;
    private String value;

    public ExcelChangeDto() {
    }

    public ExcelChangeDto(int rowIndex, int colIndex, String value) {
        this.rowIndex = rowIndex;
        this.colIndex = colIndex;
        this.value = value;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public int getColIndex() {
        return colIndex;
    }

    public void setColIndex(int colIndex) {
        this.colIndex = colIndex;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }
}