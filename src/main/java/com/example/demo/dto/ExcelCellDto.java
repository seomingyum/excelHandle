package com.example.demo.dto;

public class ExcelCellDto {

    private String value;
    private int rowspan;
    private int colspan;
    private boolean hidden;

    public ExcelCellDto() {
    }

    public ExcelCellDto(String value, int rowspan, int colspan, boolean hidden) {
        this.value = value;
        this.rowspan = rowspan;
        this.colspan = colspan;
        this.hidden = hidden;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public int getRowspan() {
        return rowspan;
    }

    public void setRowspan(int rowspan) {
        this.rowspan = rowspan;
    }

    public int getColspan() {
        return colspan;
    }

    public void setColspan(int colspan) {
        this.colspan = colspan;
    }

    public boolean isHidden() {
        return hidden;
    }

    public void setHidden(boolean hidden) {
        this.hidden = hidden;
    }
}