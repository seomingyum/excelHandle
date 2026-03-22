package com.example.demo.dto;

public class ExcelCellDto {

    private String value;
    private int rowspan;
    private int colspan;
    private boolean hidden;
    private int rowIndex;
    private int colIndex;
    private boolean formulaCell;
    private String cellType;

    public ExcelCellDto() {
    }

    public ExcelCellDto(String value, int rowspan, int colspan, boolean hidden,
                        int rowIndex, int colIndex, boolean formulaCell, String cellType) {
        this.value = value;
        this.rowspan = rowspan;
        this.colspan = colspan;
        this.hidden = hidden;
        this.rowIndex = rowIndex;
        this.colIndex = colIndex;
        this.formulaCell = formulaCell;
        this.cellType = cellType;
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

    public boolean isFormulaCell() {
        return formulaCell;
    }

    public void setFormulaCell(boolean formulaCell) {
        this.formulaCell = formulaCell;
    }

    public String getCellType() {
        return cellType;
    }

    public void setCellType(String cellType) {
        this.cellType = cellType;
    }
}