package com.example.demo.dto;

import java.util.List;

public class ExcelParseResponseDto {

    private List<List<ExcelCellDto>> headers;
    private List<List<ExcelCellDto>> rows;

    public ExcelParseResponseDto() {
    }

    public ExcelParseResponseDto(List<List<ExcelCellDto>> headers, List<List<ExcelCellDto>> rows) {
        this.headers = headers;
        this.rows = rows;
    }

    public List<List<ExcelCellDto>> getHeaders() {
        return headers;
    }

    public void setHeaders(List<List<ExcelCellDto>> headers) {
        this.headers = headers;
    }

    public List<List<ExcelCellDto>> getRows() {
        return rows;
    }

    public void setRows(List<List<ExcelCellDto>> rows) {
        this.rows = rows;
    }
}