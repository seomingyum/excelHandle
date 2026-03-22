package com.example.demo.dto;

import java.util.List;

public class ExcelUploadResponseDto {

    private String token;
    private String originalFileName;
    private List<List<ExcelCellDto>> rows;

    public ExcelUploadResponseDto() {
    }

    public ExcelUploadResponseDto(String token, String originalFileName, List<List<ExcelCellDto>> rows) {
        this.token = token;
        this.originalFileName = originalFileName;
        this.rows = rows;
    }

    public String getToken() {
        return token;
    }

    public void setToken(String token) {
        this.token = token;
    }

    public String getOriginalFileName() {
        return originalFileName;
    }

    public void setOriginalFileName(String originalFileName) {
        this.originalFileName = originalFileName;
    }

    public List<List<ExcelCellDto>> getRows() {
        return rows;
    }

    public void setRows(List<List<ExcelCellDto>> rows) {
        this.rows = rows;
    }
}