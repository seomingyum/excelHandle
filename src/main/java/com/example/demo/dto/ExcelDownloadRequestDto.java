package com.example.demo.dto;

import java.util.List;

public class ExcelDownloadRequestDto {

    private String token;
    private List<ExcelChangeDto> changes;

    public ExcelDownloadRequestDto() {
    }

    public ExcelDownloadRequestDto(String token, List<ExcelChangeDto> changes) {
        this.token = token;
        this.changes = changes;
    }

    public String getToken() {
        return token;
    }

    public void setToken(String token) {
        this.token = token;
    }

    public List<ExcelChangeDto> getChanges() {
        return changes;
    }

    public void setChanges(List<ExcelChangeDto> changes) {
        this.changes = changes;
    }
}