package com.example.demo.service;

import com.example.demo.dto.ExcelDownloadRequestDto;
import com.example.demo.dto.ExcelUploadResponseDto;
import org.springframework.web.multipart.MultipartFile;

public interface ExcelMemoryService {

    ExcelUploadResponseDto uploadExcel(MultipartFile file) throws Exception;

    byte[] downloadExcel(ExcelDownloadRequestDto request) throws Exception;

    String peekDownloadFileName();
}