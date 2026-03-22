package com.example.demo.service;

import com.example.demo.dto.ExcelParseResponseDto;
import org.springframework.web.multipart.MultipartFile;

public interface ExcelParseService {
    ExcelParseResponseDto parseExcel(MultipartFile file) throws Exception;
}