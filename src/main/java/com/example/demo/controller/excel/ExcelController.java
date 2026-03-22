package com.example.demo.controller.excel;

import com.example.demo.dto.ExcelParseResponseDto;
import com.example.demo.service.ExcelParseService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.Map;

@RestController
@RequestMapping("/excel")
public class ExcelController {
	
	@Autowired
    private ExcelParseService excelParseService;

    @PostMapping("/upload")
    public ResponseEntity<?> uploadExcel(@RequestParam("file") MultipartFile file) {
        try {
            ExcelParseResponseDto result = excelParseService.parseExcel(file);
            return ResponseEntity.ok(result);
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(Map.of(
                    "message", e.getMessage()
            ));
        }
    }
}