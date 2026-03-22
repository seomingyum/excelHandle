package com.example.demo.controller.excel;

import com.example.demo.dto.ExcelDownloadRequestDto;
import com.example.demo.dto.ExcelUploadResponseDto;
import com.example.demo.service.ExcelMemoryService;
import jakarta.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Map;

@RestController
@RequestMapping("/excel")
public class ExcelController {
	
	@Autowired
    private ExcelMemoryService excelMemoryService;

    @PostMapping(value = "/upload", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ExcelUploadResponseDto upload(@RequestParam("file") MultipartFile file) throws Exception {
        return excelMemoryService.uploadExcel(file);
    }

    @PostMapping("/download")
    public void download(@RequestBody ExcelDownloadRequestDto request,
                         HttpServletResponse response) throws Exception {

        String downloadFileName = excelMemoryService.peekDownloadFileName();
        byte[] fileBytes = excelMemoryService.downloadExcel(request);

        String encodedFileName = URLEncoder.encode(downloadFileName, StandardCharsets.UTF_8)
                .replaceAll("\\+", "%20");

        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename*=UTF-8''" + encodedFileName);
        response.setContentLength(fileBytes.length);

        response.getOutputStream().write(fileBytes);
        response.getOutputStream().flush();
    }

    @ExceptionHandler(Exception.class)
    @ResponseBody
    public Map<String, Object> handleException(Exception e, HttpServletResponse response) {
        response.setStatus(400);

        Map<String, Object> result = new HashMap<>();
        result.put("message", e.getMessage());
        return result;
    }
}