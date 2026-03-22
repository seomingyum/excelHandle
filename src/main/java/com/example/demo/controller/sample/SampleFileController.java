package com.example.demo.controller.sample;

import jakarta.servlet.http.HttpServletResponse;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Controller;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.bind.annotation.GetMapping;

import java.io.InputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;

@Controller
public class SampleFileController {

    @GetMapping("/sample/download")
    public void downloadSampleExcel(HttpServletResponse response) throws Exception {
        ClassPathResource resource = new ClassPathResource("static/sample/sample.xlsx");

        if (!resource.exists()) {
            response.setStatus(HttpServletResponse.SC_NOT_FOUND);
            response.getWriter().write("샘플 엑셀 파일이 없습니다.");
            return;
        }

        String fileName = URLEncoder.encode("sample.xlsx", StandardCharsets.UTF_8)
                .replaceAll("\\+", "%20");

        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename*=UTF-8''" + fileName);
        response.setContentLengthLong(resource.contentLength());

        try (InputStream in = resource.getInputStream()) {
            FileCopyUtils.copy(in, response.getOutputStream());
        }
    }
}