package com.example.demo.exception;

import jakarta.servlet.http.HttpServletResponse;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;

import java.util.HashMap;
import java.util.Map;

@RestControllerAdvice
public class GlobalExceptionHandler {

    @ExceptionHandler(IllegalArgumentException.class)
    public Map<String, Object> handleIllegalArgumentException(IllegalArgumentException e,
                                                              HttpServletResponse response) {
        response.setStatus(400);

        Map<String, Object> result = new HashMap<>();
        result.put("message", e.getMessage());
        return result;
    }

    @ExceptionHandler(Exception.class)
    public Map<String, Object> handleException(Exception e,
                                               HttpServletResponse response) {
        response.setStatus(500);

        Map<String, Object> result = new HashMap<>();
        result.put("message", "서버 내부 오류가 발생했습니다.");
        return result;
    }
}