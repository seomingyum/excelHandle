package com.example.demo.service.impl;

import com.example.demo.dto.ExcelCellDto;
import com.example.demo.dto.ExcelParseResponseDto;
import com.example.demo.service.ExcelParseService;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelParseServiceImpl implements ExcelParseService {

    @Override
    public ExcelParseResponseDto parseExcel(MultipartFile file) throws Exception {
        validateExcelFile(file);

        try (
                InputStream is = file.getInputStream();
                Workbook workbook = createWorkbook(file, is)
        ) {
            Sheet sheet = workbook.getSheetAt(0);

            List<List<ExcelCellDto>> headers = new ArrayList<>();
            List<List<ExcelCellDto>> rows = new ArrayList<>();

            if (sheet == null || sheet.getPhysicalNumberOfRows() == 0) {
                return new ExcelParseResponseDto(headers, rows);
            }

            int firstRowNum = sheet.getFirstRowNum();
            Row firstRow = sheet.getRow(firstRowNum);

            if (firstRow == null) {
                return new ExcelParseResponseDto(headers, rows);
            }

            int columnCount = getMaxColumnCount(sheet);
            if (columnCount <= 0) {
                return new ExcelParseResponseDto(headers, rows);
            }

            // 첫 번째 행을 헤더로 사용
            headers.add(buildRowCells(sheet, firstRowNum, columnCount));

            // 나머지 행을 데이터로 사용
            for (int rowIndex = firstRowNum + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                if (isRowEmpty(sheet, rowIndex, columnCount)) {
                    continue;
                }

                rows.add(buildRowCells(sheet, rowIndex, columnCount));
            }

            return new ExcelParseResponseDto(headers, rows);
        }
    }

    private List<ExcelCellDto> buildRowCells(Sheet sheet, int rowIndex, int columnCount) {
        List<ExcelCellDto> rowCells = new ArrayList<>();
        Row row = sheet.getRow(rowIndex);

        for (int cellIndex = 0; cellIndex < columnCount; cellIndex++) {
            CellRangeAddress mergedRegion = getMergedRegion(sheet, rowIndex, cellIndex);

            if (mergedRegion != null) {
                boolean isTopLeft = mergedRegion.getFirstRow() == rowIndex
                        && mergedRegion.getFirstColumn() == cellIndex;

                if (isTopLeft) {
                    Cell cell = null;
                    Row firstRow = sheet.getRow(mergedRegion.getFirstRow());
                    if (firstRow != null) {
                        cell = firstRow.getCell(mergedRegion.getFirstColumn(), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    }

                    int rowspan = mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1;
                    int colspan = mergedRegion.getLastColumn() - mergedRegion.getFirstColumn() + 1;

                    rowCells.add(new ExcelCellDto(
                            getCellStringValue(cell),
                            rowspan,
                            colspan,
                            false
                    ));
                } else {
                    rowCells.add(new ExcelCellDto("", 1, 1, true));
                }
            } else {
                Cell cell = null;
                if (row != null) {
                    cell = row.getCell(cellIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                }

                rowCells.add(new ExcelCellDto(
                        getCellStringValue(cell),
                        1,
                        1,
                        false
                ));
            }
        }

        return rowCells;
    }

    private int getMaxColumnCount(Sheet sheet) {
        int maxColumnCount = 0;

        for (int rowIndex = sheet.getFirstRowNum(); rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null && row.getLastCellNum() > maxColumnCount) {
                maxColumnCount = row.getLastCellNum();
            }
        }

        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            int mergedColumnCount = merged.getLastColumn() + 1;
            if (mergedColumnCount > maxColumnCount) {
                maxColumnCount = mergedColumnCount;
            }
        }

        return maxColumnCount;
    }

    private CellRangeAddress getMergedRegion(Sheet sheet, int rowIndex, int columnIndex) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.isInRange(rowIndex, columnIndex)) {
                return region;
            }
        }
        return null;
    }

    private Workbook createWorkbook(MultipartFile file, InputStream is) throws Exception {
        String originalFilename = file.getOriginalFilename();

        if (originalFilename == null) {
            throw new IllegalArgumentException("파일명이 없습니다.");
        }

        String lowerName = originalFilename.toLowerCase();

        if (lowerName.endsWith(".xlsx")) {
            return new XSSFWorkbook(is);
        }

        if (lowerName.endsWith(".xls")) {
            return new HSSFWorkbook(is);
        }

        throw new IllegalArgumentException("지원하지 않는 파일 형식입니다. xls 또는 xlsx 파일만 업로드하세요.");
    }

    private void validateExcelFile(MultipartFile file) {
        if (file == null || file.isEmpty()) {
            throw new IllegalArgumentException("업로드된 파일이 없습니다.");
        }

        String originalFilename = file.getOriginalFilename();
        if (originalFilename == null) {
            throw new IllegalArgumentException("파일명이 없습니다.");
        }

        String lowerName = originalFilename.toLowerCase();
        if (!lowerName.endsWith(".xlsx") && !lowerName.endsWith(".xls")) {
            throw new IllegalArgumentException("엑셀 파일만 업로드할 수 있습니다.");
        }
    }

    private boolean isRowEmpty(Sheet sheet, int rowIndex, int columnCount) {
        for (int cellIndex = 0; cellIndex < columnCount; cellIndex++) {
            CellRangeAddress mergedRegion = getMergedRegion(sheet, rowIndex, cellIndex);

            if (mergedRegion != null) {
                Row mergedStartRow = sheet.getRow(mergedRegion.getFirstRow());
                Cell mergedCell = null;
                if (mergedStartRow != null) {
                    mergedCell = mergedStartRow.getCell(mergedRegion.getFirstColumn(), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                }

                if (!getCellStringValue(mergedCell).isBlank()) {
                    return false;
                }
            } else {
                Row row = sheet.getRow(rowIndex);
                Cell cell = null;
                if (row != null) {
                    cell = row.getCell(cellIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                }

                if (!getCellStringValue(cell).isBlank()) {
                    return false;
                }
            }
        }

        return true;
    }

    private String getCellStringValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();

            case NUMERIC -> {
                double num = cell.getNumericCellValue();

                if (num == (long) num) {
                    long longValue = (long) num;
                    String numericText = String.valueOf(longValue);

                    if (numericText.matches("\\d{8}")) {
                        String yyyy = numericText.substring(0, 4);
                        String mm = numericText.substring(4, 6);
                        String dd = numericText.substring(6, 8);

                        if (isValidYyyyMmDd(yyyy, mm, dd)) {
                            yield numericText;
                        }
                    }
                }

                if (DateUtil.isCellDateFormatted(cell)) {
                    yield new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(cell.getDateCellValue());
                }

                if (num == (long) num) {
                    yield String.valueOf((long) num);
                }

                yield String.valueOf(num);
            }

            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());

            case FORMULA -> {
                FormulaEvaluator evaluator = cell.getSheet()
                        .getWorkbook()
                        .getCreationHelper()
                        .createFormulaEvaluator();

                CellValue evaluated = evaluator.evaluate(cell);

                if (evaluated == null) {
                    yield "";
                }

                yield switch (evaluated.getCellType()) {
                    case STRING -> evaluated.getStringValue().trim();

                    case NUMERIC -> {
                        double num = evaluated.getNumberValue();

                        if (num == (long) num) {
                            long longValue = (long) num;
                            String numericText = String.valueOf(longValue);

                            if (numericText.matches("\\d{8}")) {
                                String yyyy = numericText.substring(0, 4);
                                String mm = numericText.substring(4, 6);
                                String dd = numericText.substring(6, 8);

                                if (isValidYyyyMmDd(yyyy, mm, dd)) {
                                    yield yyyy + "-" + mm + "-" + dd;
                                }
                            }

                            yield String.valueOf(longValue);
                        }

                        yield String.valueOf(num);
                    }

                    case BOOLEAN -> String.valueOf(evaluated.getBooleanValue());
                    case BLANK -> "";
                    default -> "";
                };
            }

            case BLANK -> "";
            default -> "";
        };
    }

    private boolean isValidYyyyMmDd(String yyyy, String mm, String dd) {
        try {
            LocalDate.of(
                    Integer.parseInt(yyyy),
                    Integer.parseInt(mm),
                    Integer.parseInt(dd)
            );
            return true;
        } catch (Exception e) {
            return false;
        }
    }
}