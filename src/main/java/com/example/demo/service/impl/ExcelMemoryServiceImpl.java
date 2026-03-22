package com.example.demo.service.impl;

import com.example.demo.dto.ExcelCellDto;
import com.example.demo.dto.ExcelChangeDto;
import com.example.demo.dto.ExcelDownloadRequestDto;
import com.example.demo.dto.ExcelUploadResponseDto;
import com.example.demo.service.ExcelMemoryService;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.locks.ReentrantLock;

@Service
public class ExcelMemoryServiceImpl implements ExcelMemoryService {

    private final ReentrantLock lock = new ReentrantLock();
    private StoredWorkbook currentWorkbook;

    @Override
    public ExcelUploadResponseDto uploadExcel(MultipartFile file) throws Exception {
        validateExcelFile(file);

        lock.lock();
        try {
            clearCurrentWorkbook();

            Workbook workbook;
            try (InputStream is = file.getInputStream()) {
                workbook = createWorkbook(file, is);
            }

            Workbook clonedWorkbook = cloneWorkbook(workbook, file.getOriginalFilename());
            workbook.close();

            String token = UUID.randomUUID().toString();

            currentWorkbook = new StoredWorkbook(
                    token,
                    file.getOriginalFilename(),
                    clonedWorkbook
            );

            Sheet firstSheet = clonedWorkbook.getSheetAt(0);
            List<List<ExcelCellDto>> previewRows = buildPreviewRows(firstSheet);

            return new ExcelUploadResponseDto(token, file.getOriginalFilename(), previewRows);
        } finally {
            lock.unlock();
        }
    }

    @Override
    public byte[] downloadExcel(ExcelDownloadRequestDto request) throws Exception {
        lock.lock();
        try {
            if (currentWorkbook == null || currentWorkbook.workbook == null) {
                throw new IllegalStateException("현재 업로드된 엑셀 파일이 없습니다.");
            }

            if (request == null || request.getToken() == null || request.getToken().isBlank()) {
                throw new IllegalArgumentException("유효하지 않은 다운로드 요청입니다.");
            }

            if (!currentWorkbook.token.equals(request.getToken())) {
                throw new IllegalArgumentException("현재 메모리에 있는 엑셀 파일과 토큰이 일치하지 않습니다.");
            }

            applyChanges(currentWorkbook.workbook, request.getChanges());

            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                currentWorkbook.workbook.write(bos);
                byte[] fileBytes = bos.toByteArray();

                clearCurrentWorkbook();

                return fileBytes;
            }
        } finally {
            lock.unlock();
        }
    }

    @Override
    public String peekDownloadFileName() {
        lock.lock();
        try {
            if (currentWorkbook == null || currentWorkbook.originalFileName == null || currentWorkbook.originalFileName.isBlank()) {
                return "edited_excel.xlsx";
            }
            return "edited_" + currentWorkbook.originalFileName;
        } finally {
            lock.unlock();
        }
    }

    private List<List<ExcelCellDto>> buildPreviewRows(Sheet sheet) {
        List<List<ExcelCellDto>> rows = new ArrayList<>();

        if (sheet == null) {
            return rows;
        }

        int maxColumnCount = getMaxColumnCount(sheet);
        if (maxColumnCount <= 0) {
            return rows;
        }

        DataFormatter formatter = new DataFormatter(Locale.KOREA);
        FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

        int lastMeaningfulRow = getLastMeaningfulRow(sheet, maxColumnCount, formatter, evaluator);
        if (lastMeaningfulRow < 0) {
            return rows;
        }

        // 0행부터 마지막 의미 있는 행까지 전부 보여줌
        for (int rowIndex = 0; rowIndex <= lastMeaningfulRow; rowIndex++) {
            List<ExcelCellDto> rowCells = new ArrayList<>();

            for (int colIndex = 0; colIndex < maxColumnCount; colIndex++) {
                CellRangeAddress mergedRegion = getMergedRegion(sheet, rowIndex, colIndex);

                if (mergedRegion != null) {
                    boolean isTopLeft = mergedRegion.getFirstRow() == rowIndex
                            && mergedRegion.getFirstColumn() == colIndex;

                    if (isTopLeft) {
                        Row topRow = sheet.getRow(mergedRegion.getFirstRow());
                        Cell realCell = topRow == null ? null
                                : topRow.getCell(mergedRegion.getFirstColumn(), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

                        int rowspan = mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1;
                        int colspan = mergedRegion.getLastColumn() - mergedRegion.getFirstColumn() + 1;

                        rowCells.add(new ExcelCellDto(
                                getDisplayValue(realCell, formatter, evaluator),
                                rowspan,
                                colspan,
                                false,
                                rowIndex,
                                colIndex,
                                isFormulaCell(realCell),
                                getReadableCellType(realCell)
                        ));
                    } else {
                        rowCells.add(new ExcelCellDto(
                                "",
                                1,
                                1,
                                true,
                                rowIndex,
                                colIndex,
                                false,
                                "HIDDEN"
                        ));
                    }
                } else {
                    Row row = sheet.getRow(rowIndex);
                    Cell cell = row == null ? null
                            : row.getCell(colIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

                    rowCells.add(new ExcelCellDto(
                            getDisplayValue(cell, formatter, evaluator),
                            1,
                            1,
                            false,
                            rowIndex,
                            colIndex,
                            isFormulaCell(cell),
                            getReadableCellType(cell)
                    ));
                }
            }

            rows.add(rowCells);
        }

        return rows;
    }
    private int getLastMeaningfulRow(Sheet sheet, int maxColumnCount,
	            DataFormatter formatter, FormulaEvaluator evaluator) {
		for (int rowIndex = sheet.getLastRowNum(); rowIndex >= 0; rowIndex--) {
			if (!isRowCompletelyEmpty(sheet, rowIndex, maxColumnCount, formatter, evaluator)) {
			return rowIndex;
			}
		}
		
	return -1;
	}

    private void applyChanges(Workbook workbook, List<ExcelChangeDto> changes) {
        if (changes == null || changes.isEmpty()) {
            return;
        }

        Sheet sheet = workbook.getSheetAt(0);

        for (ExcelChangeDto change : changes) {
            if (change == null) {
                continue;
            }

            int rowIndex = change.getRowIndex();
            int colIndex = change.getColIndex();
            String newValue = change.getValue() == null ? "" : change.getValue();

            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }

            Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

            // 원래 셀이 없던 경우: 새 셀 생성 후 일반 텍스트로 입력
            if (cell == null) {
                if (newValue.isBlank()) {
                    continue;
                }

                Cell newCell = row.createCell(colIndex, CellType.STRING);
                newCell.setCellValue(newValue);
                continue;
            }

            // 수식 셀은 수정 불가
            if (cell.getCellType() == CellType.FORMULA) {
                continue;
            }

            setCellValueKeepingOriginalStyleAndMeaning(cell, newValue);
        }

        workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
    }

    private void setCellValueKeepingOriginalStyleAndMeaning(Cell cell, String newValue) {
        if (newValue == null) {
            newValue = "";
        }

        CellType originalType = cell.getCellType();

        try {
            if (originalType == CellType.NUMERIC) {
                if (DateUtil.isCellDateFormatted(cell)) {
                    if (newValue.isBlank()) {
                        cell.setBlank();
                        return;
                    }

                    Date parsedDate = tryParseDate(newValue);
                    if (parsedDate == null) {
                        throw new IllegalArgumentException("날짜 셀에는 날짜 형식 값만 입력할 수 있습니다: " + newValue);
                    }

                    cell.setCellValue(parsedDate);
                    return;
                }

                if (newValue.isBlank()) {
                    cell.setBlank();
                    return;
                }

                String normalized = normalizeNumericText(newValue);
                double numericValue = Double.parseDouble(normalized);
                cell.setCellValue(numericValue);
                return;
            }

            if (originalType == CellType.BOOLEAN) {
                if (newValue.isBlank()) {
                    cell.setBlank();
                    return;
                }

                if (!"true".equalsIgnoreCase(newValue) && !"false".equalsIgnoreCase(newValue)) {
                    throw new IllegalArgumentException("BOOLEAN 셀에는 true 또는 false만 입력할 수 있습니다: " + newValue);
                }

                cell.setCellValue(Boolean.parseBoolean(newValue));
                return;
            }

            if (originalType == CellType.BLANK) {
                if (newValue.isBlank()) {
                    cell.setBlank();
                } else {
                    cell.setCellValue(newValue);
                }
                return;
            }

            if (newValue.isBlank()) {
                cell.setBlank();
            } else {
                cell.setCellValue(newValue);
            }
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException("숫자 셀에는 숫자만 입력할 수 있습니다: " + newValue);
        }
    }

    private String normalizeNumericText(String value) {
        return value.replace(",", "")
                .replace("₩", "")
                .replace("$", "")
                .replace("€", "")
                .replace("%", "")
                .trim();
    }

    private Date tryParseDate(String value) {
        List<String> patterns = Arrays.asList(
                "yyyy-MM-dd",
                "yyyy/MM/dd",
                "yyyy.MM.dd",
                "yyyy-MM-dd HH:mm:ss",
                "yyyy/MM/dd HH:mm:ss",
                "yyyy.MM.dd HH:mm:ss",
                "yyyy-MM-dd HH:mm",
                "yyyy/MM/dd HH:mm",
                "yyyy.MM.dd HH:mm",
                "HH:mm:ss",
                "HH:mm",
                "yy-MM-dd",
                "yy/MM/dd",
                "yy.MM.dd"
        );

        for (String pattern : patterns) {
            try {
                SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                sdf.setLenient(false);
                return sdf.parse(value);
            } catch (Exception ignored) {
            }
        }
        return null;
    }

    private boolean isFormulaCell(Cell cell) {
        return cell != null && cell.getCellType() == CellType.FORMULA;
    }

    private String getReadableCellType(Cell cell) {
        if (cell == null) {
            return "BLANK";
        }

        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) {
            return "FORMULA";
        }
        return type.name();
    }

    private String getDisplayValue(Cell cell, DataFormatter formatter, FormulaEvaluator evaluator) {
        if (cell == null) {
            return "";
        }

        try {
            CellStyle style = cell.getCellStyle();
            short formatIndex = style != null ? style.getDataFormat() : 0;
            String formatString = style != null ? style.getDataFormatString() : null;

            CellType type = cell.getCellType();

            // 수식 셀
            if (type == CellType.FORMULA) {
                CellValue evaluated = evaluator.evaluate(cell);
                if (evaluated == null) {
                    return "";
                }

                CellType resultType = evaluated.getCellType();

                if (resultType == CellType.NUMERIC) {
                    double value = evaluated.getNumberValue();

                    // 날짜/시간 서식 셀인 경우
                    if (DateUtil.isADateFormat(formatIndex, formatString)) {
                        Date date = DateUtil.getJavaDate(value);
                        String normalizedFormat = normalizeExcelDateFormat(formatString);

                        if (isTimeOnlyFormat(normalizedFormat)) {
                            return new java.text.SimpleDateFormat("HH:mm:ss").format(date);
                        }

                        if (hasTimePart(normalizedFormat)) {
                            return new java.text.SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(date);
                        }

                        return new java.text.SimpleDateFormat("yyyy-MM-dd").format(date);
                    }

                    return formatter.formatRawCellContents(value, formatIndex, formatString);
                }

                if (resultType == CellType.STRING) {
                    return evaluated.getStringValue();
                }

                if (resultType == CellType.BOOLEAN) {
                    return String.valueOf(evaluated.getBooleanValue());
                }

                return "";
            }

            // 일반 숫자 셀
            if (type == CellType.NUMERIC) {
                double value = cell.getNumericCellValue();

                // 날짜/시간 서식 셀인 경우
                if (DateUtil.isADateFormat(formatIndex, formatString)) {
                    Date date = DateUtil.getJavaDate(value);
                    String normalizedFormat = normalizeExcelDateFormat(formatString);

                    if (isTimeOnlyFormat(normalizedFormat)) {
                        return new java.text.SimpleDateFormat("HH:mm:ss").format(date);
                    }

                    if (hasTimePart(normalizedFormat)) {
                        return new java.text.SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(date);
                    }

                    return new java.text.SimpleDateFormat("yyyy-MM-dd").format(date);
                }

                return formatter.formatRawCellContents(value, formatIndex, formatString);
            }

            if (type == CellType.STRING) {
                return cell.getStringCellValue();
            }

            if (type == CellType.BOOLEAN) {
                return String.valueOf(cell.getBooleanCellValue());
            }

            if (type == CellType.BLANK) {
                return "";
            }

            return formatter.formatCellValue(cell, evaluator);
        } catch (Exception e) {
            return formatter.formatCellValue(cell, evaluator);
        }
    }

    private String normalizeExcelDateFormat(String formatString) {
        if (formatString == null) {
            return "";
        }

        String s = formatString.toLowerCase();

        // locale 태그 제거: [$-412], [$-ko-KR]
        s = s.replaceAll("\\[\\$-[^\\]]*\\]", "");
        s = s.replaceAll("\\[\\$[^\\]]*\\]", "");

        // 색상 태그 제거: [Red]
        s = s.replaceAll("\\[[a-zA-Z]+\\]", "");

        // escape 제거
        s = s.replace("\\", "");

        // AM/PM 제거
        s = s.replace("am/pm", "");

        return s.trim();
    }

    private boolean hasTimePart(String format) {
        return format.contains("h") || format.contains("s");
    }

    private boolean isTimeOnlyFormat(String format) {
        boolean hasTime = format.contains("h") || format.contains("s");
        boolean hasDate = format.contains("y") || format.contains("d");
        return hasTime && !hasDate;
    }

    private boolean isRowCompletelyEmpty(Sheet sheet, int rowIndex, int maxColumnCount,
                                         DataFormatter formatter, FormulaEvaluator evaluator) {
        for (int colIndex = 0; colIndex < maxColumnCount; colIndex++) {
            CellRangeAddress mergedRegion = getMergedRegion(sheet, rowIndex, colIndex);

            if (mergedRegion != null) {
                Row mergedStartRow = sheet.getRow(mergedRegion.getFirstRow());
                Cell mergedCell = mergedStartRow == null ? null :
                        mergedStartRow.getCell(mergedRegion.getFirstColumn(), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

                if (!getDisplayValue(mergedCell, formatter, evaluator).isBlank()) {
                    return false;
                }
            } else {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row == null ? null :
                        row.getCell(colIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

                if (!getDisplayValue(cell, formatter, evaluator).isBlank()) {
                    return false;
                }
            }
        }
        return true;
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

    private CellRangeAddress getMergedRegion(Sheet sheet, int rowIndex, int colIndex) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.isInRange(rowIndex, colIndex)) {
                return region;
            }
        }
        return null;
    }

    private Workbook createWorkbook(MultipartFile file, InputStream is) throws Exception {
        String fileName = file.getOriginalFilename();
        if (fileName == null) {
            throw new IllegalArgumentException("파일명이 없습니다.");
        }

        String lower = fileName.toLowerCase();
        if (lower.endsWith(".xlsx")) {
            return new XSSFWorkbook(is);
        }
        if (lower.endsWith(".xls")) {
            return new HSSFWorkbook(is);
        }

        throw new IllegalArgumentException("엑셀 파일만 업로드할 수 있습니다.");
    }

    private Workbook cloneWorkbook(Workbook source, String originalFileName) throws Exception {
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            source.write(bos);
            byte[] bytes = bos.toByteArray();

            if (originalFileName != null && originalFileName.toLowerCase().endsWith(".xls")) {
                return new HSSFWorkbook(new ByteArrayInputStream(bytes));
            }
            return new XSSFWorkbook(new ByteArrayInputStream(bytes));
        }
    }

    private void validateExcelFile(MultipartFile file) {
        if (file == null || file.isEmpty()) {
            throw new IllegalArgumentException("업로드할 파일이 없습니다.");
        }

        String fileName = file.getOriginalFilename();
        if (fileName == null || fileName.isBlank()) {
            throw new IllegalArgumentException("파일명이 없습니다.");
        }

        String lower = fileName.toLowerCase();
        if (!lower.endsWith(".xls") && !lower.endsWith(".xlsx")) {
            throw new IllegalArgumentException("엑셀 파일만 업로드할 수 있습니다.");
        }
    }

    private void clearCurrentWorkbook() throws Exception {
        if (currentWorkbook != null && currentWorkbook.workbook != null) {
            currentWorkbook.workbook.close();
        }
        currentWorkbook = null;
    }

    private static class StoredWorkbook {
        private final String token;
        private final String originalFileName;
        private final Workbook workbook;

        private StoredWorkbook(String token, String originalFileName, Workbook workbook) {
            this.token = token;
            this.originalFileName = originalFileName;
            this.workbook = workbook;
        }
    }
}