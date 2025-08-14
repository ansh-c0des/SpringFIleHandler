package com.Truboard.ExcelFileDetector.service;

import com.Truboard.ExcelFileDetector.DTO.ColumnValidationRule;
import com.Truboard.ExcelFileDetector.DTO.ExcelInfoResponse;
import com.Truboard.ExcelFileDetector.config.ExcelValidationConfig;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

@Service
public class ExcelService {

    private final ExcelValidationConfig validationConfig;

    public ExcelService(ExcelValidationConfig validationConfig) {
        this.validationConfig = validationConfig;
    }

    public ExcelInfoResponse extractExcelInfo(MultipartFile file) throws Exception {
        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            int sheetCount = workbook.getNumberOfSheets();
            List<String> sheetNames = new ArrayList<>();
            for (int i = 0; i < sheetCount; i++) {
                sheetNames.add(workbook.getSheetName(i));
            }

            Sheet sheetToRead = (sheetCount == 1)
                    ? workbook.getSheetAt(0)
                    : workbook.getSheet("Data");
            if (sheetToRead == null) {
                throw new Exception("Sheet named 'Data' not found in the workbook");
            }

            Row headerRow = sheetToRead.getRow(0);
            if (headerRow == null) {
                throw new Exception("No header row found in sheet");
            }

            Map<String, List<String>> columnData = new LinkedHashMap<>();
            List<String> errors = new ArrayList<>();
            Map<String, ColumnValidationRule> rules = validationConfig.getValidations();

            // Collect actual header names
            Set<String> actualHeaders = new HashSet<>();
            for (Cell cell : headerRow) {
                actualHeaders.add(cell.toString().trim());
            }

            // 1️⃣ Check for missing required columns
            for (Map.Entry<String, ColumnValidationRule> entry : rules.entrySet()) {
                String expectedCol = entry.getKey().replace("_", " "); // match Excel header style
                ColumnValidationRule rule = entry.getValue();

                if (rule.isRequired() && !actualHeaders.contains(expectedCol)) {
                    errors.add("Missing required column: " + expectedCol);
                }
            }

            // If required column(s) missing, return immediately
            if (!errors.isEmpty()) {
                return new ExcelInfoResponse(sheetCount, sheetNames, columnData, errors);
            }

            // 2️⃣ Read and validate each column
            int maxColumns = headerRow.getLastCellNum();
            for (int colIndex = 0; colIndex < maxColumns; colIndex++) {
                Cell headerCell = headerRow.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String colName = headerCell.toString().trim();
                if (colName.isEmpty()) {
                    colName = "Column_" + (colIndex + 1);
                }

                List<String> colValues = new ArrayList<>();
                String propertyKey = colName.replace(" ", "_");
                ColumnValidationRule rule = rules.get(propertyKey);

                for (int rowIndex = 1; rowIndex <= sheetToRead.getLastRowNum(); rowIndex++) {
                    Row row = sheetToRead.getRow(rowIndex);
                    Cell cell = (row == null)
                            ? null
                            : row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                    String value = (cell == null) ? "" : cell.toString().trim();
                    colValues.add(value);

                    if (rule != null) {
                        validateCell(value, rule, rowIndex + 1, colName, errors);
                    }
                }

                columnData.put(colName, colValues);
            }

            return new ExcelInfoResponse(sheetCount, sheetNames, columnData, errors);
        }
    }

    private void validateCell(String value, ColumnValidationRule rule, int rowNum,
                               String colName, List<String> errors) {
        // Required check
        if (rule.isRequired() && (value == null || value.isEmpty())) {
            errors.add("Row " + rowNum + ": " + colName + " is required");
            return;
        }

        // Skip further checks if empty and not required
        if (value == null || value.isEmpty()) return;

        switch (rule.getType().toLowerCase()) {
            case "number":
                try {
                    double num = Double.parseDouble(value);
                    if (rule.getMin() != null && num < rule.getMin()) {
                        errors.add("Row " + rowNum + ": " + colName + " must be >= " + rule.getMin());
                    }
                    if (rule.getMax() != null && num > rule.getMax()) {
                        errors.add("Row " + rowNum + ": " + colName + " must be <= " + rule.getMax());
                    }
                } catch (NumberFormatException e) {
                    errors.add("Row " + rowNum + ": " + colName + " must be a number");
                }
                break;

            case "date":
                try {
                    new SimpleDateFormat(rule.getFormat()).parse(value);
                } catch (ParseException e) {
                    errors.add("Row " + rowNum + ": " + colName + " must match date format " + rule.getFormat());
                }
                break;

            case "text":
                if (rule.getRegex() != null && !value.matches(rule.getRegex())) {
                    errors.add("Row " + rowNum + ": " + colName + " format is invalid");
                }
                break;
        }
    }
}
