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

            // Pick correct sheet (Data if >1, otherwise first)
            Sheet sheetToRead = (sheetCount == 1)
                    ? workbook.getSheetAt(0)
                    : workbook.getSheet("Data");
            if (sheetToRead == null) {
                throw new Exception("Sheet named 'Data' not found in the workbook");
            }

            // Header row
            Row headerRow = sheetToRead.getRow(0);
            if (headerRow == null) {
                throw new Exception("No header row found in sheet");
            }

            Map<String, List<String>> columnData = new LinkedHashMap<>();
            List<String> errors = new ArrayList<>();

            // Load rules and required columns from config
            Map<String, ColumnValidationRule> rules = validationConfig.getValidations();
            List<String> requiredColsFromConfig = validationConfig.getRequiredColumns();

            // Build normalized header set for lookup (normalize: replace '_' with ' ', trim, lowercase)
            Set<String> normalizedHeaders = new HashSet<>();
            int headerCellCount = headerRow.getLastCellNum();
            for (int ci = 0; ci < headerCellCount; ci++) {
                Cell hc = headerRow.getCell(ci, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String headerName = hc.toString().trim();
                String normalized = normalizeForCompare(headerName);
                normalizedHeaders.add(normalized);
            }

            // 1) Check required columns presence (configurable)
            if (requiredColsFromConfig != null) {
                for (String required : requiredColsFromConfig) {
                    if (required == null || required.trim().isEmpty()) continue;
                    String reqNorm = normalizeForCompare(required);
                    if (!normalizedHeaders.contains(reqNorm)) {
                        // keep original text from config in message for clarity
                        errors.add("Missing required column: " + required);
                    }
                }
            }

            // If required columns missing, return early (no further data reading)
            if (!errors.isEmpty()) {
                return new ExcelInfoResponse(sheetCount, sheetNames, columnData, errors);
            }

            // 2) Read columns and validate rows (lookup rules by header -> propertyKey with underscores)
            int maxColumns = headerRow.getLastCellNum();
            for (int colIndex = 0; colIndex < maxColumns; colIndex++) {
                Cell headerCell = headerRow.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String colName = headerCell.toString().trim();
                if (colName.isEmpty()) {
                    colName = "Column_" + (colIndex + 1);
                }

                List<String> colValues = new ArrayList<>();
                // property key uses underscores as in application.properties: "Joining_Date"
                String propertyKey = colName.replace(" ", "_");
                ColumnValidationRule rule = (rules == null) ? null : rules.get(propertyKey);

                for (int rowIndex = 1; rowIndex <= sheetToRead.getLastRowNum(); rowIndex++) {
                    Row row = sheetToRead.getRow(rowIndex);
                    Cell cell = (row == null)
                            ? null
                            : row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                    String value = (cell == null) ? "" : getCellString(cell).trim();
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

    // Normalization helper: convert underscores to spaces, trim, lower-case to compare flexibly
    private String normalizeForCompare(String s) {
        if (s == null) return "";
        return s.replace('_', ' ').trim().toLowerCase();
    }

    // Convert cell to string, handling dates/numerics properly
    private String getCellString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // default formatting for numeric date cells; doesn't validate format here
                    return new SimpleDateFormat("dd/MM/yyyy").format(cell.getDateCellValue());
                } else {
                    // remove trailing .0 for integer-like values
                    double d = cell.getNumericCellValue();
                    if (d == Math.floor(d)) {
                        return String.valueOf((long) d);
                    } else {
                        return String.valueOf(d);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                // evaluate formula result as string where possible
                try {
                    FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                    CellValue evaluated = evaluator.evaluate(cell);
                    if (evaluated == null) return "";
                    switch (evaluated.getCellType()) {
                        case STRING: return evaluated.getStringValue();
                        case NUMERIC:
                            double dn = evaluated.getNumberValue();
                            if (dn == Math.floor(dn)) return String.valueOf((long) dn);
                            return String.valueOf(dn);
                        case BOOLEAN: return String.valueOf(evaluated.getBooleanValue());
                        default: return "";
                    }
                } catch (Exception e) {
                    return cell.getCellFormula();
                }
            case BLANK:
            default:
                return "";
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

        String type = (rule.getType() == null) ? "" : rule.getType().toLowerCase(Locale.ROOT);
        switch (type) {
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
                    SimpleDateFormat sdf = new SimpleDateFormat(rule.getFormat());
                    sdf.setLenient(false);
                    sdf.parse(value);
                } catch (Exception e) {
                    errors.add("Row " + rowNum + ": " + colName + " must match date format " + rule.getFormat());
                }
                break;

            case "text":
                if (rule.getRegex() != null && !value.matches(rule.getRegex())) {
                    errors.add("Row " + rowNum + ": " + colName + " format is invalid");
                }
                break;

            default:
                break;
        }
    }
}
