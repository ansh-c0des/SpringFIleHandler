package com.Truboard.ExcelFileDetector.service;

import com.Truboard.ExcelFileDetector.DTO.ColumnValidationRule;
import com.Truboard.ExcelFileDetector.DTO.ExcelInfoResponse;
import com.Truboard.ExcelFileDetector.config.ExcelValidationConfig;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

@Service
public class ExcelService {

    private final ExcelValidationConfig validationConfig;
    private final ObjectMapper objectMapper = new ObjectMapper();

    public ExcelService(ExcelValidationConfig validationConfig) {
        this.validationConfig = validationConfig;
    }

    /**
     * Existing Excel processing (unchanged behavior except refactored process step).
     */
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

            // Build columnData: Map<HeaderName, List<rowValues>>
            Map<String, List<String>> columnData = new LinkedHashMap<>();
            int maxColumns = headerRow.getLastCellNum();
            for (int colIndex = 0; colIndex < maxColumns; colIndex++) {
                Cell headerCell = headerRow.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String colName = headerCell.toString().trim();
                if (colName.isEmpty()) {
                    colName = "Column_" + (colIndex + 1);
                }

                List<String> colValues = new ArrayList<>();
                for (int rowIndex = 1; rowIndex <= sheetToRead.getLastRowNum(); rowIndex++) {
                    Row row = sheetToRead.getRow(rowIndex);
                    Cell cell = (row == null)
                            ? null
                            : row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String value = (cell == null) ? "" : getCellString(cell).trim();
                    colValues.add(value);
                }
                columnData.put(colName, colValues);
            }

            // Delegate to common validation/response builder
            return processDataAndValidate(columnData, sheetCount, sheetNames);
        }
    }

    /**
     * New: process a JSON file upload. Accepts two JSON shapes:
     * 1) Array of objects: [ { "Name": "John", "Age": 25 }, { ... } ]
     * 2) Object of arrays: { "Name": ["John","Alice"], "Age":[25,30] }
     */
    public ExcelInfoResponse extractJsonInfo(MultipartFile file) throws Exception {
        try (InputStream in = file.getInputStream()) {
            // Try array-of-objects first
            try {
                List<Map<String, Object>> rows = objectMapper.readValue(in, new TypeReference<List<Map<String, Object>>>() {});
                if (rows == null) rows = Collections.emptyList();

                // Build columnData from rows (columns discovered from keys)
                LinkedHashMap<String, List<String>> columnData = new LinkedHashMap<>();
                // preserve order: first collect keys in order of appearance across rows
                LinkedHashSet<String> keysOrder = new LinkedHashSet<>();
                for (Map<String, Object> row : rows) {
                    if (row == null) continue;
                    keysOrder.addAll(row.keySet());
                }
                for (String key : keysOrder) {
                    columnData.put(key, new ArrayList<>());
                }

                for (Map<String, Object> row : rows) {
                    for (String key : keysOrder) {
                        Object val = (row == null) ? null : row.get(key);
                        columnData.get(key).add(val == null ? "" : String.valueOf(val));
                    }
                }

                // sheetCount = 1, name "JSON"
                return processDataAndValidate(columnData, 1, Collections.singletonList("JSON"));
            } catch (Exception eArray) {
                // reset stream to try second format - need to reopen stream
            }
        }

        // second attempt: map of arrays
        try (InputStream in2 = file.getInputStream()) {
            Map<String, List<Object>> cols = objectMapper.readValue(in2, new TypeReference<Map<String, List<Object>>>() {});
            if (cols == null) cols = Collections.emptyMap();

            LinkedHashMap<String, List<String>> columnData = new LinkedHashMap<>();
            int maxRows = 0;
            for (Map.Entry<String, List<Object>> e : cols.entrySet()) {
                maxRows = Math.max(maxRows, (e.getValue() == null) ? 0 : e.getValue().size());
            }
            // Convert each col to List<String>
            for (Map.Entry<String, List<Object>> e : cols.entrySet()) {
                String key = e.getKey();
                List<Object> objList = e.getValue();
                List<String> stringList = new ArrayList<>();
                if (objList != null) {
                    for (Object o : objList) stringList.add(o == null ? "" : String.valueOf(o));
                }
                // pad shorter lists to maxRows so structure is rectangular
                while (stringList.size() < maxRows) stringList.add("");
                columnData.put(key, stringList);
            }

            return processDataAndValidate(columnData, 1, Collections.singletonList("JSON"));
        } catch (Exception ex) {
            throw new Exception("JSON parsing failed: " + ex.getMessage(), ex);
        }
    }

    /**
     * Shared logic: validate required columns and run per-cell validation.
     * Returns ExcelInfoResponse with sheetCount, sheetNames, sheetData and errors.
     */
    private ExcelInfoResponse processDataAndValidate(Map<String, List<String>> columnData,
                                                     int sheetCount,
                                                     List<String> sheetNames) {
        List<String> errors = new ArrayList<>();
        Map<String, ColumnValidationRule> rules = validationConfig.getValidations();
        List<String> requiredColsFromConfig = validationConfig.getRequiredColumns();

        // Build normalized header set from provided columnData keys
        Set<String> normalizedHeaders = new HashSet<>();
        for (String header : columnData.keySet()) {
            normalizedHeaders.add(normalizeForCompare(header));
        }

        // 1) Check required columns presence (configurable), normalize comparison
        if (requiredColsFromConfig != null) {
            for (String required : requiredColsFromConfig) {
                if (required == null || required.trim().isEmpty()) continue;
                String reqNorm = normalizeForCompare(required);
                if (!normalizedHeaders.contains(reqNorm)) {
                    errors.add("Missing required column: " + required);
                }
            }
        }

        // If missing required columns, return early (no further validation)
        if (!errors.isEmpty()) {
            return new ExcelInfoResponse(sheetCount, sheetNames, Collections.emptyMap(), errors);
        }

        // 2) Validate each column by rule (if rule exists)
        if (rules == null) rules = Collections.emptyMap();
        for (Map.Entry<String, List<String>> entry : columnData.entrySet()) {
            String colName = entry.getKey();
            List<String> values = entry.getValue();

            // propertyKey uses underscores as in application.properties
            String propertyKey = colName.replace(" ", "_");
            ColumnValidationRule rule = rules.get(propertyKey);

            if (rule != null) {
                for (int i = 0; i < values.size(); i++) {
                    String value = values.get(i);
                    // row number for messaging: i+2 if we think of header row 1 for excel; but for JSON we'll use i+1
                    int displayRowNum = (sheetCount == 1 && sheetNames.size() == 1 && "JSON".equals(sheetNames.get(0)))
                            ? (i + 1)
                            : (i + 2);
                    validateCell(value, rule, displayRowNum, colName, errors);
                }
            }
        }

        return new ExcelInfoResponse(sheetCount, sheetNames, columnData, errors);
    }

    // Normalization helper: convert underscores to spaces, trim, lower-case to compare flexibly
    private String normalizeForCompare(String s) {
        if (s == null) return "";
        return s.replace('_', ' ').trim().toLowerCase(Locale.ROOT);
    }

    // Convert cell to string, handling dates/numerics properly
    private String getCellString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // default formatting for numeric date cells
                    return new SimpleDateFormat("dd/MM/yyyy").format(cell.getDateCellValue());
                } else {
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
                // unknown type: ignore
                break;
        }
    }
}
