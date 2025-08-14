package com.Truboard.ExcelFileDetector.service;

import com.Truboard.ExcelFileDetector.DTO.ColumnValidationRule;
import com.Truboard.ExcelFileDetector.DTO.ExcelInfoResponse;
import com.Truboard.ExcelFileDetector.DTO.ValidationError;
import com.Truboard.ExcelFileDetector.config.ExcelValidationConfig;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

@Service
public class ExcelService {

    private final ExcelValidationConfig validationConfig;
    private final ObjectMapper objectMapper = new ObjectMapper();
    private final FileStorageService fileStorageService;

    public ExcelService(ExcelValidationConfig validationConfig, FileStorageService fileStorageService) {
        this.validationConfig = validationConfig;
        this.fileStorageService = fileStorageService;
    }

    /**
     * Process an uploaded .xlsx file: extract column-wise data and validate.
     */
    public ExcelInfoResponse extractExcelInfo(MultipartFile file) throws Exception {
        String fileId = fileStorageService.storeFile(file);
        
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
            Map<String, Integer> columnIndexMap = new LinkedHashMap<>(); // Track column positions
            
            int maxColumns = headerRow.getLastCellNum();
            for (int colIndex = 0; colIndex < maxColumns; colIndex++) {
                Cell headerCell = headerRow.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String colName = headerCell.toString().trim();
                if (colName.isEmpty()) {
                    colName = "Column_" + (colIndex + 1);
                }

                columnIndexMap.put(colName, colIndex);
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
            ExcelInfoResponse response = processDataAndValidate(columnData, sheetCount, sheetNames, columnIndexMap, "xlsx");
            response.setFileId(fileId);
            response.setFileType("xlsx");
            return response;
        }
    }

    /**
     * Process an uploaded .json file. Accepts either:
     * - array of objects: [ {"Name":"John","Age":25}, ... ]
     * - object of arrays: { "Name": ["John","Alice"], "Age":[25,30] }
     */
    public ExcelInfoResponse extractJsonInfo(MultipartFile file) throws Exception {
        String fileId = fileStorageService.storeFile(file);
        
        // Try JSON array-of-objects first
        try (InputStream in = file.getInputStream()) {
            try {
                List<Map<String, Object>> rows = objectMapper.readValue(in, new TypeReference<List<Map<String, Object>>>() {});
                if (rows == null) rows = Collections.emptyList();

                LinkedHashMap<String, List<String>> columnData = new LinkedHashMap<>();
                LinkedHashMap<String, Integer> columnIndexMap = new LinkedHashMap<>();
                LinkedHashSet<String> keysOrder = new LinkedHashSet<>();
                
                for (Map<String, Object> row : rows) {
                    if (row == null) continue;
                    keysOrder.addAll(row.keySet());
                }
                
                int colIndex = 0;
                for (String key : keysOrder) {
                    columnData.put(key, new ArrayList<>());
                    columnIndexMap.put(key, colIndex++);
                }

                for (Map<String, Object> row : rows) {
                    for (String key : keysOrder) {
                        Object val = (row == null) ? null : row.get(key);
                        columnData.get(key).add(val == null ? "" : String.valueOf(val));
                    }
                }

                ExcelInfoResponse response = processDataAndValidate(columnData, 1, Collections.singletonList("JSON"), columnIndexMap, "json");
                response.setFileId(fileId);
                response.setFileType("json");
                return response;
            } catch (Exception eArray) {
                // fall through to try object-of-arrays
            }
        }

        // Try object-of-arrays
        try (InputStream in2 = file.getInputStream()) {
            Map<String, List<Object>> cols = objectMapper.readValue(in2, new TypeReference<Map<String, List<Object>>>() {});
            if (cols == null) cols = Collections.emptyMap();

            LinkedHashMap<String, List<String>> columnData = new LinkedHashMap<>();
            LinkedHashMap<String, Integer> columnIndexMap = new LinkedHashMap<>();
            
            int maxRows = 0;
            for (Map.Entry<String, List<Object>> e : cols.entrySet()) {
                maxRows = Math.max(maxRows, (e.getValue() == null) ? 0 : e.getValue().size());
            }
            
            int colIndex = 0;
            for (Map.Entry<String, List<Object>> e : cols.entrySet()) {
                String key = e.getKey();
                columnIndexMap.put(key, colIndex++);
                List<Object> objList = e.getValue();
                List<String> stringList = new ArrayList<>();
                if (objList != null) {
                    for (Object o : objList) stringList.add(o == null ? "" : String.valueOf(o));
                }
                while (stringList.size() < maxRows) stringList.add("");
                columnData.put(key, stringList);
            }

            ExcelInfoResponse response = processDataAndValidate(columnData, 1, Collections.singletonList("JSON"), columnIndexMap, "json");
            response.setFileId(fileId);
            response.setFileType("json");
            return response;
        } catch (Exception ex) {
            throw new Exception("JSON parsing failed: " + ex.getMessage(), ex);
        }
    }

    /**
     * Generate Excel file with validation errors highlighted
     */
    public byte[] generateErrorHighlightedExcel(String fileId) throws Exception {
        if (!fileStorageService.fileExists(fileId)) {
            throw new Exception("File not found or expired");
        }

        byte[] originalContent = fileStorageService.getFileContent(fileId);
        String fileName = fileStorageService.getFileName(fileId);
        
        if (fileName == null || !fileName.toLowerCase().endsWith(".xlsx")) {
            throw new Exception("Error highlighting is only supported for Excel files");
        }

        // Re-process the file to get validation errors
        try (InputStream inputStream = new ByteArrayInputStream(originalContent);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            int sheetCount = workbook.getNumberOfSheets();
            Sheet sheetToRead = (sheetCount == 1) ? workbook.getSheetAt(0) : workbook.getSheet("Data");
            
            if (sheetToRead == null) {
                throw new Exception("Sheet named 'Data' not found in the workbook");
            }

            Row headerRow = sheetToRead.getRow(0);
            if (headerRow == null) {
                throw new Exception("No header row found in sheet");
            }

            // Build column data and get validation errors
            Map<String, List<String>> columnData = new LinkedHashMap<>();
            Map<String, Integer> columnIndexMap = new LinkedHashMap<>();
            
            int maxColumns = headerRow.getLastCellNum();
            for (int colIndex = 0; colIndex < maxColumns; colIndex++) {
                Cell headerCell = headerRow.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String colName = headerCell.toString().trim();
                if (colName.isEmpty()) {
                    colName = "Column_" + (colIndex + 1);
                }

                columnIndexMap.put(colName, colIndex);
                List<String> colValues = new ArrayList<>();
                for (int rowIndex = 1; rowIndex <= sheetToRead.getLastRowNum(); rowIndex++) {
                    Row row = sheetToRead.getRow(rowIndex);
                    Cell cell = (row == null) ? null : row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String value = (cell == null) ? "" : getCellString(cell).trim();
                    colValues.add(value);
                }
                columnData.put(colName, colValues);
            }

            // Get detailed validation errors
            List<ValidationError> detailedErrors = getDetailedValidationErrors(columnData, columnIndexMap);
            
            // Apply highlighting and comments
            applyErrorHighlighting(workbook, sheetToRead, detailedErrors);
            
            // Write modified workbook to byte array
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    /**
     * Apply error highlighting and comments to Excel cells
     */
    private void applyErrorHighlighting(Workbook workbook, Sheet sheet, List<ValidationError> errors) {
        // Create red background style
        CellStyle errorStyle = workbook.createCellStyle();
        errorStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        errorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        // Create drawing patriarch for comments
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        CreationHelper creationHelper = workbook.getCreationHelper();
        
        for (ValidationError error : errors) {
            // CRITICAL FIX: Ensure valid row and column indices
            int colIndex = error.getColumnIndex();
            // A rowNumber of 0 indicates a missing column error, which cannot be highlighted on a cell
            if (error.getRowNumber() <= 0 || colIndex < 0) {
                continue; 
            }
            int rowIndex = error.getRowNumber() - 1; // Convert to 0-based index
            
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;
            
            // CRITICAL FIX: Get cell, but create it if it doesn't exist
            Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

            // Apply red background
            cell.setCellStyle(errorStyle);
            
            // Add comment with error message
            ClientAnchor anchor = creationHelper.createClientAnchor();
            anchor.setCol1(colIndex);
            anchor.setCol2(colIndex + 3);
            anchor.setRow1(rowIndex);
            anchor.setRow2(rowIndex + 3);
            
            Comment comment = drawing.createCellComment(anchor);
            RichTextString richTextString = creationHelper.createRichTextString(
                "Validation Error:\n" + error.getMessage() + "\nCurrent value: " + error.getCellValue()
            );
            comment.setString(richTextString);
            comment.setAuthor("Excel Validator");
            cell.setCellComment(comment);
        }
    }

    /**
     * Get detailed validation errors with cell positions
     */
    private List<ValidationError> getDetailedValidationErrors(Map<String, List<String>> columnData, 
                                                              Map<String, Integer> columnIndexMap) {
        List<ValidationError> detailedErrors = new ArrayList<>();
        Map<String, ColumnValidationRule> rules = validationConfig.getValidations();
        
        if (rules == null) rules = Collections.emptyMap();
        
        for (Map.Entry<String, List<String>> entry : columnData.entrySet()) {
            String colName = entry.getKey();
            List<String> values = entry.getValue();
            Integer colIndex = columnIndexMap.get(colName);
            
            if (colIndex == null) colIndex = -1;
            
            String propertyKey = colName.replace(" ", "_");
            ColumnValidationRule rule = rules.get(propertyKey);
            
            if (rule != null) {
                for (int i = 0; i < values.size(); i++) {
                    String value = values.get(i);
                    int displayRowNum = i + 2; // Excel row number (1-based + header)
                    
                    List<String> cellErrors = new ArrayList<>();
                    validateCellDetailed(value, rule, cellErrors);
                    
                    for (String errorMsg : cellErrors) {
                        detailedErrors.add(new ValidationError(
                            colName, 
                            displayRowNum, 
                            colIndex, 
                            "Row " + displayRowNum + ": " + colName + " " + errorMsg,
                            value
                        ));
                    }
                }
            }
        }
        
        return detailedErrors;
    }

    /**
     * Shared logic: validate required columns and run per-cell validation.
     * Returns ExcelInfoResponse with sheetCount, sheetNames, sheetData and errors.
     */
    private ExcelInfoResponse processDataAndValidate(Map<String, List<String>> columnData,
                                                     int sheetCount,
                                                     List<String> sheetNames,
                                                     Map<String, Integer> columnIndexMap,
                                                     String fileType) {
        List<String> errors = new ArrayList<>();
        List<ValidationError> detailedErrors = new ArrayList<>();
        Map<String, ColumnValidationRule> rules = validationConfig.getValidations();
        List<String> requiredColsFromConfig = validationConfig.getRequiredColumns();

        // Build normalized header set from provided columnData keys
        Set<String> normalizedHeaders = new HashSet<>();
        for (String header : columnData.keySet()) {
            normalizedHeaders.add(normalizeForCompare(header));
        }

        // 1) Check required columns presence (configurable)
        if (requiredColsFromConfig != null) {
            for (String required : requiredColsFromConfig) {
                if (required == null || required.trim().isEmpty()) continue;
                String reqNorm = normalizeForCompare(required);
                if (!normalizedHeaders.contains(reqNorm)) {
                    String missingError = "Missing required column: " + required;
                    errors.add(missingError);
                    // Add a detailed error for the missing column itself
                    detailedErrors.add(new ValidationError(required, 0, -1, missingError, null));
                }
            }
        }

        // 2) Validate each column by rule (if rule exists)
        if (rules == null) rules = Collections.emptyMap();
        for (Map.Entry<String, List<String>> entry : columnData.entrySet()) {
            String colName = entry.getKey();
            List<String> values = entry.getValue();
            Integer colIndex = columnIndexMap.get(colName);

            if (colIndex == null) colIndex = -1;

            // propertyKey uses underscores as in application.properties: "Joining_Date"
            String propertyKey = colName.replace(" ", "_");
            ColumnValidationRule rule = rules.get(propertyKey);

            if (rule != null) {
                for (int i = 0; i < values.size(); i++) {
                    String value = values.get(i);
                    int displayRowNum = ("json".equals(fileType)) ? (i + 1) : (i + 2);
                    
                    List<String> cellErrors = new ArrayList<>();
                    validateCellDetailed(value, rule, cellErrors);
                    
                    for (String errorMsg : cellErrors) {
                        String fullErrorMsg = "Row " + displayRowNum + ": " + colName + " " + errorMsg;
                        errors.add(fullErrorMsg);
                        detailedErrors.add(new ValidationError(colName, displayRowNum, colIndex, fullErrorMsg, value));
                    }
                }
            }
        }

        return new ExcelInfoResponse(sheetCount, sheetNames, columnData, errors, detailedErrors, null, fileType);
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

    /**
     * Enhanced cell validation that collects all errors for a single cell
     */
    private void validateCellDetailed(String value, ColumnValidationRule rule, List<String> errors) {
        // Required check
        if (rule.isRequired() && (value == null || value.isEmpty())) {
            errors.add("is required");
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
                        errors.add("must be >= " + rule.getMin());
                    }
                    if (rule.getMax() != null && num > rule.getMax()) {
                        errors.add("must be <= " + rule.getMax());
                    }
                } catch (NumberFormatException e) {
                    errors.add("must be a number");
                }
                break;

            case "date":
                try {
                    SimpleDateFormat sdf = new SimpleDateFormat(rule.getFormat());
                    sdf.setLenient(false);
                    sdf.parse(value);
                } catch (Exception e) {
                    errors.add("must match date format " + rule.getFormat());
                }
                break;

            case "text":
                if (rule.getRegex() != null && !value.matches(rule.getRegex())) {
                    errors.add("format is invalid");
                }
                break;

            default:
                break;
        }
    }
}