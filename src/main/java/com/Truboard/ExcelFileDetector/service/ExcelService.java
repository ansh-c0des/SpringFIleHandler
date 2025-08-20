package com.Truboard.ExcelFileDetector.service;

import com.Truboard.ExcelFileDetector.DTO.ColumnValidationRule;
import com.Truboard.ExcelFileDetector.DTO.ExcelInfoResponse;
import com.Truboard.ExcelFileDetector.DTO.ValidationError;
import com.Truboard.ExcelFileDetector.config.ExcelValidationConfig;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * ExcelService - validation + highlighting logic with robust header-rule lookup,
 * rule-aware date formatting, and correct display-aware cell reading using
 * DataFormatter + FormulaEvaluator so percent signs and displayed date/currency
 * formats are preserved for validation.
 */
@Service
public class ExcelService {

    private final ExcelValidationConfig validationConfig;
    private final ObjectMapper objectMapper = new ObjectMapper();
    private final FileStorageService fileStorageService;

    // Normalized map for rule lookup: normalizedHeader -> ColumnValidationRule
    private final Map<String, ColumnValidationRule> normalizedRules = new HashMap<>();

    public ExcelService(ExcelValidationConfig validationConfig, FileStorageService fileStorageService) {
        this.validationConfig = validationConfig;
        this.fileStorageService = fileStorageService;

        // Build normalized rules map for robust lookup (normalize keys like "Joining_Date" -> "joining date")
        Map<String, ColumnValidationRule> rules = validationConfig.getValidations();
        if (rules != null) {
            for (Map.Entry<String, ColumnValidationRule> e : rules.entrySet()) {
                String configuredKey = e.getKey(); // e.g. "ACQUISITION_DATE" or "PENAL_RATE"
                ColumnValidationRule rule = e.getValue();

                // Store multiple normalized versions of the same rule
                String norm1 = normalizeForCompare(configuredKey); // normalized
                String norm2 = normalizeForCompare(configuredKey.replace('_', ' ')); // spaces
                String norm3 = normalizeForCompare(configuredKey).replace(' ', '_'); // underscores

                normalizedRules.put(norm1, rule);
                normalizedRules.put(norm2, rule);
                normalizedRules.put(norm3, rule);

                System.out.println("Registered rule '" + configuredKey + "' with variants: " +
                        norm1 + ", " + norm2 + ", " + norm3);
            }
            System.out.println("Total normalized rules registered: " + normalizedRules.size());
        }
    }

    /**
     * Process an uploaded .xlsx file: extract column-wise data and validate.
     */
    public ExcelInfoResponse extractExcelInfo(MultipartFile file) throws Exception {
        String fileId = fileStorageService.storeFile(file);

        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            DataFormatter dataFormatter = new DataFormatter();

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

            // Determine last meaningful data row (ignore trailing empty rows)
            int lastDataRow = findLastNonEmptyRow(sheetToRead, maxColumns, dataFormatter, evaluator);
            System.out.println("Determined last meaningful data row: " + lastDataRow);

            for (int colIndex = 0; colIndex < maxColumns; colIndex++) {
                Cell headerCell = headerRow.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String colName = headerCell.toString().trim();
                if (colName.isEmpty()) {
                    colName = "Column_" + (colIndex + 1);
                }

                columnIndexMap.put(colName, colIndex);
                List<String> colValues = new ArrayList<>();

                // Determine rule for this column (normalized lookup)
                ColumnValidationRule ruleForColumn = getRuleForColumnName(colName);
                System.out.println("Column '" + colName + "' -> Rule: " +
                        (ruleForColumn != null ? ruleForColumn.getType() + " (required: " + ruleForColumn.isRequired() + ")" : "NONE"));

                // Only iterate up to lastDataRow (ignore trailing empty rows)
                for (int rowIndex = 1; rowIndex <= lastDataRow; rowIndex++) {
                    Row row = sheetToRead.getRow(rowIndex);
                    Cell cell = (row == null)
                            ? null
                            : row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                    String value;
                    if (cell == null) {
                        value = "";
                    } else {
                        // For validation purposes, always get the displayed value using DataFormatter + evaluator
                        value = dataFormatter.formatCellValue(cell, evaluator);
                    }

                    colValues.add(value == null ? "" : value.trim());
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

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            DataFormatter dataFormatter = new DataFormatter();

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

            // Determine last meaningful data row (ignore trailing empty rows)
            int lastDataRow = findLastNonEmptyRow(sheetToRead, maxColumns, dataFormatter, evaluator);
            System.out.println("generateErrorHighlightedExcel - lastDataRow: " + lastDataRow);

            for (int colIndex = 0; colIndex < maxColumns; colIndex++) {
                Cell headerCell = headerRow.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String colName = headerCell.toString().trim();
                if (colName.isEmpty()) {
                    colName = "Column_" + (colIndex + 1);
                }

                columnIndexMap.put(colName, colIndex);
                List<String> colValues = new ArrayList<>();

                for (int rowIndex = 1; rowIndex <= lastDataRow; rowIndex++) {
                    Row row = sheetToRead.getRow(rowIndex);
                    Cell cell = (row == null) ? null : row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String value;
                    if (cell == null) {
                        value = "";
                    } else {
                        // For validation, always use the displayed value without special formatting
                        value = dataFormatter.formatCellValue(cell, evaluator);
                    }
                    colValues.add(value == null ? "" : value.trim());
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
     * Apply error highlighting and comments to Excel cells - FIXED VERSION
     */
    private void applyErrorHighlighting(Workbook workbook, Sheet sheet, List<ValidationError> errors) {
        if (errors == null || errors.isEmpty()) return;

        // Create ONE red background style and reuse it
        CellStyle errorStyle = workbook.createCellStyle();
        errorStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        errorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Create drawing patriarch for comments
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        CreationHelper creationHelper = workbook.getCreationHelper();

        System.out.println("Applying highlighting for " + errors.size() + " errors");

        for (ValidationError error : errors) {
            int colIndex = error.getColumnIndex();
            // A rowNumber of 0 indicates a missing column error, which cannot be highlighted on a cell
            if (error.getRowNumber() <= 0 || colIndex < 0) {
                System.out.println("Skipping error for missing column or invalid position: " + error.getMessage());
                continue;
            }

            int rowIndex = error.getRowNumber(); // 1-based Excel row number
            int zeroBasedRowIndex = rowIndex - 1;

            System.out.println("Processing error: Column=" + colIndex + ", Row=" + rowIndex + " (0-based: " + zeroBasedRowIndex + "), Message=" + error.getMessage());

            Row row = sheet.getRow(zeroBasedRowIndex);
            if (row == null) {
                System.out.println("Row " + zeroBasedRowIndex + " is null, creating it");
                row = sheet.createRow(zeroBasedRowIndex);
            }

            // Get or create cell
            Cell cell = row.getCell(colIndex);
            if (cell == null) {
                System.out.println("Cell at column " + colIndex + " is null, creating it");
                cell = row.createCell(colIndex);
            }

            try {
                // Clone the existing cell style and add red background
                CellStyle newStyle = workbook.createCellStyle();
                if (cell.getCellStyle() != null) {
                    newStyle.cloneStyleFrom(cell.getCellStyle());
                }
                newStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                // Apply the style
                cell.setCellStyle(newStyle);
                System.out.println("Applied red background to cell at column " + colIndex + ", row " + zeroBasedRowIndex);

                // Add comment with error message
                ClientAnchor anchor = creationHelper.createClientAnchor();
                anchor.setCol1(colIndex);
                anchor.setCol2(colIndex + 3);
                anchor.setRow1(zeroBasedRowIndex);
                anchor.setRow2(zeroBasedRowIndex + 3);

                Comment comment = drawing.createCellComment(anchor);
                RichTextString richTextString = creationHelper.createRichTextString(
                        "Validation Error:\n" + error.getMessage() +
                                "\nCurrent value: " + (error.getCellValue() == null ? "" : error.getCellValue())
                );
                comment.setString(richTextString);
                comment.setAuthor("Excel Validator");
                cell.setCellComment(comment);
                System.out.println("Added comment to cell");

            } catch (Exception e) {
                System.err.println("Error applying formatting to cell at column " + colIndex + ", row " + zeroBasedRowIndex + ": " + e.getMessage());
                e.printStackTrace();
            }
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

        System.out.println("Getting detailed validation errors for " + columnData.size() + " columns");

        for (Map.Entry<String, List<String>> entry : columnData.entrySet()) {
            String colName = entry.getKey();
            List<String> values = entry.getValue();
            Integer colIndex = columnIndexMap.get(colName);

            if (colIndex == null) colIndex = -1;

            // Use normalized lookup
            ColumnValidationRule rule = getRuleForColumnName(colName);
            System.out.println("Column: " + colName + ", Rule: " + (rule != null ? rule.getType() : "none"));

            if (rule != null) {
                for (int i = 0; i < values.size(); i++) {
                    String value = values.get(i);
                    int displayRowNum = i + 2; // Excel row number (1-based + header)

                    List<String> cellErrors = new ArrayList<>();
                    validateCellDetailed(value, rule, cellErrors);

                    if (!cellErrors.isEmpty()) {
                        System.out.println("Found errors in column " + colName + ", row " + displayRowNum + ", value: '" + value + "'");
                        System.out.println("Errors: " + cellErrors);
                    }

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

        System.out.println("Total detailed errors found: " + detailedErrors.size());
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

            // Use normalized rule lookup
            ColumnValidationRule rule = getRuleForColumnName(colName);

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

    /**
     * Return the ColumnValidationRule for a column header using enhanced lookup.
     */
    private ColumnValidationRule getRuleForColumnName(String columnHeader) {
        return findRuleForColumn(columnHeader);
    }

    /**
     * Normalization for header/rule matching:
     * - replace underscores with spaces and spaces with underscores for bidirectional matching
     * - remove punctuation (non alnum/space/underscore)
     * - collapse multiple spaces and lowercase
     */
    private String normalizeForCompare(String s) {
        if (s == null) return "";
        String cleaned = s.replaceAll("[^A-Za-z0-9 _ ]+", "")  // Keep only alphanumeric, spaces, and underscores
                .trim()
                .replaceAll("\\s{2,}", " ")  // Collapse multiple spaces
                .toLowerCase(Locale.ROOT);
        return cleaned;
    }

    /**
     * Enhanced rule lookup that tries multiple normalization strategies
     */
    private ColumnValidationRule findRuleForColumn(String columnHeader) {
        if (columnHeader == null) return null;

        String normalized = normalizeForCompare(columnHeader);

        // Strategy 1: Direct normalized lookup
        ColumnValidationRule rule = normalizedRules.get(normalized);
        if (rule != null) {
            System.out.println("Found rule for '" + columnHeader + "' using direct lookup: " + normalized);
            return rule;
        }

        // Strategy 2: Try with spaces replaced by underscores
        String withUnderscores = normalized.replace(' ', '_');
        rule = normalizedRules.get(withUnderscores);
        if (rule != null) {
            System.out.println("Found rule for '" + columnHeader + "' using underscore replacement: " + withUnderscores);
            return rule;
        }

        // Strategy 3: Try with underscores replaced by spaces
        String withSpaces = normalized.replace('_', ' ');
        rule = normalizedRules.get(withSpaces);
        if (rule != null) {
            System.out.println("Found rule for '" + columnHeader + "' using space replacement: " + withSpaces);
            return rule;
        }

        // Strategy 4: Try exact match with original keys (case insensitive)
        for (Map.Entry<String, ColumnValidationRule> entry : normalizedRules.entrySet()) {
            if (entry.getKey().equalsIgnoreCase(columnHeader)) {
                System.out.println("Found rule for '" + columnHeader + "' using case-insensitive exact match");
                return entry.getValue();
            }
        }

        System.out.println("No rule found for column: '" + columnHeader + "' (normalized: '" + normalized + "')");
        System.out.println("Available rules: " + normalizedRules.keySet());
        return null;
    }

    /**
     * Helper that finds the last non-empty row index (1-based row index values) for data.
     * Returns 0 if no data rows found (i.e., only header exists).
     *
     * We inspect displayed cell text using DataFormatter + FormulaEvaluator to match what the user sees.
     */
    private int findLastNonEmptyRow(Sheet sheet, int maxColumns, DataFormatter formatter, FormulaEvaluator evaluator) {
        int lastRowNum = sheet.getLastRowNum();
        for (int r = lastRowNum; r >= 1; r--) { // start from bottom, ignore header row (0)
            Row row = sheet.getRow(r);
            if (row == null) continue;
            boolean anyNonEmpty = false;
            for (int c = 0; c < maxColumns; c++) {
                Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String text = formatter.formatCellValue(cell, evaluator);
                if (text != null && !text.trim().isEmpty()) {
                    anyNonEmpty = true;
                    break;
                }
            }
            if (anyNonEmpty) {
                return r; // r is 0-based index; we use it in loops as 1..r
            }
        }
        return 0;
    }

    /**
     * Enhanced cell validation that collects all errors for a single cell
     *
     * Supported rule types:
     *  - number   : numeric (accepts commas), floats like 0.02, 0.2833294
     *  - percent  : requires trailing % e.g. 12.00% (validator expects % in the string)
     *  - currency : numeric with commas allowed e.g. 86,000,000.00
     *  - date     : validated with rule.format (SimpleDateFormat)
     *  - text     : validated with regex if provided
     */
    private void validateCellDetailed(String value, ColumnValidationRule rule, List<String> errors) {
        // Required check
        if (rule.isRequired() && (value == null || value.trim().isEmpty())) {
            errors.add("is required");
            return;
        }

        // Skip further checks if empty and not required
        if (value == null || value.trim().isEmpty()) return;

        String type = (rule.getType() == null) ? "" : rule.getType().toLowerCase(Locale.ROOT);
        switch (type) {
            case "number": {
                String normalized = value.trim().replaceAll(",", "");
                if (normalized.endsWith("%")) {
                    errors.add("must be a numeric value (no % sign)");
                    return;
                }
                try {
                    double num = Double.parseDouble(normalized);
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
            }

            case "percent": {
                String v = value.trim();
                if (!v.endsWith("%")) {
                    errors.add("must be a percentage string ending with % (e.g. 12.00%)");
                    break;
                }
                String numericPart = v.substring(0, v.length() - 1).trim().replaceAll(",", "");
                try {
                    double num = Double.parseDouble(numericPart);
                    if (rule.getMin() != null && num < rule.getMin()) {
                        errors.add("must be >= " + rule.getMin() + "%");
                    }
                    if (rule.getMax() != null && num > rule.getMax()) {
                        errors.add("must be <= " + rule.getMax() + "%");
                    }
                } catch (NumberFormatException e) {
                    errors.add("must be a percentage number like 12.00%");
                }
                break;
            }

            case "currency": {
                String normalized = value.trim().replaceAll(",", "");
                try {
                    double num = Double.parseDouble(normalized);
                    if (rule.getMin() != null && num < rule.getMin()) {
                        errors.add("must be >= " + rule.getMin());
                    }
                    if (rule.getMax() != null && num > rule.getMax()) {
                        errors.add("must be <= " + rule.getMax());
                    }
                } catch (NumberFormatException e) {
                    errors.add("must be a currency numeric value (e.g. 86,000,000.00)");
                }
                break;
            }

            case "date": {
                String format = rule.getFormat();
                if (format == null || format.trim().isEmpty()) {
                    errors.add("date format not specified in configuration");
                    return;
                }

                try {
                    SimpleDateFormat sdf = new SimpleDateFormat(format);
                    sdf.setLenient(false);
                    sdf.parse(value.trim());
                } catch (Exception e) {
                    errors.add("must match date format " + format + " (current value: '" + value + "')");
                }
                break;
            }

            case "text": {
                if (rule.getRegex() != null && !rule.getRegex().isEmpty()) {
                    if (!value.matches(rule.getRegex())) {
                        errors.add("format is invalid");
                    }
                }
                break;
            }

            default:
                break;
        }
    }
}
