package com.Truboard.ExcelFileDetector.DTO;

import java.util.List;
import java.util.Map;

public class ExcelInfoResponse {
    private int sheetCount;
    private List<String> sheetNames;
    private Map<String, List<String>> sheetData; // column → list of values
    private List<String> errors; // validation errors (string format for backward compatibility)
    private List<ValidationError> detailedErrors; // New field for detailed error information
    private String fileId; // New field for file tracking
    private String fileType; // "xlsx" or "json"

    public ExcelInfoResponse(int sheetCount, List<String> sheetNames,
                              Map<String, List<String>> sheetData,
                              List<String> errors) {
        this.sheetCount = sheetCount;
        this.sheetNames = sheetNames;
        this.sheetData = sheetData;
        this.errors = errors;
    }

    // Enhanced constructor
    public ExcelInfoResponse(int sheetCount, List<String> sheetNames,
                              Map<String, List<String>> sheetData,
                              List<String> errors,
                              List<ValidationError> detailedErrors,
                              String fileId,
                              String fileType) {
        this.sheetCount = sheetCount;
        this.sheetNames = sheetNames;
        this.sheetData = sheetData;
        this.errors = errors;
        this.detailedErrors = detailedErrors;
        this.fileId = fileId;
        this.fileType = fileType;
    }

    public int getSheetCount() { return sheetCount; }
    public List<String> getSheetNames() { return sheetNames; }
    public Map<String, List<String>> getSheetData() { return sheetData; }
    public List<String> getErrors() { return errors; }
    public List<ValidationError> getDetailedErrors() { return detailedErrors; }
    public String getFileId() { return fileId; }
    public String getFileType() { return fileType; }

    public void setDetailedErrors(List<ValidationError> detailedErrors) {
        this.detailedErrors = detailedErrors;
    }

    public void setFileId(String fileId) {
        this.fileId = fileId;
    }

    public void setFileType(String fileType) {
        this.fileType = fileType;
    }
}