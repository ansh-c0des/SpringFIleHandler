package com.Truboard.ExcelFileDetector.DTO;

import java.util.List;
import java.util.Map;

public class ExcelInfoResponse {
    private int sheetCount;
    private List<String> sheetNames;
    private Map<String, List<String>> sheetData; // column → list of values

    public ExcelInfoResponse(int sheetCount, List<String> sheetNames, Map<String, List<String>> sheetData) {
        this.sheetCount = sheetCount;
        this.sheetNames = sheetNames;
        this.sheetData = sheetData;
    }

    public int getSheetCount() {
        return sheetCount;
    }

    public List<String> getSheetNames() {
        return sheetNames;
    }

    public Map<String, List<String>> getSheetData() {
        return sheetData;
    }
}
