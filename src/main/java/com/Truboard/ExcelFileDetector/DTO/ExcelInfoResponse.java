package com.Truboard.ExcelFileDetector.DTO;

import java.util.List;

public class ExcelInfoResponse {
    private int sheetCount;
    private List<String> sheetNames;
    private List<List<String>> sheetData; 

    public ExcelInfoResponse(int sheetCount, List<String> sheetNames, List<List<String>> sheetData) {
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

    public List<List<String>> getSheetData() {
        return sheetData;
    }
}
