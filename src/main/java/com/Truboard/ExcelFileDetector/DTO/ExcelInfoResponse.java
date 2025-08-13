package com.Truboard.ExcelFileDetector.DTO;

import java.util.List;

public class ExcelInfoResponse {
    private int sheetCount;
    private List<String> sheetNames;

    public ExcelInfoResponse(int sheetCount, List<String> sheetNames) {
        this.sheetCount = sheetCount;
        this.sheetNames = sheetNames;
    }

    public int getSheetCount() {
        return sheetCount;
    }

    public List<String> getSheetNames() {
        return sheetNames;
    }
}