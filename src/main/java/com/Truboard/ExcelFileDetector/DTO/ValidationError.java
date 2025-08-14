package com.Truboard.ExcelFileDetector.DTO;

public class ValidationError {
    private final String columnName;
    private final int rowNumber;
    private final int columnIndex; // New field for exact cell position
    private final String message;
    private final String cellValue; // New field for current cell value

    public ValidationError(String columnName, int rowNumber, int columnIndex, String message, String cellValue) {
        this.columnName = columnName;
        this.rowNumber = rowNumber;
        this.columnIndex = columnIndex;
        this.message = message;
        this.cellValue = cellValue;
    }

    // Legacy constructor for backward compatibility
    public ValidationError(String columnName, int rowNumber, String message) {
        this(columnName, rowNumber, -1, message, "");
    }

    public String getColumnName() {
        return columnName;
    }

    public int getRowNumber() {
        return rowNumber;
    }

    public int getColumnIndex() {
        return columnIndex;
    }

    public String getMessage() {
        return message;
    }

    public String getCellValue() {
        return cellValue;
    }
}