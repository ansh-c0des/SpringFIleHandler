package com.Truboard.ExcelFileDetector.DTO;

public class ErrorHighlightResponse {
    private final String fileId;
    private final String fileName;
    private final String downloadUrl;
    private final int errorCount;
    private final String message;

    public ErrorHighlightResponse(String fileId, String fileName, String downloadUrl, int errorCount, String message) {
        this.fileId = fileId;
        this.fileName = fileName;
        this.downloadUrl = downloadUrl;
        this.errorCount = errorCount;
        this.message = message;
    }

    public String getFileId() {
        return fileId;
    }

    public String getFileName() {
        return fileName;
    }

    public String getDownloadUrl() {
        return downloadUrl;
    }

    public int getErrorCount() {
        return errorCount;
    }

    public String getMessage() {
        return message;
    }
}