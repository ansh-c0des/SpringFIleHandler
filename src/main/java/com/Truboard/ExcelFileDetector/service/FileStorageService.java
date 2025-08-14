package com.Truboard.ExcelFileDetector.service;

import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.UUID;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

@Service
public class FileStorageService {
    
    private static class FileData {
        private final byte[] content;
        private final String originalFileName;
        private final String contentType;
        private final long timestamp;
        
        public FileData(byte[] content, String originalFileName, String contentType) {
            this.content = content;
            this.originalFileName = originalFileName;
            this.contentType = contentType;
            this.timestamp = System.currentTimeMillis();
        }
        
        public byte[] getContent() { return content; }
        public String getOriginalFileName() { return originalFileName; }
        public String getContentType() { return contentType; }
        public long getTimestamp() { return timestamp; }
    }
    
    private final ConcurrentHashMap<String, FileData> fileStorage = new ConcurrentHashMap<>();
    private final ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(1);
    
    // Clean up files older than 1 hour
    private static final long FILE_EXPIRY_TIME = 60 * 60 * 1000; // 1 hour in milliseconds
    
    public FileStorageService() {
        // Schedule cleanup task to run every 30 minutes
        scheduler.scheduleAtFixedRate(this::cleanupExpiredFiles, 30, 30, TimeUnit.MINUTES);
    }
    
    /**
     * Store uploaded file and return unique file ID
     */
    public String storeFile(MultipartFile file) throws IOException {
        String fileId = UUID.randomUUID().toString();
        byte[] content = file.getBytes();
        String originalFileName = file.getOriginalFilename();
        String contentType = file.getContentType();
        
        fileStorage.put(fileId, new FileData(content, originalFileName, contentType));
        return fileId;
    }
    
    /**
     * Store processed file content (for modified Excel files)
     */
    public String storeProcessedFile(ByteArrayOutputStream fileContent, String fileName, String contentType) {
        String fileId = UUID.randomUUID().toString();
        fileStorage.put(fileId, new FileData(fileContent.toByteArray(), fileName, contentType));
        return fileId;
    }
    
    /**
     * Retrieve stored file by ID
     */
    public FileData getFile(String fileId) {
        return fileStorage.get(fileId);
    }
    
    /**
     * Get file content as byte array
     */
    public byte[] getFileContent(String fileId) {
        FileData fileData = fileStorage.get(fileId);
        return fileData != null ? fileData.getContent() : null;
    }
    
    /**
     * Get original file name
     */
    public String getFileName(String fileId) {
        FileData fileData = fileStorage.get(fileId);
        return fileData != null ? fileData.getOriginalFileName() : null;
    }
    
    /**
     * Get content type
     */
    public String getContentType(String fileId) {
        FileData fileData = fileStorage.get(fileId);
        return fileData != null ? fileData.getContentType() : "application/octet-stream";
    }
    
    /**
     * Remove file from storage
     */
    public void removeFile(String fileId) {
        fileStorage.remove(fileId);
    }
    
    /**
     * Check if file exists
     */
    public boolean fileExists(String fileId) {
        return fileStorage.containsKey(fileId);
    }
    
    /**
     * Clean up expired files
     */
    private void cleanupExpiredFiles() {
        long currentTime = System.currentTimeMillis();
        fileStorage.entrySet().removeIf(entry -> 
            (currentTime - entry.getValue().getTimestamp()) > FILE_EXPIRY_TIME
        );
    }
    
    /**
     * Get storage statistics
     */
    public int getStoredFileCount() {
        return fileStorage.size();
    }
}