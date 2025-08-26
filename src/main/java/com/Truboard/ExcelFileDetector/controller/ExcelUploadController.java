package com.Truboard.ExcelFileDetector.controller;

import com.Truboard.ExcelFileDetector.DTO.ExcelInfoResponse;
import com.Truboard.ExcelFileDetector.DTO.ErrorHighlightResponse;
import com.Truboard.ExcelFileDetector.service.ExcelService;
import com.Truboard.ExcelFileDetector.service.FileStorageService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RestController
@CrossOrigin(origins = "http://localhost:5173")
@RequestMapping("/api/excel")
public class ExcelUploadController {

    @Autowired
    private final ExcelService excelService;

    @Autowired
    private final FileStorageService fileStorageService;

    public ExcelUploadController(ExcelService excelService, FileStorageService fileStorageService) {
        this.excelService = excelService;
        this.fileStorageService = fileStorageService;
    }

    /**
     * Accepts either .xlsx or .json file.
     * - .xlsx -> validated and auto-filled (with yellow highlighting for auto-filled cells)
     * - .json -> validated using same rules (JSON can be array of objects or object-of-arrays)
     *
     * Auto-fill functionality now happens during upload:
     * - Empty critical cells are filled with default values
     * - Auto-filled cells are highlighted in yellow with comments
     * - Modified file is stored in the system
     * - Response includes auto-fill information
     */
    @PostMapping("/upload")
    public ResponseEntity<?> uploadExcelOrJson(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty()) {
            return ResponseEntity.badRequest().body("File is empty");
        }

        String filename = file.getOriginalFilename();
        if (filename == null) {
            return ResponseEntity.badRequest().body("File must have a name/extension");
        }

        String lower = filename.toLowerCase();

        try {
            ExcelInfoResponse response;

            if (lower.endsWith(".xlsx")) {
                // For Excel files: perform auto-fill during upload
                response = excelService.extractAndProcessExcelInfo(file);
            } else if (lower.endsWith(".json")) {
                // JSON files: validation only (no auto-fill needed)
                response = excelService.extractJsonInfo(file);
            } else {
                return ResponseEntity.badRequest().body("Only .xlsx and .json files are supported");
            }

            return ResponseEntity.ok(response);

        } catch (Exception e) {
            // Return 500 with message; validation errors are returned in response.errors (200)
            return ResponseEntity.status(500).body("Error processing file: " + e.getMessage());
        }
    }

    /**
     * Generate and download Excel file with validation errors highlighted
     *
     * Now only handles error highlighting (red background + comments).
     * Auto-fill functionality has been moved to upload phase.
     */
    @GetMapping("/download-highlighted/{fileId}")
    public ResponseEntity<?> downloadHighlightedExcel(@PathVariable String fileId) {
        try {
            if (!fileStorageService.fileExists(fileId)) {
                return ResponseEntity.notFound().build();
            }

            String originalFileName = fileStorageService.getFileName(fileId);
            if (originalFileName == null || !originalFileName.toLowerCase().endsWith(".xlsx")) {
                return ResponseEntity.badRequest()
                        .body("Error highlighting is only supported for Excel (.xlsx) files");
            }

            // Generate highlighted Excel with only error highlighting (no auto-fill)
            byte[] highlightedFileContent = excelService.generateErrorHighlightedExcel(fileId);

            // Generate highlighted filename
            String highlightedFileName = generateHighlightedFileName(originalFileName);

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDispositionFormData("attachment", highlightedFileName);
            headers.setContentLength(highlightedFileContent.length);

            return ResponseEntity.ok()
                    .headers(headers)
                    .body(highlightedFileContent);

        } catch (Exception e) {
            return ResponseEntity.status(500)
                    .body("Error generating highlighted file: " + e.getMessage());
        }
    }

    /**
     * Get information about highlighted file availability
     */
    @GetMapping("/highlight-info/{fileId}")
    public ResponseEntity<?> getHighlightInfo(@PathVariable String fileId) {
        try {
            if (!fileStorageService.fileExists(fileId)) {
                return ResponseEntity.notFound().build();
            }

            String originalFileName = fileStorageService.getFileName(fileId);
            if (originalFileName == null || !originalFileName.toLowerCase().endsWith(".xlsx")) {
                return ResponseEntity.ok(new ErrorHighlightResponse(
                        fileId,
                        originalFileName,
                        null,
                        0,
                        "Error highlighting is only supported for Excel (.xlsx) files"
                ));
            }

            String downloadUrl = "/api/excel/download-highlighted/" + fileId;
            String highlightedFileName = generateHighlightedFileName(originalFileName);

            return ResponseEntity.ok(new ErrorHighlightResponse(
                    fileId,
                    highlightedFileName,
                    downloadUrl,
                    -1, // Will be calculated when file is generated
                    "Highlighted file is ready for download"
            ));

        } catch (Exception e) {
            return ResponseEntity.status(500)
                    .body("Error getting highlight info: " + e.getMessage());
        }
    }

    /**
     * Delete stored file
     */
    @DeleteMapping("/files/{fileId}")
    public ResponseEntity<?> deleteStoredFile(@PathVariable String fileId) {
        try {
            if (!fileStorageService.fileExists(fileId)) {
                return ResponseEntity.notFound().build();
            }

            fileStorageService.removeFile(fileId);
            return ResponseEntity.ok().body("File deleted successfully");

        } catch (Exception e) {
            return ResponseEntity.status(500)
                    .body("Error deleting file: " + e.getMessage());
        }
    }

    /**
     * Get storage statistics (for monitoring)
     */
    @GetMapping("/storage/stats")
    public ResponseEntity<?> getStorageStats() {
        try {
            int fileCount = fileStorageService.getStoredFileCount();
            return ResponseEntity.ok().body(
                    "Stored files count: " + fileCount
            );
        } catch (Exception e) {
            return ResponseEntity.status(500)
                    .body("Error getting storage stats: " + e.getMessage());
        }
    }

    /**
     * Generate filename for highlighted Excel file
     */
    private String generateHighlightedFileName(String originalFileName) {
        if (originalFileName == null) {
            return "highlighted_file.xlsx";
        }

        String nameWithoutExtension;
        String extension;

        int lastDotIndex = originalFileName.lastIndexOf('.');
        if (lastDotIndex > 0) {
            nameWithoutExtension = originalFileName.substring(0, lastDotIndex);
            extension = originalFileName.substring(lastDotIndex);
        } else {
            nameWithoutExtension = originalFileName;
            extension = ".xlsx";
        }

        return nameWithoutExtension + "_highlighted" + extension;
    }
}