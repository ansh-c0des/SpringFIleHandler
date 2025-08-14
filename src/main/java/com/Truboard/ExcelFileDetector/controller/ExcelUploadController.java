package com.Truboard.ExcelFileDetector.controller;

import com.Truboard.ExcelFileDetector.DTO.ExcelInfoResponse;
import com.Truboard.ExcelFileDetector.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/api/excel")
public class ExcelUploadController {

    @Autowired
    private final ExcelService excelService;

    public ExcelUploadController(ExcelService excelService) {
        this.excelService = excelService;
    }

    /**
     * Accepts either .xlsx or .json file.
     * - .xlsx -> validated as before
     * - .json -> validated using same rules (JSON can be array of objects or object-of-arrays)
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
                response = excelService.extractExcelInfo(file);
            } else if (lower.endsWith(".json")) {
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
}
