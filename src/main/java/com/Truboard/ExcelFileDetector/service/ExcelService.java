package com.Truboard.ExcelFileDetector.service;

import com.Truboard.ExcelFileDetector.DTO.ExcelInfoResponse;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelService {

    public ExcelInfoResponse extractExcelInfo(MultipartFile file) throws Exception {
        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            int sheetCount = workbook.getNumberOfSheets();
            List<String> sheetNames = new ArrayList<>();

            for (int i = 0; i < sheetCount; i++) {
                sheetNames.add(workbook.getSheetName(i));
            }

            // Decide which sheet to read
            Sheet sheetToRead;
            if (sheetCount == 1) {
                sheetToRead = workbook.getSheetAt(0);
            } else {
                sheetToRead = workbook.getSheet("Data");
                if (sheetToRead == null) {
                    throw new Exception("Sheet named 'Data' not found in the workbook");
                }
            }

            // Read sheet data into a list of lists
            List<List<String>> sheetData = new ArrayList<>();
            for (Row row : sheetToRead) {
                List<String> rowData = new ArrayList<>();
                for (Cell cell : row) {
                    cell.setCellType(CellType.STRING); // force text for simplicity
                    rowData.add(cell.getStringCellValue());
                }
                sheetData.add(rowData);
            }

            return new ExcelInfoResponse(sheetCount, sheetNames, sheetData);
        }
    }
}
