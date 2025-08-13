package com.Truboard.ExcelFileDetector.service;

import com.Truboard.ExcelFileDetector.DTO.ExcelInfoResponse;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.*;

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

            // Select the target sheet
            Sheet sheetToRead;
            if (sheetCount == 1) {
                sheetToRead = workbook.getSheetAt(0);
            } else {
                sheetToRead = workbook.getSheet("Data");
                if (sheetToRead == null) {
                    throw new Exception("Sheet named 'Data' not found in the workbook");
                }
            }

            // Determine max column count
            int maxColumns = 0;
            for (Row row : sheetToRead) {
                if (row.getLastCellNum() > maxColumns) {
                    maxColumns = row.getLastCellNum();
                }
            }

            // Create column-wise data
            Map<String, List<String>> columnData = new LinkedHashMap<>();
            for (int colIndex = 0; colIndex < maxColumns; colIndex++) {
                String colName = "Column " + (char) ('A' + colIndex);
                List<String> colValues = new ArrayList<>();

                for (int rowIndex = 0; rowIndex <= sheetToRead.getLastRowNum(); rowIndex++) {
                    Row row = sheetToRead.getRow(rowIndex);
                    if (row == null) {
                        colValues.add("");
                        continue;
                    }

                    Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cell.setCellType(CellType.STRING);
                    colValues.add(cell.getStringCellValue());
                }

                columnData.put(colName, colValues);
            }

            return new ExcelInfoResponse(sheetCount, sheetNames, columnData);
        }
    }
}
