package com.example.booting.service;
import com.example.booting.model.IPOwner;
import com.example.booting.repository.IpOwnerRepository;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
@Service
public class ExcelProcessorService {

@Autowired
    private IpOwnerRepository ipOwnerRepository;

    public void updateExcelWithOwners(String filePath) {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Get first sheet

            // Find or create the "Owner" column in header row
            Row headerRow = sheet.getRow(0); // Assuming first row contains headers
            int ownerColumnIndex = -1;

            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                if (headerRow.getCell(i).getStringCellValue().equalsIgnoreCase("Owner")) {
                    ownerColumnIndex = i;
                    break;
                }
            }

            if (ownerColumnIndex == -1) { 
                ownerColumnIndex = headerRow.getLastCellNum(); // New column index
                headerRow.createCell(ownerColumnIndex).setCellValue("Owner"); // Add "Owner" column
            }

            // Process each row and fetch corresponding owner from database
            // for (int i = 1; i <= sheet.getLastRowNum(); i++) { 
            //     Row row = sheet.getRow(i);
            //     if (row != null) {
            //         Cell ipCell = row.getCell(0); // Assume IP is in first column (index 0)
            //         if (ipCell != null) {
            //             String ipAddress = ipCell.getStringCellValue().trim();
            //           //  IPOwner ipOwner = ipOwnerRepository.findByIpAddress(ipAddress);

            //             String ownerName = (ipOwner != null) ? ipOwner.getOwnerName() : "Not found";
            //             Cell ownerCell = row.getCell(ownerColumnIndex);

            //             if (ownerCell == null) {
            //                 ownerCell = row.createCell(ownerColumnIndex);
            //             }
            //             ownerCell.setCellValue(ownerName); // Update owner column
            //         }
            //     }
            // }

            // Save the updated file
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }

            System.out.println("Excel file updated successfully!");

        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Error processing Excel file.");
        }
    }

}
