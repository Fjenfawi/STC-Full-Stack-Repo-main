package com.example.booting.service;

import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.scheduling.annotation.Async;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.CompletableFuture;

import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import java.util.Calendar;
import com.example.booting.model.IPOwner;
import com.example.booting.repository.IpOwnerRepository;
@Service
public class ExcelFileService {
    private final IpOwnerRepository ipOwnerRepository;
    @Autowired
    public ExcelFileService(IpOwnerRepository ipOwnerRepository) {
        this.ipOwnerRepository = ipOwnerRepository;
    }

    @Async
    public CompletableFuture<byte[]> modifyExcelFileAsync(MultipartFile file, String department, String location) {
        try {
            // Your heavy Excel processing logic
            byte[] modified = modifyExcelFile(file, department, location);
            return CompletableFuture.completedFuture(modified);
        } catch (Exception e) {
            e.printStackTrace();
            return CompletableFuture.failedFuture(e);
        }
    }

    // Method to modify the Excel file based on department and location
    public byte[] modifyExcelFile(MultipartFile file, String department, String location) throws Exception {
        InputStream inputStream = file.getInputStream();
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream); // Load the Excel file

        // Get the first sheet (assuming there's only one sheet)
        XSSFSheet sheet = workbook.getSheetAt(0);

        XSSFRow headRow = sheet.getRow(0);
        // Step 1: Get header row and find index of "Host" column
        XSSFRow headerRow = sheet.getRow(0);
        int hostColIndex = -1;

        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            XSSFCell cell = headerRow.getCell(i);
            if (cell != null && cell.getCellType() == CellType.STRING
                    && cell.getStringCellValue().equalsIgnoreCase("Host")) {
                hostColIndex = i;
                break;
            }
        }

        if (hostColIndex == -1) {

            workbook.close();
            return new byte[] {};
        }

        // Step 2: Insert new column after "Host"
        int insertColIndex = hostColIndex + 1;

        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            XSSFRow row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }

            // Shift cells to the right starting from the last cell to insertColIndex
            for (int colIndex = row.getLastCellNum(); colIndex > insertColIndex; colIndex--) {
                XSSFCell oldCell = row.getCell(colIndex - 1);
                XSSFCell newCell = row.createCell(colIndex);

                if (oldCell != null) {
                    newCell.setCellStyle(oldCell.getCellStyle());
                    switch (oldCell.getCellType()) {
                        case STRING ->
                            newCell.setCellValue(oldCell.getStringCellValue());
                        case NUMERIC ->
                            newCell.setCellValue(oldCell.getNumericCellValue());
                        case BOOLEAN ->
                            newCell.setCellValue(oldCell.getBooleanCellValue());
                        case FORMULA ->
                            newCell.setCellFormula(oldCell.getCellFormula());
                        default -> {
                        }
                    }
                }
            }

            // Create a new empty cell in the inserted column
            row.createCell(insertColIndex);
        }

        sheet.getRow(0).getCell(insertColIndex).setCellValue("Owner");

        int riskCoulmnIndex = -1;

        for (int i = 0; i < headRow.getLastCellNum(); i++) {
            if (headRow.getCell(i).getStringCellValue().equalsIgnoreCase("Risk")) {
                riskCoulmnIndex = i;
                break;
            }

        }
        if (riskCoulmnIndex == -1) {

            workbook.close();
            return new byte[] {};
        }
        // Iterate through rows and delete if risk is none or low
        List<Integer> rowsToDelete = new ArrayList<>();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);
            if (row != null) {
                XSSFCell cell = row.getCell(riskCoulmnIndex);
                if (cell != null) {
                    String riskValue = cell.getStringCellValue().trim();
                    if (riskValue.equalsIgnoreCase("None") || riskValue.equalsIgnoreCase("Low")) {
                        rowsToDelete.add(i);
                    }
                }
            }
        }
        Collections.reverse(rowsToDelete);
        for (int rowIndex : rowsToDelete) {
            removeRow(sheet, rowIndex);
            // sheet.shiftRows(rowIndex+1,sheet.getLastRowNum(),-1);
        }

        int numRowsToInsert = 5;
        int startRowToShift = 0;
        int lastROw = sheet.getLastRowNum();
        sheet.shiftRows(startRowToShift, lastROw, numRowsToInsert, true, false);

        // sheet.shiftRows(insertRowIndex, sheet.getLastRowNum(), numRowsToInsert);
        for (int i = 0; i < numRowsToInsert; i++) {
            sheet.createRow(startRowToShift + i);
        }

        XSSFRow firstRow = sheet.getRow(0);
        if (firstRow == null) {
            firstRow = sheet.createRow(0);
        }
        XSSFCellStyle style = workbook.createCellStyle();
        XSSFColor customColor = new XSSFColor();
        customColor.setARGBHex("FF7030A0"); // Adding 'FF' for full opacity

        ((XSSFCellStyle) style).setFillForegroundColor(customColor);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for (int j = 0; j < 4; j++) {

            for (int i = 0; i < 15; i++) {
                XSSFCell cell = firstRow.getCell(i);
                if (cell == null) {
                    cell = firstRow.createCell(i);

                }
                cell.setCellStyle(style);

            }
            firstRow = sheet.getRow(j);
        }
        XSSFFont font = workbook.createFont();
        XSSFColor fontColor = new XSSFColor(new byte[] { (byte) 255, (byte) 255, (byte) 255 }, null);
        font.setColor(fontColor);

        style.setFont(font);
        XSSFRow mainRow = sheet.getRow(5);
        int lastCoulmnIndex = 0;
        for (int i = 0; i < mainRow.getLastCellNum(); i++) {
            if (mainRow.getCell(i) == null) {
                lastCoulmnIndex = i - 1;
                break;
            }

        }
        for (int i = 0; i < 28; i++) {
            XSSFCell cell = mainRow.getCell(i);
            if (cell == null) {
                cell = mainRow.createCell(i);

            }
            cell.setCellStyle(style);

        }

        XSSFRow first = sheet.getRow(0);
        XSSFRow second = sheet.getRow(1);
        XSSFRow third = sheet.getRow(2);
        // Merge cells A1 to O1 (i.e., 15 cells: 0 to 14)
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 14));

        // Create or get the first cell (merged cell anchor)
        XSSFCell cell = first.getCell(0);
        if (cell == null) {
            cell = first.createCell(0);
        }

        // Set the value
        cell.setCellValue("stc");
        XSSFCellStyle style2 = workbook.createCellStyle();
        XSSFColor customColor2 = new XSSFColor(new byte[] { (byte) 112, (byte) 48, (byte) 160 }, null);
        ((XSSFCellStyle) style2).setFillForegroundColor(customColor2);
        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFFont font2 = workbook.createFont();
        font2.setBold(true);

        font2.setColor(fontColor);
        font2.setFontHeight(20);
        style2.setFont(font2);
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(style2);

        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 14));
        XSSFCell secondCell = second.getCell(0);
        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
        String formattedDate = sdf.format(new Date());
        secondCell.setCellValue("status as on: " + formattedDate);
        XSSFCellStyle style3 = workbook.createCellStyle();
        XSSFColor customColor3 = new XSSFColor(new byte[] { (byte) 112, (byte) 48, (byte) 160 }, null);
        ((XSSFCellStyle) style3).setFillForegroundColor(customColor);
        style3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        secondCell.setCellStyle(style3);
        XSSFFont font3 = workbook.createFont();
        font3.setBold(true);

        font3.setColor(fontColor);
        font3.setFontHeight(20);
        style3.setFont(font3);
        style3.setAlignment(HorizontalAlignment.CENTER);
        style3.setVerticalAlignment(VerticalAlignment.CENTER);
        secondCell.setCellStyle(style3);

        sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 14));
        XSSFCell thirdCell = third.getCell(0);
        Calendar cal = Calendar.getInstance();
        int month = cal.get(Calendar.MONTH) + 1; // Calendar.MONTH starts from 0 (Jan = 0, Feb = 1, ...)

        String quarter;
        if (month >= 1 && month <= 3) {
            quarter = "Q1";
        } else if (month >= 4 && month <= 6) {
            quarter = "Q2";
        } else if (month >= 7 && month <= 9) {
            quarter = "Q3";
        } else {
            quarter = "Q4";
        }
        SimpleDateFormat sdfYear = new SimpleDateFormat("yyyy");
        String year = sdfYear.format(new Date());
        thirdCell.setCellValue("stc " + location + " " + department + " VA Report " + year + " Final Tracker sheet"
                + year + "-" + quarter);
        thirdCell.setCellStyle(style3);

        // Utility function to normalize hostnames

        // Load the second workbook with host-owner mapping
        // FileInputStream refInput = new FileInputStream(
        //         "C:\\Users\\User\\Desktop\\STC-Full-Stack-Repo-main\\stcIPs.xlsx");

        // XSSFWorkbook refWorkbook = new XSSFWorkbook(refInput);
        // XSSFSheet refSheet = refWorkbook.getSheetAt(0); // Assuming the first sheet contains the mapping

        // // Create host → owner map
        // Map<String, String> hostOwnerMap = new HashMap<>();
        // for (int i = 1; i <= refSheet.getLastRowNum(); i++) {
        //     XSSFRow row = refSheet.getRow(i);
        //     if (row != null) {
        //         XSSFCell hostCell = row.getCell(0); // Host column (IP Address)
        //         XSSFCell ownerCell = row.getCell(1); // Owner column
        //         if (hostCell != null && ownerCell != null) {
        //             String host = hostCell.getStringCellValue().trim();
        //             String owner = ownerCell.getStringCellValue().trim();
        //             hostOwnerMap.put(host, owner);

        //         }
        //     }
        // }
        // refWorkbook.close(); // Close the reference workbook
        
        // // Now update the "Owner" column in your main sheet
        // int ownerColumnIndex = -1;
        // XSSFRow headRow1 = sheet.getRow(5); // The header row

        // // Find the index of the "Owner" column in the main sheet
        // for (int i = 0; i < headRow1.getLastCellNum(); i++) {
        //     String cellValue = headRow1.getCell(i).getStringCellValue();

        //     if (cellValue.equalsIgnoreCase("Owner")) {
        //         ownerColumnIndex = i;
        //         break;
        //     }
        // }

        // if (ownerColumnIndex == -1) {

        //     return new byte[] {}; // Exit if the "Owner" column is not found
        // }

        // // Populate the "Owner" column with the corresponding owner name from the
        // // host-owner map
        // for (int i = 6; i <= sheet.getLastRowNum(); i++) {
        //     XSSFRow row = sheet.getRow(i);
        //     if (row != null) {
        //         XSSFCell hostCell = row.getCell(4); // Assuming "Host" is in the 5th column (index 4)
        //         if (hostCell != null) {
        //             String host = hostCell.getStringCellValue().trim();
        //             String owner = hostOwnerMap.getOrDefault(host, "Not there"); // Default to "Not there" if not found

        //             XSSFCell ownerCell = row.getCell(ownerColumnIndex); // Get the "Owner" cell
        //             if (ownerCell == null) {
        //                 ownerCell = row.createCell(ownerColumnIndex); // If the cell is null, create it
        //             }
        //             ownerCell.setCellValue(owner); // Set the owner value in the cell

        //         }
        //     }
        // } // Get the original sheet











    // Get header row (same as old version, row index 5)
    Row headerRow2 = sheet.getRow(5); 
    int ownerColumnIndex = -1;

    // Find or create the "Owner" column
    for (int i = 0; i < headerRow2.getLastCellNum(); i++) {
        Cell cell2 = headerRow2.getCell(i);
        if (cell2 != null && cell2.getStringCellValue().equalsIgnoreCase("Owner")) {
            ownerColumnIndex = i;
            break;
        }
    }

    if (ownerColumnIndex == -1) {
        ownerColumnIndex = headerRow2.getLastCellNum(); // Add new "Owner" column
        headerRow2.createCell(ownerColumnIndex).setCellValue("Owner");
    }

    // Start from row 6 as in the old version
    for (int i = 6; i <= sheet.getLastRowNum(); i++) {
        Row row = sheet.getRow(i);
        if (row != null) {
            Cell ipCell = row.getCell(4); // IP is in column index 4
            if (ipCell != null) {
                String ipAddress = ipCell.getStringCellValue().trim();
               List<IPOwner> ipOwners = ipOwnerRepository.findByIpAddress(ipAddress);
if (ipOwners.isEmpty()) {
    // handle missing data
    return new byte[] {};
}
IPOwner ipOwner = ipOwners.get(0); // this is now a valid assignment

                String ownerName = (ipOwner != null) ? ipOwner.getOwnerName() : "Not there";
                Cell ownerCell = row.getCell(ownerColumnIndex);

                if (ownerCell == null) {
                    ownerCell = row.createCell(ownerColumnIndex);
                }
                ownerCell.setCellValue(ownerName);
            }
        }
    }



        XSSFSheet summarySheet = workbook.createSheet("Summary");

        int lastRowNum = sheet.getLastRowNum();
        int lastColumnNum = sheet.getRow(5).getLastCellNum();

        String startCell = "A6";
        String endCell = CellReference.convertNumToColString(lastColumnNum - 1) + (lastRowNum + 1);

        AreaReference reference = new AreaReference(startCell + ":" + endCell, workbook.getSpreadsheetVersion());
        CellReference pivotPosition = new CellReference("A1");

        XSSFPivotTable pivotTable = summarySheet.createPivotTable(reference, pivotPosition, sheet);
        pivotTable.addRowLabel(5);
        pivotTable.addRowLabel(3);


        // Save the modified file to a byte array
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        workbook.write(bos);
        workbook.close();

        return bos.toByteArray(); // Return the byte array representing the modified file
    }

    private static String normalize(String s) {
        return s.trim().toLowerCase().replaceAll("[\\u00A0\\s]+", "");
    }

    private static void removeRow(XSSFSheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        } else if (rowIndex == lastRowNum) {
            XSSFRow removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }
}