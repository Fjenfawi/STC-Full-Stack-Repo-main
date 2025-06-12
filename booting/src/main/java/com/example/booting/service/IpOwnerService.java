package com.example.booting.service;
import com.example.booting.model.IPOwner;
import com.example.booting.repository.IpOwnerRepository;

import jakarta.annotation.PostConstruct;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
@Service
public class IpOwnerService {
@Autowired
private IpOwnerRepository IpOwnerRepository;

//Method to read excel and store data in the database
public void processExcelFile() {
    String filePath = "C:\\Users\\User\\Desktop\\STC-Full-Stack-Repo-main\\stcIPs.xlsx"; // Hardcoded file path

    File file = new File(filePath);
    if (!file.exists()) {
        System.out.println("File not found at: " + filePath);
        return;
    }

    try (FileInputStream fis = new FileInputStream(file);
         Workbook workbook = new XSSFWorkbook(fis)) {

        Sheet sheet = workbook.getSheetAt(0);
        List<IPOwner> ipOwners = new ArrayList<>();

        for (Row row : sheet) {
            if (row.getRowNum() == 0) continue; // Skip header row

            Cell ipCell = row.getCell(0);
            Cell ownerCell = row.getCell(1);

            if (ipCell != null && ownerCell != null) {
                String ipAddress = ipCell.getStringCellValue().trim();
                String ownerName = ownerCell.getStringCellValue().trim();
        
                ipOwners.add(new IPOwner(ipAddress, ownerName));
            }
        }

        IpOwnerRepository.saveAll(ipOwners);
        System.out.println("Data successfully stored in H2 database!");
    } catch (IOException e) {
        e.printStackTrace();
        System.out.println("Error reading Excel file.");
    }
}
    @PostConstruct
public void init() {
    processExcelFile();
}
// Fetch all stored IP-owner records
public List<IPOwner> getAllIpOwners() {
    return IpOwnerRepository.findAll();
}
}
