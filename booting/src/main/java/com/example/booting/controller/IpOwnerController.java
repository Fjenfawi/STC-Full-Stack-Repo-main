package com.example.booting.controller;
import com.example.booting.model.IPOwner;
import com.example.booting.service.IpOwnerService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
@RestController
@RequestMapping("/api/ip-owner")
public class IpOwnerController {
    @Autowired IpOwnerService ipOwnerService;


    //upload all process Excel file
    @PostMapping("/process")
    public ResponseEntity<String>processFile(){
        
       
        try{
           
           // process the file
           ipOwnerService.processExcelFile();
           return ResponseEntity.ok("file uploaded successfully");
        }catch (Exception e) {
            return ResponseEntity.status(500).body("Error processing file: " + e.getMessage());
        }

    }
        // Retrieve all IP-owner records
    @GetMapping("/all")
    public ResponseEntity<List<IPOwner>> getAllIpOwners() {
        List<IPOwner> ipOwners = ipOwnerService.getAllIpOwners();
        return ResponseEntity.ok(ipOwners);
    }



}
