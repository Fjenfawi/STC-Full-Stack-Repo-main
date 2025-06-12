package com.example.booting.controller;

import com.example.booting.service.ExcelFileService;

import java.util.concurrent.CompletableFuture;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
@RestController
@RequestMapping("/api/excel")
@CrossOrigin(origins = "http://localhost:5173")
public class FileUploadController {

    private final ExcelFileService excelFileService;

    @Autowired
    public FileUploadController(ExcelFileService excelFileService) {
        this.excelFileService = excelFileService;
    }

    @PostMapping("/modify")
    public ResponseEntity<byte[]> modifyExcelFile(
            @RequestParam("file") MultipartFile file,
            @RequestParam("department") String department,
            @RequestParam("location") String location) {

        try {
            CompletableFuture<byte[]> future = excelFileService.modifyExcelFileAsync(file, department, location);
            byte[] modifiedFile = future.get();

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            headers.setContentDisposition(ContentDisposition.attachment().filename(location+" "+department+" VA Trackersheet.xlsx").build());

            return new ResponseEntity<>(modifiedFile, headers, HttpStatus.OK);
        } catch (Exception e) {
            e.printStackTrace();
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}
