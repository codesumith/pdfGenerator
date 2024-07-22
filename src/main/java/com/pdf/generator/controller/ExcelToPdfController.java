package com.pdf.generator.controller;

import java.io.IOException;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import com.pdf.generator.service.ExcelToPdfService;

@RestController
public class ExcelToPdfController {

    @Autowired
    private ExcelToPdfService excelToPdfService;

    @GetMapping("/generate-pdf")
    public String generatePdf() {
        try {
            excelToPdfService.generatePdfFromExcel();
            return "PDF generated successfully.";
        } catch (IOException e) {
            e.printStackTrace();
            return "Error generating PDF: " + e.getMessage();
        }
    }
}


