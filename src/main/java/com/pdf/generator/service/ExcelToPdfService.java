package com.pdf.generator.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

@Service
public class ExcelToPdfService {

    private static final String EXCEL_FILE_PATH = "/Users/sumith/Downloads/Excel BI_ETLI_PGI_Plan_without Filter_V1 1.xlsm";
    private static final String PDF_FILE_PATH = "/Users/sumith/Downloads/file.pdf";
    private static final String SHEET_NAME = "output";

    public void generatePdfFromExcel() throws IOException {
        try (FileInputStream fis = new FileInputStream(new File(EXCEL_FILE_PATH));
             XSSFWorkbook workbook = new XSSFWorkbook(fis);
             PdfWriter writer = new PdfWriter(new FileOutputStream(PDF_FILE_PATH));
             PdfDocument pdfDoc = new PdfDocument(writer);
             Document document = new Document(pdfDoc)) {

            XSSFSheet sheet = workbook.getSheet(SHEET_NAME);
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            // Read the Excel sheet and write to PDF
            for (Row row : sheet) {
                StringBuilder rowText = new StringBuilder();
                for (Cell cell : row) {
                    String cellValue = getCellValue(cell, evaluator);
                    rowText.append(cellValue).append("\t"); // Tab-separated for columns
                }
                document.add(new Paragraph(rowText.toString()));
            }
        }
    }

    private String getCellValue(Cell cell, FormulaEvaluator evaluator) {
        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue().toString();
                    } else {
                        return String.valueOf(cell.getNumericCellValue());
                    }
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    // Evaluate the formula and get the result
                    CellValue cellValue = evaluator.evaluate(cell);
                    switch (cellValue.getCellType()) {
                        case STRING:
                            return cellValue.getStringValue();
                        case NUMERIC:
                            return String.valueOf(cellValue.getNumberValue());
                        case BOOLEAN:
                            return String.valueOf(cellValue.getBooleanValue());
                        default:
                            return "Unsupported formula result";
                    }
                default:
                    return "";
            }
        } catch (Exception e) {
            return "Error evaluating cell: " + e.getMessage();
        }
    }
}
