package com.task.controller;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

@Controller
public class csvfileupload {

	@GetMapping("/")
    public String index() {
        return "upload"; 
    }

    @PostMapping("/upload")
    public String handleFileUpload(@RequestParam("file") MultipartFile file, 
                                   @RequestParam("startRow") int startRow, 
                                   Model model) {
        List<List<String>> data = new ArrayList<>();
        
        try {
            String fileName = file.getOriginalFilename();
            if (fileName != null && fileName.endsWith(".csv")) {
                data = processCSV(file.getInputStream(), startRow);
            } else if (fileName != null && (fileName.endsWith(".xls") || fileName.endsWith(".xlsx"))) {
                data = processExcel(file.getInputStream(), startRow);
            } else {
                model.addAttribute("error", "Unsupported file format. Please upload CSV or Excel.");
                return "upload";
            }
        } catch (Exception e) {
            model.addAttribute("error", "Error processing file: " + e.getMessage());
            return "upload";
        }

        model.addAttribute("data", data);
        return "upload";
    }

    
    private List<List<String>> processCSV(InputStream inputStream, int startRow) throws IOException {
        List<List<String>> data = new ArrayList<>();
        BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream));
        CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT);
        int currentRow = 0;

        for (CSVRecord record : csvParser) {
            if (currentRow >= startRow) {
                List<String> row = new ArrayList<>();
                record.forEach(row::add);
                data.add(row);
            }
            currentRow++;
        }
        return data;
    }

    
    private List<List<String>> processExcel(InputStream inputStream, int startRow) throws IOException {
        List<List<String>> data = new ArrayList<>();
        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        int currentRow = 0;

        for (Row row : sheet) {
            if (currentRow >= startRow) {
                List<String> rowData = new ArrayList<>();
                for (Cell cell : row) {
                    rowData.add(getCellValue(cell));
                }
                data.add(rowData);
            }
            currentRow++;
        }
        return data;
    }

   
    private String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            default: return "";
        }
    }
}
