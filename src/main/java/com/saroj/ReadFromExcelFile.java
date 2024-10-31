package com.saroj;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ReadFromExcelFile {
    public static void main(String[] args) {
        String excelFilePath = "/Users/sarojshrestha/Desktop/Zorba/Java_1016_Batch_Class_Notes/FileInputOutputOperations/src/main/resources/students.xlsx";

        try (FileInputStream fis = new FileInputStream(new File(excelFilePath));
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = workbook.getSheetAt(0);
            List<Map<String, Object>> studentList = new ArrayList<>();

            // Read rows and store data as a map
            for (int i = 1; i <= sheet.getLastRowNum(); i++) { // Start from row 1 to skip header
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Map<String, Object> studentData = new HashMap<>();
                studentData.put("Student Id", (int) row.getCell(0).getNumericCellValue());
                studentData.put("Student Name", row.getCell(1).getStringCellValue());
                studentData.put("Sub1 Score", (float) row.getCell(2).getNumericCellValue());
                studentData.put("Sub2 Score", (float) row.getCell(3).getNumericCellValue());
                studentData.put("Sub3 Score", (float) row.getCell(4).getNumericCellValue());
                studentData.put("Final Score", (float) row.getCell(5).getNumericCellValue());

                studentList.add(studentData);
            }

            System.out.println(studentList);


        } catch (IOException e) {
            System.err.println("Error processing the Excel file: " + e.getMessage());
        }
    }
}
