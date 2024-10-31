package com.saroj;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteGradesToExcel {
    public static void main(String[] args) {
        String excelFilePath = "/Users/sarojshrestha/Desktop/Zorba/ExcelFile/src/main/resources/students.xlsx";

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            int gradeColumnIndex = 6;

            // Add "Grade" header if it doesn't exist
            if (headerRow.getLastCellNum() <= gradeColumnIndex ||
                    headerRow.getCell(gradeColumnIndex) == null ||
                    !headerRow.getCell(gradeColumnIndex).getStringCellValue().equals("Grade")) {
                headerRow.createCell(gradeColumnIndex).setCellValue("Grade");
            }

            // Calculate and write grades for each student
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    float finalScore = (float) row.getCell(5).getNumericCellValue();
                    String grade;

                    // Determine the grade based on the final score
                    if (finalScore > 80) {
                        grade = "A";
                    } else if (finalScore > 60 && finalScore <= 80) {
                        grade = "B";
                    } else if (finalScore > 40 && finalScore <= 60) {
                        grade = "C";
                    } else {
                        grade = "D";
                    }

                    // Print grade to console for debugging
                    System.out.println("Student " + i + ": Final Score = " + finalScore + ", Grade = " + grade);

                    // Write the grade to the corresponding cell
                    row.createCell(gradeColumnIndex).setCellValue(grade);
                }
            }

            // Write changes back to the Excel file
            try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
                workbook.write(fos);
                System.out.println("Grades successfully added to the Excel file!");
            }

        } catch (IOException e) {
            System.err.println("Error processing the Excel file: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
