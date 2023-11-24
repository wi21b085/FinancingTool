package com.example.financingtool;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcel {

    public static void main(String[] args) {
        try {
            // Erstelle ein neues Workbook und eine Tabelle
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("MeineDaten");

            // Daten zum Schreiben
            Object[][] daten = {
                    {"Name", "Alter", "Stadt"},
                    {"John Doe", 25, "New York"},
                    {"Jane Doe", 30, "London"},
                    {"Max Mustermann", 22, "Berlin"}
            };

            // Iteriere durch die Daten und schreibe sie in die Tabelle
            int rowNum = 0;
            for (Object[] row : daten) {
                Row excelRow = sheet.createRow(rowNum++);
                int colNum = 0;
                for (Object field : row) {
                    Cell cell = excelRow.createCell(colNum++);
                    if (field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);
                    }
                    // Füge weitere Bedingungen hinzu, falls andere Datentypen unterstützt werden sollen
                }
            }

            // Schreibe das Workbook in eine Excel-Datei
            try (FileOutputStream outputStream = new FileOutputStream("Ausgabe.xlsx")) {
                workbook.write(outputStream);
            }

            // Schließe das Workbook
            workbook.close();

            System.out.println("Excel-Datei wurde erfolgreich erstellt.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
