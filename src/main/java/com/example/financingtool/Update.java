package com.example.financingtool;

import javafx.scene.control.Label;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Update {

   // private static Label resultLabel=new Label();

    public static void updateRangeOfCells(String[] newValues, int startRow, int endRow, int colIdx, Label resultLabel, String sheetName) {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
           // String sheetName = "Gesamtinvestitionskosten";
            //int startRow = 1;
            //int endRow = 9;
            //int colIdx = 1;

            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheet(sheetName);


            // Annahme: newValues enthält die neuen Werte für B3 bis B10
            for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
                String newValueStr = newValues[rowIdx - startRow];

                // Überprüfen Sie, ob die Zeichenkette nicht leer ist und nicht null ist, bevor Sie sie parsen
                if (newValueStr != null && !newValueStr.isEmpty()) {
                    double newValue = Double.parseDouble(newValueStr);
                    updateCellValue(sheet, rowIdx, colIdx, newValue);
                }
            }


            // Automatische Auswertung der Formeln im gesamten Arbeitsblatt
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            fileInputStream.close();

            FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();

            resultLabel.setText("Bereich von B2 bis B10 erfolgreich aktualisiert.");
            System.out.println("Zellen erfolgreich aktualisiert.");
            // exportExcelToWord();

        } catch (NumberFormatException | IOException e) {
            e.printStackTrace();
           resultLabel.setText("Fehler bei der Aktualisierung.");
        }
    }
    public static void updateCellValue(Sheet sheet, int rowIdx, int colIdx, double newValue) {
        Row row = sheet.getRow(rowIdx);
        if (row != null ) {
            Cell cell = row.getCell(colIdx);
            cell.setCellValue(newValue);
        }
    }

    public static void updateRangeOfCellsString(String[] newValues, int startRow, int endRow, int colIdx, Label resultLabel, String sheetName) {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            // String sheetName = "Gesamtinvestitionskosten";
            //int startRow = 1;
            //int endRow = 9;
            //int colIdx = 1;

            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheet(sheetName);


            // Annahme: newValues enthält die neuen Werte für B3 bis B10
            for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
                String newValueStr = newValues[rowIdx - startRow];

                // Überprüfen Sie, ob die Zeichenkette nicht leer ist und nicht null ist, bevor Sie sie parsen
                if (newValueStr != null && !newValueStr.isEmpty()) {
                    updateCellValueString(sheet, rowIdx, colIdx, newValueStr);
                }
            }


            // Automatische Auswertung der Formeln im gesamten Arbeitsblatt
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            fileInputStream.close();

            FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();

            resultLabel.setText("Bereich von B2 bis B10 erfolgreich aktualisiert.");
            System.out.println("Zellen erfolgreich aktualisiert.");
            // exportExcelToWord();

        } catch (NumberFormatException | IOException e) {
            e.printStackTrace();
            resultLabel.setText("Fehler bei der Aktualisierung.");
        }
    }

    public static void updateCellValueString(Sheet sheet, int rowIdx, int colIdx, String newValue) {
        Row row = sheet.getRow(rowIdx);
        Cell cell = row.getCell(colIdx);
        cell.setCellValue(newValue);
    }

    /*private void updateCellD(int rowIdx, int colIdx, String newValue) {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Gesamtinvestitionskosten";

            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheet(sheetName);

            // Überprüfen Sie, ob die Zeichenkette nicht leer ist und nicht null ist, bevor Sie sie parsen
            if (newValue != null && !newValue.isEmpty()) {
                double newCellValue = Double.parseDouble(newValue);
                updateCellValue(sheet, rowIdx, colIdx, newCellValue);
            }

            // Automatische Auswertung der Formeln im gesamten Arbeitsblatt
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            fileInputStream.close();

            FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();

            // resultLabel.setText("Zelle erfolgreich aktualisiert.");
            // System.out.println("Zelle erfolgreich aktualisiert.");

        } catch (NumberFormatException | IOException e) {
            e.printStackTrace();
            resultLabel.setText("Fehler bei der Aktualisierung von Zelle D" + rowIdx);
        }
    }*/
}
