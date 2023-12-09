package com.example.financingtool;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.*;

public class ExcelToWordConverter {

    public static void exportExcelToWord() {
        String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
        String wordFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.docx";

        try {
            FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("Gesamtinvestitionskosten");

            XWPFDocument document = new XWPFDocument();
            FileOutputStream out = new FileOutputStream(wordFilePath);

            XWPFTable table = document.createTable();

            String[] headers = {"Name", "Netto", "Ust", "% der Ust", "Brutto", "in %"};
            XWPFTableRow headerRow = table.getRow(0);
            for (int i = 0; i < headers.length; i++) {
                XWPFTableCell cell = headerRow.getCell(i);
                if (cell == null) {
                    cell = headerRow.createCell();
                }
                cell.setText(headers[i]);
            }

            for (int rowIdx = 1; rowIdx <= 9; rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                XWPFTableRow dataRow = table.createRow();

                for (int colIdx = 0; colIdx < headers.length; colIdx++) {
                    XWPFTableCell cell = dataRow.getCell(colIdx);
                    if (cell == null) {
                        cell = dataRow.createCell();
                    }

                    if (row.getCell(colIdx) != null) {
                        if (row.getCell(colIdx).getCellType() == CellType.FORMULA) {
                            cell.setText(String.valueOf(row.getCell(colIdx).getNumericCellValue()));
                        } else {
                            cell.setText(row.getCell(colIdx).toString());
                        }
                    }
                }
            }

            document.write(out);
            out.close();
            workbook.close();

            System.out.println("Daten erfolgreich von Excel nach Word exportiert.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

