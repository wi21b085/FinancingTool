package com.example.financingtool;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;


public class Excel_Plugin_V1 {


    public static void main(String[] args) {
        try {

            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";

            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            String sheetName = "Basisinformation";

            Sheet sheet = workbook.getSheet(sheetName);

            // Iteriere durch die Zeilen
            for (Row row : sheet) {
                // Iteriere durch die Zellen
                for (Cell cell : row) {
                    // Lies den Zellenwert abhängig vom Zelltyp
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t|");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t|");
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t|");
                            break;
                        case BLANK:
                            System.out.print("\t");
                            break;
                        default:
                            System.out.print("\t");
                    }
                }
                System.out.println();
            }
            fileInputStream.close();
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
