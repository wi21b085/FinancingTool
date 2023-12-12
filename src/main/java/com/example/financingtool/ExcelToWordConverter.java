package com.example.financingtool;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import java.io.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;


public class ExcelToWordConverter {

    public static void exportExcelToWord() {
        String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
        String wordFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.docx";

        try {
            FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("Gesamtinvestitionskosten");
            XWPFDocument document = new XWPFDocument();

            // Create a paragraph and run to add content
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText("This is a sample document content.");
            System.out.println(15840/2);
            double x=15840/27.94;
            System.out.println("x= "+x);
            System.out.println(x*29.7);
            // Set Word document in landscape orientation
            document.getDocument().getBody().addNewSectPr().addNewPgSz().setW(x*29.7);
            document.getDocument().getBody().addNewSectPr().addNewPgSz().setH(x*21);

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
                            // Round the numerical value to two decimal places
                            double roundedValue = Math.round(row.getCell(colIdx).getNumericCellValue() * 100.0) / 100.0;
                            cell.setText(String.format("%.2f", roundedValue));
                        } else {
                            cell.setText(row.getCell(colIdx).toString());
                        }
                    }
                }
            }

            document.write(out);
            out.close();
            workbook.close();

            System.out.println("Data successfully exported from Excel to Word.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
