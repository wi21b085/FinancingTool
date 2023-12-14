package com.example.financingtool;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.List;


public class ExcelToWordConverter {
    static String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
    static String wordFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.docx";
    static  String pdfFilePath= "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.pdf";

    public static void exportExcelToWord() {


        try {
            FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("Gesamtinvestitionskosten");
            XWPFDocument document = new XWPFDocument();

            // Create a paragraph and run to add content
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
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


    public static void convertWordToPDF() {
        try {
            FileInputStream in = new FileInputStream(wordFilePath);
            XWPFDocument document = new XWPFDocument(in);

            PDDocument pdfDocument = new PDDocument();

            // Erstellen Sie die PDF-Seite im Querformat
            PDPage pdfPage = new PDPage(new PDRectangle(PDRectangle.A4.getHeight(), PDRectangle.A4.getWidth()));
            pdfDocument.addPage(pdfPage);

            PDPageContentStream contentStream = new PDPageContentStream(pdfDocument, pdfPage);

            try {
                List<XWPFParagraph> paragraphs = document.getParagraphs();
                for (XWPFParagraph paragraph : paragraphs) {
                    String text = paragraph.getText();
                    contentStream.setFont(PDType1Font.HELVETICA_BOLD, 12);
                    contentStream.beginText();
                    contentStream.newLineAtOffset(20, pdfPage.getMediaBox().getHeight() - 20);
                    contentStream.showText(text);
                    contentStream.newLine();
                    contentStream.endText();
                }

                List<XWPFTable> tables = document.getTables();
                for (XWPFTable table : tables) {
                    float margin = 20;
                    float yStart = pdfPage.getMediaBox().getHeight() - margin;
                    float tableWidth = pdfPage.getMediaBox().getWidth() - 2 * margin;
                    float yPosition = yStart;
                    float yBottom = margin;

                    // Draw table on the PDF page
                    drawPdfTable(contentStream, yPosition, tableWidth, yStart, yBottom, table);
                }

            } finally {
                contentStream.close();

                // Verwenden Sie die gleiche Dateipfadvariable wie für das Word-Dokument
                pdfDocument.save(pdfFilePath);
                pdfDocument.close();
                in.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private static void drawPdfTable(PDPageContentStream contentStream, float yPosition, float tableWidth, float yStart, float yBottom, XWPFTable table) throws IOException {
        float margin = 20;
        float fontSize = 12;
        float cellMargin = 5f;
        float yPositionNew = yPosition;

        for (int rowIdx = 0; rowIdx < table.getRows().size(); rowIdx++) {
            XWPFTableRow wordRow = table.getRow(rowIdx);
            List<XWPFTableCell> wordCells = wordRow.getTableCells();

            float maxHeight = 0;

            for (int colIdx = 0; colIdx < wordCells.size(); colIdx++) {
                XWPFTableCell wordCell = wordCells.get(colIdx);

                float width = tableWidth / (float) wordCells.size();
                float height = calculateCellHeight(wordCell, fontSize);
                String text = wordCell.getText();

                contentStream.beginText();
                contentStream.setFont(PDType1Font.HELVETICA_BOLD, fontSize);
                contentStream.newLineAtOffset(margin + colIdx * width + cellMargin, yPositionNew - height);
                contentStream.showText(text);
                contentStream.endText();

                maxHeight = Math.max(maxHeight, height);
            }

            yPositionNew -= maxHeight;
        }
    }

    private static float calculateCellHeight(XWPFTableCell cell, float fontSize) {
        // You may need to adjust this calculation based on your specific requirements
        // This is a simple estimation, and the actual cell height may depend on various factors.
        float lineSpacing = 1.5f; // Adjust as needed
        int numberOfLines = cell.getText().split("\\r?\\n").length;
        return numberOfLines * fontSize * lineSpacing;
    }








}
