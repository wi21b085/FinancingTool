package com.example.financingtool;

import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;

import java.io.*;
import java.util.List;

public class ExcelToWordConverter {
    private static String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
    private static String wordFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.docx";
    private static String pdfFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.pdf";
    private static XWPFDocument document;
    static ExecutiveSummary executiveSummary = new ExecutiveSummary();

    public static void setExecutiveSummary(ExecutiveSummary executiveSummary) {
        ExcelToWordConverter.executiveSummary = executiveSummary;
    }

    public static void initializeDocument() {
        document = new XWPFDocument();
    }

    public static void exportExcelToWord() {
        try {
            FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(excelFile);

            if (document == null) {
                initializeDocument();
            }

            exportSheetToWord(workbook, "Basisinformation");
            document.createParagraph().setPageBreak(true);
            exportSheetToWord(workbook, "Gesamtinvestitionskosten");
            document.createParagraph().setPageBreak(true);
            exportSheetToWord(workbook, "Mittelverwendung - Mittelherkun");
            document.createParagraph().setPageBreak(true);
            exportSheetToWord(workbook, "Wirtschaftlichkeitsrechnung");

            workbook.close();
            saveDocument();
            convertWordToPDF();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void exportSheetToWord(Workbook workbook, String sheetName) throws FileNotFoundException {
        Sheet sheet = workbook.getSheet(sheetName);
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();

        document.getDocument().getBody().addNewSectPr().addNewPgSz().setW(15840);
        document.getDocument().getBody().addNewSectPr().addNewPgSz().setH(11280);

        FileOutputStream out = new FileOutputStream(wordFilePath);

        if (sheetName.equals("Mittelverwendung - Mittelherkun")) {
            System.out.println("Mittelverwendung");
            createTable(document, sheet, 0, 5, "Mittelverwendung");
        } else if (sheetName.equals("Basisinformation")) {
            createTable(document, sheet, 0, 2, "Basisinformation");
            document.createParagraph().setPageBreak(true);
            System.out.println("Bas");
            createTable(document, sheet, 7, 8,"Stammdaten");
            createTable(document, sheet, 14, 14,"Keine Ahnung");
            document.createParagraph().setPageBreak(true);
        } else if (sheetName.equals("Gesamtinvestitionskosten")) {
            System.out.println("Ges");
            createGIKtable(sheet);
        } else if (sheetName.equals("Wirtschaftlichkeitsrechnung")) {
            document.createParagraph().setPageBreak(true);
            createTable(document, sheet, 0, 7, "Wirtschaftlichkeitsrechnung");
        }

        try {
            document.write(out);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void createTable(XWPFDocument document, Sheet sheet, int startColumn, int endColumn, String title) {
        if (title != null && !title.isEmpty()) {
            XWPFParagraph titleParagraph = document.createParagraph();
            titleParagraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun titleRun = titleParagraph.createRun();
            titleRun.setText(title);
            titleRun.setBold(true);
            titleRun.setFontSize(14);
            titleRun.addBreak(BreakType.TEXT_WRAPPING);
        }

        XWPFTable table = document.createTable();
        XWPFTableRow headerRow = table.getRow(0);

        for (int i = startColumn; i <= endColumn; i++) {
            Cell excelCell = sheet.getRow(0).getCell(i);
            if (excelCell != null && excelCell.getCellType() == CellType.STRING) {
                String cellValue = excelCell.getStringCellValue();
                if (!cellValue.isEmpty()) {
                    XWPFTableCell cell = headerRow.getCell(i - startColumn);
                    if (cell == null) {
                        cell = headerRow.createCell();
                    }
                    cell.setText(cellValue);
                }
            }
        }

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                XWPFTableRow tableRow = table.createRow();
                for (int j = startColumn; j <= endColumn; j++) {
                    Cell excelCell = row.getCell(j);
                    XWPFTableCell cell = tableRow.getCell(j - startColumn);

                    if (excelCell != null) {
                        if (excelCell.getCellType() == CellType.STRING) {
                            String cellValue = excelCell.getStringCellValue();
                            if (!cellValue.isEmpty()) {
                                if (cell == null) {
                                    cell = tableRow.createCell();
                                }
                                cell.setText(cellValue);
                            }
                        } else if (excelCell.getCellType() == CellType.NUMERIC) {
                            double numericValue = excelCell.getNumericCellValue();
                            if (cell == null) {
                                cell = tableRow.createCell();
                            }
                            cell.setText(String.valueOf(numericValue));
                        }
                    }
                }
            }
        }
    }

    private static void createGIKtable(Sheet sheet) {
        document.createParagraph().setPageBreak(true);
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
                        double roundedValue = Math.round(row.getCell(colIdx).getNumericCellValue() * 100.0) / 100.0;
                        cell.setText(String.format("%.2f", roundedValue));
                    } else {
                        cell.setText(row.getCell(colIdx).toString());
                    }
                }
            }
        }
    }

    private static void saveDocument() {
        try {
            FileOutputStream out = new FileOutputStream(wordFilePath);
            document.write(out);
            out.close();
            System.out.println("Data successfully exported from Excel to Word.");
            mergePDFs();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void mergePDFs() {
        String file1 = "src/main/resources/com/example/financingtool/Stammblattimg.pdf";
        String file2 = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.pdf";
        String outputFile = "src/main/resources/com/example/financingtool/final.pdf";

        try {
            PDDocument pdfDocument1 = PDDocument.load(new java.io.File(file1));
            PDDocument pdfDocument2 = PDDocument.load(new java.io.File(file2));

            for (int i = 0; i < pdfDocument1.getNumberOfPages(); i++) {
                PDPage page = pdfDocument1.getPage(i);
                pdfDocument2.addPage(page);
            }

            pdfDocument2.save(outputFile);
            System.out.println("Erfolgreiche Kombination der pdf's");
            pdfDocument1.close();
            pdfDocument2.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void convertWordToPDF() {
        try {
            FileInputStream in = new FileInputStream(wordFilePath);
            XWPFDocument document = new XWPFDocument(in);
            PDDocument pdfDocument = new PDDocument();

            List<XWPFParagraph> paragraphs = document.getParagraphs();
            List<XWPFTable> tables = document.getTables();

            for (int pageIndex = 0; pageIndex < Math.max(paragraphs.size(), tables.size()); pageIndex++) {
                PDPage pdfPage = new PDPage(new PDRectangle(PDRectangle.A4.getHeight(), PDRectangle.A4.getWidth()));
                pdfDocument.addPage(pdfPage);

                PDPageContentStream contentStream = new PDPageContentStream(pdfDocument, pdfPage);

                try {
                    if (pageIndex < paragraphs.size()) {
                        String text = cleanText(paragraphs.get(pageIndex).getText());
                        contentStream.setFont(PDType1Font.HELVETICA_BOLD, 12);
                        contentStream.beginText();
                        contentStream.newLineAtOffset(20, pdfPage.getMediaBox().getHeight() - 20);
                        contentStream.showText(text);
                        contentStream.newLine();
                        contentStream.endText();
                    }

                    if (pageIndex < tables.size()) {
                        XWPFTable table = tables.get(pageIndex);
                        float margin = 20;
                        float yStart = pdfPage.getMediaBox().getHeight() - margin;
                        float tableWidth = pdfPage.getMediaBox().getWidth() - 2 * margin;
                        float yPosition = yStart;
                        float yBottom = margin;

                        drawPdfTable(pdfDocument, tableWidth, yStart, yBottom, table, pageIndex == 0);
                    }
                } finally {
                    contentStream.close();
                }
            }

            pdfDocument.save(pdfFilePath);
            pdfDocument.close();
            in.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String cleanText(String text) {
        // Ersetze oder entferne unerwünschte Zeichen
        return text.replace("\n", "").replace("\r", ""); // Hier kannst du weitere Ersetzungen hinzufügen
    }


    private static void drawPdfTable(PDDocument document, float tableWidth, float yStart, float yBottom, XWPFTable table, boolean isHeader) throws IOException {
        float margin = 20;
        float fontSize = isHeader ? 12.0f : 10.0f;
        float cellMargin = 5f;
        float pageHeight = PDRectangle.A4.getHeight() - 2 * margin;
        float yPosition = yStart;
        float currentYPosition = yPosition;

        int rowIdx = 0;
        PDPageContentStream contentStream = null;

        while (rowIdx < table.getRows().size()) {
            XWPFTableRow wordRow = table.getRow(rowIdx);
            List<XWPFTableCell> wordCells = wordRow.getTableCells();

            float maxHeight = 0;

            if (contentStream == null) {
                PDPage newPage = new PDPage(new PDRectangle(PDRectangle.A4.getHeight(), PDRectangle.A4.getWidth()));
                document.addPage(newPage);
                contentStream = new PDPageContentStream(document, newPage, true, true);
                contentStream.setFont(PDType1Font.HELVETICA_BOLD, 12);
                yPosition = yStart;
            }

            for (int colIdx = 0; colIdx < wordCells.size(); colIdx++) {
                XWPFTableCell wordCell = wordCells.get(colIdx);

                float width = tableWidth / (float) wordCells.size();
                float height = calculateCellHeight(wordCell, fontSize);
                String text = wordCell.getText();

                String[] words = text.split("\\s+");
                float currentWidth = 0;

                for (String word : words) {
                    float wordWidth = PDType1Font.HELVETICA_BOLD.getStringWidth(word) / 990 * fontSize;

                    if (currentWidth + wordWidth > width && currentWidth > 0) {
                        yPosition -= maxHeight;
                        currentWidth = 0;
                    }

                    if (yPosition - height < yBottom) {
                        PDPage newPage = new PDPage(PDRectangle.A4);
                        document.addPage(newPage);
                        contentStream.close();
                        contentStream = new PDPageContentStream(document, newPage, true, true);
                        contentStream.setFont(PDType1Font.HELVETICA_BOLD, 12);

                        yPosition = yStart;
                    }

                    contentStream.beginText();
                    contentStream.newLineAtOffset(margin + colIdx * width + cellMargin + currentWidth, yPosition - height);
                    contentStream.showText(word);
                    contentStream.endText();

                    currentWidth += wordWidth;
                }

                maxHeight = Math.max(maxHeight, height);
            }

            yPosition -= maxHeight;

            if (yPosition < yBottom && rowIdx + 1 < table.getRows().size()) {
                PDPage newPage = new PDPage(PDRectangle.A4);
                document.addPage(newPage);
                contentStream.close();
                contentStream = new PDPageContentStream(document, newPage, true, true);
                contentStream.setFont(PDType1Font.HELVETICA_BOLD, 12);
                yPosition = yStart;
            } else {
                rowIdx++;
            }
        }

        if (contentStream != null) {
            contentStream.close();
        }
    }

    private static float calculateCellHeight(XWPFTableCell cell, float fontSize) {
        float lineSpacing = 1.5f;
        int numberOfLines = cell.getText().split("\\r?\\n").length;
        return numberOfLines * fontSize * lineSpacing;
    }

    static public void openNewJavaFXWindow() {
        Stage newStage = new Stage();
        javafx.scene.control.Button convertToWordButton = new javafx.scene.control.Button("Konvertierung in eine PDF");
        convertToWordButton.setOnAction(e -> handleConvertToWordButtonClick());

        VBox newRoot = new VBox(10);
        newRoot.setAlignment(Pos.CENTER);
        newRoot.getChildren().add(convertToWordButton);

        Scene newScene = new Scene(newRoot, 300, 200);
        newStage.setTitle("PDF Konvertierung");
        newStage.setScene(newScene);
        newStage.show();
    }

    private static void handleConvertToWordButtonClick() {
        ExcelToWordConverter.convertWordToPDF();
        executiveSummary.getBasDaten();
        executiveSummary.getGikData();
        executiveSummary.getWIODaten();
        executiveSummary.getWireData();
        executiveSummary.setDaten();
    }
}


