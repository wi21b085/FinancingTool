package com.example.financingtool;

import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.List;

public class ExcelToWordConverter {
    private static String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
    private static String wordFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.docx";
    private static String pdfFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.pdf";
    private static XWPFDocument document;

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

            // Export Basisinformation
            exportSheetToWord(workbook, "Basisinformation");

            // Add a newline between the two tables
            document.createParagraph();

            // Export Gesamtinvestitionskosten
            exportSheetToWord(workbook, "Gesamtinvestitionskosten");
            document.createParagraph();
            exportSheetToWord(workbook,"Mittelverwendung - Mittelherkun");
            document.createParagraph();
            exportSheetToWord(workbook, "Wirtschaftlichkeitsrechnung");
            document.createParagraph();
            openNewJavaFXWindow();

            workbook.close();

            // Save the document to file
            saveDocument();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void exportSheetToWord(Workbook workbook, String sheetName) throws FileNotFoundException {
        Sheet sheet = workbook.getSheet(sheetName);

        // Create a paragraph and run to add content
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        double x = 15840 / 27.94;


        // Set Word document in landscape orientation
        document.getDocument().getBody().addNewSectPr().addNewPgSz().setW(x * 29.7);
        document.getDocument().getBody().addNewSectPr().addNewPgSz().setH(x * 21);

        FileOutputStream out = new FileOutputStream(wordFilePath);
        if(sheetName.equals("Mittelverwendung - Mittelherkun")){
            System.out.println("Mittelverwendung");
            createBasisinformationTable(document,sheet,0,5);
        }

        else if (sheetName.equals("Basisinformation")) {
            // Export columns A-C to Word
            createBasisinformationTable(document, sheet, 0, 2);

            // Add a newline between the two tables
            document.createParagraph().setPageBreak(true);
            System.out.println("Bas");

            // Export columns H-I to Word
            createBasisinformationTable(document, sheet, 7, 8);
        }else if (sheetName.equals("Gesamtinvestitionskosten")) {
            System.out.println("Ges");
            createGIKtable(sheet);
            document.createParagraph().setPageBreak(true);
            createBasisinformationTable(document, sheet,0,5);
        }else if(sheetName.equals("Wirtschaftlichkeitsrechnung")){
            document.createParagraph().setPageBreak(true);
            createBasisinformationTable(document, sheet, 0,7);
        }

        // Write Word document to output file
        try {
            document.write(out);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private static void createGIKtable(Sheet sheet){
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
                        // Round the numerical value to two decimal places
                        double roundedValue = Math.round(row.getCell(colIdx).getNumericCellValue() * 100.0) / 100.0;
                        cell.setText(String.format("%.2f", roundedValue));
                    } else {
                        cell.setText(row.getCell(colIdx).toString());
                    }
                }
            }
        }
    }

    private static void createBasisinformationTable(XWPFDocument document, Sheet sheet, int startColumn, int endColumn) {
        // Create table in Word
        XWPFTable table = document.createTable();
        XWPFTableRow headerRow = table.getRow(0);

        // Populate table header
        for (int i = startColumn; i <= endColumn; i++) {
            XWPFTableCell cell = headerRow.getCell(i - startColumn);
            if (cell == null) {
                cell = headerRow.createCell();
            }
            cell.setText(sheet.getRow(0).getCell(i).getStringCellValue());
        }

        // Populate table with Excel data
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) { // Überprüfen, ob die Zeile nicht null ist
                XWPFTableRow tableRow = table.createRow();
                for (int j = startColumn; j <= endColumn; j++) {
                    Cell excelCell = row.getCell(j);
                    XWPFTableCell cell = tableRow.getCell(j - startColumn);

                    if (excelCell != null) {
                        if (excelCell.getCellType() == CellType.STRING) {
                            // String-Wert
                            String cellValue = excelCell.getStringCellValue();
                            if (!cellValue.isEmpty()) {
                                if (cell == null) {
                                    cell = tableRow.createCell();
                                }
                                cell.setText(cellValue);
                            }
                        } else if (excelCell.getCellType() == CellType.NUMERIC) {
                            // Numerischer Wert
                            double numericValue = excelCell.getNumericCellValue();
                            if (cell == null) {
                                cell = tableRow.createCell();
                            }
                            cell.setText(String.valueOf(numericValue));
                        }
                        // Hier können weitere Bedingungen für andere Zellentypen hinzugefügt werden
                    }
                }
            }
        }
    }



    public static void saveDocument() {
        try {
            // Write the entire document to the file
            FileOutputStream out = new FileOutputStream(wordFilePath);
            document.write(out);
            out.close();

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

            PDPage pdfPage = new PDPage(new PDRectangle(PDRectangle.A4.getHeight(), PDRectangle.A4.getWidth()));


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
                    drawPdfTable(pdfDocument, tableWidth, yStart, yBottom, table);

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

    private static void drawPdfTable(PDDocument document, float tableWidth, float yStart, float yBottom, XWPFTable table) throws IOException {
        float margin = 20;
        float fontSize = 12;
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

            // Erstellen Sie einen neuen contentStream für die erste Zeile
            if (contentStream == null) {
                PDPage newPage = new PDPage(new PDRectangle(PDRectangle.A4.getHeight(), PDRectangle.A4.getWidth()));
                document.addPage(newPage);
                contentStream = new PDPageContentStream(document, newPage, true, true);
                contentStream.setFont(PDType1Font.HELVETICA_BOLD, fontSize);
                yPosition = yStart;
            }

            for (int colIdx = 0; colIdx < wordCells.size(); colIdx++) {
                XWPFTableCell wordCell = wordCells.get(colIdx);

                float width = tableWidth / (float) wordCells.size();
                float height = calculateCellHeight(wordCell, fontSize);
                String text = wordCell.getText();

                // Manuelle Aufteilung von langen Wörtern bei Überlappung
                String[] words = text.split("\\s+");
                float currentWidth = 0;

                for (String word : words) {
                    float wordWidth = PDType1Font.HELVETICA_BOLD.getStringWidth(word) / 990 * fontSize;

                    if (currentWidth + wordWidth > width && currentWidth > 0) {
                        yPosition -= maxHeight; // Neue Zeile
                        currentWidth = 0;
                    }

                    // Prüfen, ob eine neue Seite benötigt wird
                    if (yPosition - height < yBottom) {
                        // Neue Seite erstellen und zum Dokument hinzufügen
                        PDPage newPage = new PDPage(new PDRectangle(PDRectangle.A4.getHeight(), PDRectangle.A4.getWidth()));
                        document.addPage(newPage);

                        // Schließen Sie den vorherigen contentStream und erstellen Sie einen neuen für die neue Seite
                        contentStream.close();
                        contentStream = new PDPageContentStream(document, newPage, true, true);
                        contentStream.setFont(PDType1Font.HELVETICA_BOLD, fontSize);
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

            // Aktualisieren Sie die y-Position für die nächste Zeile
            yPosition -= maxHeight;

            // Überprüfen, ob eine neue Seite benötigt wird (falls die Zeile nicht vollständig auf die aktuelle Seite passt)
            if (yPosition < yBottom && rowIdx + 1 < table.getRows().size()) {
                // Neue Seite erstellen und zum Dokument hinzufügen
                PDPage newPage = new PDPage(PDRectangle.A4);
                document.addPage(newPage);

                // Schließen Sie den vorherigen contentStream und erstellen Sie einen neuen für die neue Seite
                contentStream.close();
                contentStream = new PDPageContentStream(document, newPage, true, true);
                contentStream.setFont(PDType1Font.HELVETICA_BOLD, fontSize);
                yPosition = yStart;
            } else {
                rowIdx++;
            }
        }

        // Schließen Sie den contentStream am Ende der Methode
        if (contentStream != null) {
            contentStream.close();
        }
    }






    private static float calculateCellHeight(XWPFTableCell cell, float fontSize) {
        // You may need to adjust this calculation based on your specific requirements
        // This is a simple estimation, and the actual cell height may depend on various factors.
        float lineSpacing = 1.5f; // Adjust as needed
        int numberOfLines = cell.getText().split("\\r?\\n").length;
        return numberOfLines * fontSize * lineSpacing;
    }

    static public void openNewJavaFXWindow() {
        Stage newStage = new Stage();

        // Button für die Konvertierung in Word im neuen Fenster
        javafx.scene.control.Button convertToWordButton = new javafx.scene.control.Button("Konvertierung in eine PDF");
        convertToWordButton.setOnAction(e -> ExcelToWordConverter.convertWordToPDF());

        // Layout für das neue Fenster
        VBox newRoot = new VBox(10);
        newRoot.setAlignment(Pos.CENTER);
        newRoot.getChildren().add(convertToWordButton);

        Scene newScene = new Scene(newRoot, 300, 200);
        newStage.setTitle("PDF Konvertierung");
        newStage.setScene(newScene);
        newStage.show();
    }

}
