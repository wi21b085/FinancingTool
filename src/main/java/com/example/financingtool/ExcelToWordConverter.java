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
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPrInner;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;


//import java.awt.Color;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ExcelToWordConverter {
    private static String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
    private static String wordFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.docx";
    private static String pdfFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.pdf";
    private static XWPFDocument document;

    static ExecutiveSummary executiveSummary = new ExecutiveSummary();


    public static void setExecutiveSummary(ExecutiveSummary executiveSummary) {
        ExcelToWordConverter.executiveSummary=executiveSummary;
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

            // Export Basisinformation
            exportSheetToWord(workbook, "Basisinformation");

            // Add a newline between the two tables
            document.createParagraph();

            // Export Gesamtinvestitionskosten
            exportSheetToWord(workbook, "Gesamtinvestitionskosten");
            //document.createParagraph();
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
    private static void addTableTitle(String tableName) {
        if (document == null) {
            initializeDocument();
        }

        // Add title to Word document
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setBold(true);
        run.setText("Titel für Tabelle: " + tableName);
        paragraph.setSpacingBefore(10); // You can adjust the spacing as needed
        paragraph.setSpacingAfter(10);

        // Add title to PDF
        addTextToFirstPage("Titel für Tabelle: " + tableName);
    }

    public static double getTranche() {
        double trancheCell = 0;

        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Mittelverwendung - Mittelherkun";
            int rowIdx = 11;
            int colIdx = 7;

            // FileInputStream und Workbook hier erstellen
            try (FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
                 Workbook workbook = new XSSFWorkbook(fileInputStream)) {

                Sheet sheet = workbook.getSheet(sheetName);

                Row row = sheet.getRow(rowIdx);
                Cell cell = row.getCell(colIdx);
                trancheCell =  cell.getNumericCellValue();
                System.out.println(trancheCell);
                //fk.setText(fkCell);

            } catch (NumberFormatException | IOException e) {
                e.printStackTrace();
                //resultLabel.setText("Fehler bei der Aktualisierung.");
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return trancheCell;
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
            /*if(getTranche()==2) {
                System.out.println("MVMH: 2");
                createTable(document, sheet, 0, 5);
            } else if (getTranche()==1) {
                System.out.println("MVMH:1");
                createTable(document,sheet,8,13);
            }else if (getTranche()==3){
                System.out.println("MVMH: 3");
                createTable(document,sheet, 16, 21);
            } else if (getTranche()==4) {
                System.out.println("MVMH: 4");
                createTable(document,sheet,24,29);
            }else if (getTranche()==5){
                System.out.println("MVMH: 5");
                createTable(document,sheet,32,37);
            }else{
                System.out.println("Entschuldigung, etwas ist beim Generieren der Tabelle schiefgegeangen.");
            }*/

            double tranche = getTranche();
            int startRow;
            int endRow;

            switch ((int) tranche) {
                case 2:
                    startRow = 0;
                    endRow= 5;
                    break;
                case 1:
                    startRow = 8;
                    endRow=13;
                    break;
                case 3:
                    startRow = 16;
                    endRow=21;
                    break;
                case 4:
                    startRow = 24;
                    endRow=29;
                    break;
                case 5:
                    startRow = 32;
                    endRow=37;
                    break;
                default:
                    System.out.println("Entschuldigung, etwas ist beim Generieren der Tabelle schiefgegangen.");
                    return; // Hier könntest du weiteren Code hinzufügen oder die Methode verlassen, je nach Bedarf
            }
            addTableTitle("Mittelverwendung und Mittelherkunft");
            createTable(document, sheet, startRow, endRow);


        }

        else if (sheetName.equals("Basisinformation")) {
            // Export columns A-C to Word
            addTableTitle("Basisinformation");
            createTable(document, sheet, 0, 2);

            // Add a newline between the two tables
            document.createParagraph().setPageBreak(true);
            System.out.println("Bas");
            // Export columns H-I to Word
            addTableTitle("Stammdaten");
            createTable(document, sheet, 7, 8);

            System.out.println("Widmung");
            addTableTitle("Widmung");
            createTable(document,sheet,14,14);
            document.createParagraph().setPageBreak(true);
        }else if (sheetName.equals("Gesamtinvestitionskosten")) {
            System.out.println("Ges");
            addTableTitle("Gesamtinvestitionskosten");
            createGIKtable(sheet);
            //document.createParagraph().setPageBreak(true);
            // createTable(document, sheet,0,5);
        }else if(sheetName.equals("Wirtschaftlichkeitsrechnung")){
            document.createParagraph().setPageBreak(true);
            addTableTitle("Wirtschaftlichkeitsrechnung");
            createTable(document, sheet, 0,7);
        }

        // Write Word document to output file
        try {
            document.write(out);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private static void alternateTableColors(XWPFTable table) {
        // Setzen Sie die gewünschte Hintergrundfarbe hier
        String backgroundColor = "CCE5FF";

        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
                CTShd shd = tcPr.isSetShd() ? tcPr.getShd() : tcPr.addNewShd();

                // Setzen Sie die Hintergrundfarbe für jede Zelle
                shd.setFill(backgroundColor);
            }
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

    private static void createTable(XWPFDocument document, Sheet sheet, int startColumn, int endColumn) {
        // Create table in Word
        XWPFTable table = document.createTable();
        alternateTableColors(table);
        XWPFTableRow headerRow = table.getRow(0);

        // Populate table header
        for (int i = startColumn; i <= endColumn; i++) {
            XWPFTableCell cell = headerRow.getCell(i - startColumn);
            if (cell == null) {
                cell = headerRow.createCell();
            }
            cell.setText(sheet.getRow(0).getCell(i).getStringCellValue());

            // Formatieren Sie die erste Zeile als Überschrift
            CTTcPr tcpr = cell.getCTTc().addNewTcPr();
            CTTcPrInner inner = tcpr;
            inner.addNewVAlign().setVal(STVerticalJc.CENTER);
            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

            // Setzen Sie die Schriftgröße der Überschrift
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            XWPFRun run = paragraph.getRuns().get(0);
            run.setFontSize(20); // Hier können Sie die gewünschte Schriftgröße einstellen
            run.setBold(true);   // Fett für die Überschrift
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


    public static void addTextToFirstPage(String text) {
        if (document == null) {
            initializeDocument();
        }

        // Erstellen Sie einen Absatz und einen Run, um den Inhalt hinzuzufügen
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(text);

        // Setzen Sie den Absatz auf die erste Seite
        paragraph.setPageBreak(true);

        // Verschieben Sie den restlichen Inhalt um eine Seite nach unten
        for (int i = document.getParagraphs().size() - 1; i >= 0; i--) {
            XWPFParagraph currentParagraph = document.getParagraphs().get(i);
            if (!currentParagraph.equals(paragraph)) {
                // Verschieben Sie den Absatz auf die nächste Seite
                document.setParagraph(currentParagraph, i + 1);
            }
        }
        saveDocument();
    }






    public static void saveDocument() {
        try {
            // Write the entire document to the file
            FileOutputStream out = new FileOutputStream(wordFilePath);
            document.write(out);
            out.close();

            System.out.println("Data successfully exported from Excel to Word.");
            mergePDFs();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //-- nazia - bilder und rechnungspdf miteinander verbinden.
    public static void mergePDFs(){
        String file1 = "src/main/resources/com/example/financingtool/Stammblattimg.pdf";
        String file2 = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.pdf";
        String outputFile ="src/main/resources/com/example/financingtool/final.pdf";


        try {
            // Laden der ersten PDF-Datei
            PDDocument pdfDocument1 = PDDocument.load(new java.io.File(file1));

            // Laden der zweiten PDF-Datei
            PDDocument pdfDocument2 = PDDocument.load(new java.io.File(file2));

            // Kopieren aller Seiten von der ersten PDF-Datei zur Ausgabedatei
            for (int i = 0; i < pdfDocument1.getNumberOfPages(); i++) {
                PDPage page = pdfDocument1.getPage(i);
                pdfDocument2.addPage(page);
            }

            // Speichern des Ergebnisses
            pdfDocument2.save(outputFile);
            System.out.println("Erfolgreiche Kombination der pdf's");

            // Schließen der geöffneten Dokumente
            pdfDocument1.close();
            pdfDocument2.close();

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
                    contentStream.close();
                    //pdfDocument.save(pdfFilePath);
                    pdfDocument.close();
                    document.close();
                    in.close();
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

        float yPosition = yStart;
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

            float[] columnWidths = getColumnWidths(tableWidth, wordCells.size());

            for (int colIdx = 0; colIdx < wordCells.size(); colIdx++) {
                XWPFTableCell wordCell = wordCells.get(colIdx);

                float width = columnWidths[colIdx];
                float height = calculateCellHeight(wordCell, fontSize);
                String text = wordCell.getText();

                // Zeichnen Sie vertikale Linien für die aktuelle Zelle
                contentStream.moveTo(margin + colIdx * width, yPosition);
                contentStream.lineTo(margin + colIdx * width, yPosition - maxHeight);

                // Überprüfen Sie den gesamten Zellentext
                float cellTextWidth = PDType1Font.HELVETICA_BOLD.getStringWidth(text) / 1000 * fontSize;

                // Überprüfen Sie, ob der Zellentext in die Zelle passt
                if (cellTextWidth > width) {
                    // Wenn der Zellentext nicht in die Zelle passt, zerlegen Sie ihn in mehrere Zeilen
                    List<String> lines = splitTextToFitWidth(text, width, fontSize);

                    for (String line : lines) {
                        // Zeichnen Sie horizontale Linien für die neue Zeile
                        contentStream.moveTo(margin, yPosition - height);
                        contentStream.lineTo(margin + tableWidth, yPosition - height);
                        contentStream.stroke();

                        // Setzen Sie die Schriftgröße für die erste Zeile
                        if (rowIdx == 0 && colIdx == 0) {
                            contentStream.setFont(PDType1Font.HELVETICA_BOLD, 16);
                            //contentStream.setNonStrokingColor(Color.RED); // Ändern Sie die Farbe nach Bedarf
                        } else {
                            contentStream.setFont(PDType1Font.HELVETICA_BOLD, fontSize);
                            //contentStream.setNonStrokingColor(Color.BLACK); // Setzen Sie die Farbe für die restlichen Zellen
                        }

                        // Zeichnen Sie den Zellentext
                        contentStream.beginText();
                        contentStream.newLineAtOffset(margin + colIdx * width + cellMargin, yPosition - height + 3f);
                        contentStream.showText(line);
                        contentStream.endText();

                        // Aktualisieren Sie die y-Position für die nächste Zeile
                        yPosition -= height;
                    }

                    // Aktualisieren Sie die y-Position für die nächste Zelle
                    yPosition += maxHeight;
                } else {
                    // Wenn der Zellentext in die Zelle passt, zeichnen Sie horizontale Linien und den Zellentext
                    contentStream.moveTo(margin, yPosition - height);
                    contentStream.lineTo(margin + tableWidth, yPosition - height);
                    contentStream.stroke();

                    // Setzen Sie die Schriftgröße für die erste Zeile
                    if (rowIdx == 0 && colIdx == 0) {
                        contentStream.setFont(PDType1Font.HELVETICA_BOLD, 16);
                        //contentStream.setNonStrokingColor(Color.RED); // Ändern Sie die Farbe nach Bedarf
                    } else {
                        contentStream.setFont(PDType1Font.HELVETICA_BOLD, fontSize);
                        //contentStream.setNonStrokingColor(Color.BLACK); // Setzen Sie die Farbe für die restlichen Zellen
                    }

                    // Zeichnen Sie den Zellentext
                    contentStream.beginText();
                    contentStream.newLineAtOffset(margin + colIdx * width + cellMargin, yPosition - height + 3f);
                    contentStream.showText(text);
                    contentStream.endText();
                }

                // Aktualisieren Sie die Zellenhöhe
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

    private static float[] getColumnWidths(float tableWidth, int numColumns) {
        float[] columnWidths = new float[numColumns];
        for (int i = 0; i < numColumns; i++) {
            columnWidths[i] = tableWidth / (float) numColumns;
        }
        return columnWidths;
    }

    private static List<String> splitTextToFitWidth(String text, float maxWidth, float fontSize) throws IOException {
        List<String> lines = new ArrayList<>();
        String[] words = text.split("\\s+");
        float currentWidth = 0;
        StringBuilder currentLine = new StringBuilder();

        for (String word : words) {
            float wordWidth = PDType1Font.HELVETICA_BOLD.getStringWidth(word) / 1000 * fontSize;

            if (currentWidth + wordWidth > maxWidth && currentWidth > 0) {
                // Wenn das aktuelle Wort die verbleibende Breite überschreitet, starten Sie eine neue Zeile
                lines.add(currentLine.toString());
                currentLine = new StringBuilder();
                currentWidth = 0;
            }

            // Fügen Sie das aktuelle Wort in die Zeile ein
            currentLine.append(word).append(" ");
            currentWidth += wordWidth;
        }

        // Fügen Sie die letzte Zeile hinzu, wenn vorhanden
        if (currentLine.length() > 0) {
            lines.add(currentLine.toString());
        }

        return lines;
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
        convertToWordButton.setOnAction(e ->handleConvertToWordButtonClick());

        // Layout für das neue Fenster
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
