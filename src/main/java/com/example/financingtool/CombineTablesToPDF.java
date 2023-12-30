package com.example.financingtool;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class CombineTablesToPDF {

    public static void main(String[] args) {
        try {
            PDDocument document = new PDDocument();
            PDPage page = new PDPage();
            document.addPage(page);

            PDPageContentStream contentStream = new PDPageContentStream(document, page);

            // Lese die Daten aus dem ersten Excel-Blatt ("Basisinformation") ein
            readExcelData("Basisinformation", contentStream,page, document);

            // Füge eine neue Seite hinzu
            document.addPage(new PDPage());

            //-n füge ein bild hinzu:
            PDImageXObject image = PDImageXObject.createFromFile("src\\main\\resources\\com\\example\\financingtool\\tree.jpg", document);
            contentStream.drawImage(image, 100, 100);


            // Lese die Daten aus dem zweiten Excel-Blatt ("Gesamtinvestitionskosten") ein
            readExcelData("Gesamtinvestitionskosten", contentStream,page,document);

            // Schließe den Content Stream und das Dokument
            contentStream.close();
            document.save("src\\main\\resources\\com\\example\\financingtool\\combined_tables.pdf");
            document.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    static void readExcelData(String sheetName, PDPageContentStream contentStream, PDPage page, PDDocument document) throws IOException {
        String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
        FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        Sheet sheet = workbook.getSheet(sheetName);

        float margin = 50;
        float yStart = page.getMediaBox().getHeight() - margin;
        float tableWidth = page.getMediaBox().getWidth() - 2 * margin;
        float yPosition = yStart;
        float tableHeight = 100f;

        float bottomMargin = 70f;

        contentStream.beginText();
        contentStream.setFont(PDType1Font.HELVETICA_BOLD, 12);
        contentStream.newLineAtOffset(margin, yStart);
        contentStream.showText(sheetName);
        contentStream.newLineAtOffset(0, -20); // Abstand zwischen Titel und Tabelle
        contentStream.endText();

        drawTable(page, contentStream, yStart, tableWidth, tableHeight, sheet, document);

        contentStream.beginText();
        contentStream.setFont(PDType1Font.HELVETICA_BOLD, 12);
        contentStream.newLineAtOffset(margin, yPosition - tableHeight - bottomMargin);
        contentStream.showText("Ende der Tabelle");
        contentStream.endText();
    }

    private static void drawTable(PDPage page, PDPageContentStream contentStream, float yStart,
                                  float tableWidth, float tableHeight, Sheet sheet, PDDocument document) throws IOException {
        float yPosition = yStart;
        float margin = 50;
        float yStartNewPage = page.getMediaBox().getHeight() - margin;
        float bottomMargin = 70;
        float yPositionNewPage = yStartNewPage - 20;

        int rowsPerPage = 20; // Anzahl der Zeilen pro Seite
        int numberOfRows = sheet.getPhysicalNumberOfRows();
        int numberOfPages = numberOfRows / rowsPerPage;

        for (int i = 0; i <= numberOfPages; i++) {
            drawTableHeader(contentStream, margin, yPosition, tableWidth, tableHeight);
            drawTableContent(contentStream, margin, yPosition, tableWidth, tableHeight, sheet, i, rowsPerPage);
            drawTableFooter(contentStream, page, yPositionNewPage, tableWidth, bottomMargin);

            yPositionNewPage -= tableHeight + bottomMargin;
            yPosition = yPositionNewPage;
            contentStream.close();

            if (i < numberOfPages - 1) {
                PDPage newPage = new PDPage();
                document.addPage(newPage);
                contentStream = new PDPageContentStream(document, newPage);
                yPositionNewPage = newPage.getMediaBox().getHeight() - margin;
            }
        }
    }

    private static void drawTableHeader(PDPageContentStream contentStream, float margin,
                                        float yPosition, float tableWidth, float tableHeight) throws IOException {
        contentStream.setLineWidth(1f);
        contentStream.moveTo(margin, yPosition);
        contentStream.lineTo(margin + tableWidth, yPosition);
        contentStream.stroke();
    }

    private static void drawTableContent(PDPageContentStream contentStream, float margin,
                                         float yPosition, float tableWidth, float tableHeight, Sheet sheet,
                                         int currentPage, int rowsPerPage) throws IOException {
        float yStart = yPosition - tableHeight;
        float yPositionTable = yStart;
        float xPosition = margin;

        int startRow = currentPage * rowsPerPage;
        int endRow = Math.min((currentPage + 1) * rowsPerPage, sheet.getPhysicalNumberOfRows());

        for (int i = startRow; i < endRow; i++) {
            Row row = sheet.getRow(i);
            float rowHeight = calculateRowHeight(row);

            drawTableContentRow(contentStream, xPosition, yPositionTable, tableWidth, rowHeight, row);
            yPositionTable -= rowHeight;
        }
    }

    private static void drawTableContentRow(PDPageContentStream contentStream, float xPosition,
                                            float yPosition, float tableWidth, float rowHeight, Row row) throws IOException {
        contentStream.setLineWidth(1f);
        contentStream.moveTo(xPosition, yPosition);
        contentStream.lineTo(xPosition + tableWidth, yPosition);
        contentStream.stroke();

        float cellMargin = 2f;
        float tableHeightMargin = rowHeight - cellMargin;

        for (Cell cell : row) {
            float cellWidth = (tableWidth / (float) row.getPhysicalNumberOfCells()) - cellMargin;

            contentStream.beginText();
            contentStream.setFont(PDType1Font.HELVETICA, 12);
            contentStream.newLineAtOffset(xPosition + cellMargin, yPosition - tableHeightMargin + cellMargin);
            contentStream.showText(cell.toString());
            contentStream.endText();

            xPosition += cellWidth;
        }
    }

    private static void drawTableFooter(PDPageContentStream contentStream, PDPage page, float yPosition,
                                        float tableWidth, float margin) throws IOException {
        contentStream.setLineWidth(1f);
        contentStream.moveTo(margin, yPosition);
        contentStream.lineTo(margin + tableWidth, yPosition);
        contentStream.stroke();
    }

    private static float calculateRowHeight(Row row) {
        float maxHeight = 0;
        for (Cell cell : row) {
            float cellHeight = calculateCellHeight(cell.toString());
            maxHeight = Math.max(maxHeight, cellHeight);
        }
        return maxHeight;
    }

    private static float calculateCellHeight(String text) {
        // Hier könntest du die Höhe basierend auf der Schriftart und -größe berechnen
        // Hier verwende ich eine feste Höhe für die Vereinfachung
        return 20;
    }
}

