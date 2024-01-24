package com.example.financingtool;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.Font.FontFamily;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.util.Iterator;
//import java.util.Scanner;
//import javax.swing.JFileChooser;
//import javax.swing.filechooser.FileNameExtensionFilter;

import javafx.application.Application;
import javafx.application.HostServices;
import javafx.stage.Stage;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelToPDF extends Application {
    public static void main(String[] args) throws DocumentException, IOException {
        String excelpath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
        ExcelToPDF excel = new ExcelToPDF();
        String pdfpath = "src/main/resources/com/example/financingtool/tester.pdf";

        excel.createpdf(pdfpath, excelpath);
        System.out.print("Data succesfully stored at: ");
        System.out.println(pdfpath);
    }

    /**
     * @param document
     * @param excelpath
     * @throws IOException
     * @throws DocumentException
     */
    public void readNwrite(Document document, String excelpath) throws IOException, DocumentException {
        try (Workbook workbook = WorkbookFactory.create(new File(excelpath))) {
            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            int sheetnum = 0;
            while (sheetIterator.hasNext()) {
                document.setPageSize(PageSize.A4.rotate());

                Sheet sheet = sheetIterator.next();
                if (sheet.getSheetName().equals("Basisinformation") ||
                        sheet.getSheetName().equals("Standort") ||
                        sheet.getSheetName().equals("GIK_Kalkulation") ||
                        sheet.getSheetName().equals("WIRE_Kalkulation") ||
                        sheet.getSheetName().equals("MVMH1") ||
                        sheet.getSheetName().equals("MVMH2") ||
                        sheet.getSheetName().equals("MVMH3") ||
                        sheet.getSheetName().equals("MVMH4") ||
                        sheet.getSheetName().equals("MVMH5") ||
                        sheet.getSheetName().equals("Restliche Fragen")) {
                    do {
                        sheet = sheetIterator.next();
                        if(!sheetIterator.hasNext())
                            return;
                    } while(sheet.getSheetName().equals("Basisinformation") ||
                            sheet.getSheetName().equals("Standort") ||
                            sheet.getSheetName().equals("GIK_Kalkulation") ||
                            sheet.getSheetName().equals("WIRE_Kalkulation") ||
                            sheet.getSheetName().equals("MVMH1") ||
                            sheet.getSheetName().equals("MVMH2") ||
                            sheet.getSheetName().equals("MVMH3") ||
                            sheet.getSheetName().equals("MVMH4") ||
                            sheet.getSheetName().equals("MVMH5"));
                }
                workbook.getSheetAt(sheetnum);
                sheetnum++;
                document.newPage();

                DataFormatter dataFormatter = new DataFormatter();
                Paragraph p = new Paragraph();
                // Füge das Logo oben rechts auf jeder Seite ein

                String logoPath = "src/main/resources/com/example/financingtool/images/logo.jpg";
                File jpeg = new File(logoPath);
                if (jpeg.exists()) {
                    insertLogo(document, logoPath);
                } else {
                    insertLogo(document, "src/main/resources/com/example/financingtool/images/logo.png");
                }



                PdfPTable table = null;
                int rows = 155;

                if (sheet.getSheetName().equals("Gesamtinvestitionskosten")) {
                    rows = 15;
                }

                int[] rowsAndColumns = new int[4];
                boolean mvmh = false;
                // Neue Seite vor jedem neuen Blatt - muss das sein??
                // document.newPage();

                //Absätze nach Logo
                // Füge fünf Absätze nach dem Logo ein

                if (sheet.getSheetName().equals("Executive Summary")) {
                    System.out.println("Executive Summary");
                    for (int i = 0; i < 5; i++) {
                        Paragraph emptyParagraph = new Paragraph(" ");
                        emptyParagraph.setSpacingAfter(12f); // Setze den Abstand nach dem Absatz
                        document.add(emptyParagraph);
                    }

                    table = new PdfPTable(1);
                    System.out.println("Sheet: " + sheet);
                    rowsAndColumns = new int[]{0, 11, 0, 0};
                } else if (sheet.getSheetName().equals("Lage")) {
                    for (int i = 0; i < 3; i++) {
                        Paragraph emptyParagraph = new Paragraph(" ");
                        emptyParagraph.setSpacingAfter(12f); // Setze den Abstand nach dem Absatz
                        document.add(emptyParagraph);
                    }
                    String imagePath = "src/main/resources/com/example/financingtool/images/standort.png";

                    insertImage(document, imagePath, 250, 150, 400);
                    // Neue Seite vor Stammdaten
                    //  document.newPage();
                    table = new PdfPTable(5);
                    rowsAndColumns = new int[]{0, 11, 0, 4};
                }  else if (sheet.getSheetName().equals("Widmung")) {
                    for (int i = 0; i < 5; i++) {
                        Paragraph emptyParagraph = new Paragraph(" ");
                        emptyParagraph.setSpacingAfter(12f); // Setze den Abstand nach dem Absatz
                        document.add(emptyParagraph);
                    }
                    String imagePath = "src/main/resources/com/example/financingtool/images/adresse.png";

                    insertImage(document, imagePath, 400, 100, 400);
                    // Neue Seite vor Stammdaten
                    //  document.newPage();
                    table = new PdfPTable(1);
                    rowsAndColumns = new int[]{0, 8, 0, 0};
                } else if (sheet.getSheetName().equals("Gesamtinvestitionskosten")) {
                    for (int i = 0; i < 5; i++) {
                        Paragraph emptyParagraph = new Paragraph(" ");
                        emptyParagraph.setSpacingAfter(12f); // Setze den Abstand nach dem Absatz
                        document.add(emptyParagraph);
                    }
                    table = new PdfPTable(6); //Wieso stand da vorher 6?
                    rowsAndColumns = new int[]{0, 14, 0, 5};
                } else if (sheet.getSheetName().equals("Mittelverwendung - Mittelherkun")){
                    for (int i = 0; i < 5; i++) {
                        Paragraph emptyParagraph = new Paragraph(" ");
                        emptyParagraph.setSpacingAfter(12f); // Setze den Abstand nach dem Absatz
                        document.add(emptyParagraph);
                    }
                    table = new PdfPTable(6);
                    int tranche = (int) getTranche();
                    switch (tranche) {
                        case 1:
                            rowsAndColumns = new int[]{0, 6, 0, 5};
                            sheet = workbook.getSheet("MVMH1");
                            mvmh = true;
                            break;
                        case 2:
                            rowsAndColumns = new int[]{0, 7, 0, 5};
                            sheet = workbook.getSheet("MVMH2");
                            mvmh = true;
                            break;
                        case 3:
                            rowsAndColumns = new int[]{0, 8, 0, 5};
                            sheet = workbook.getSheet("MVMH3");
                            mvmh = true;
                            break;
                        case 4:
                            rowsAndColumns = new int[]{0, 9, 0, 5};
                            sheet = workbook.getSheet("MVMH4");
                            mvmh = true;
                            break;
                        case 5:
                            rowsAndColumns = new int[]{0, 10, 0, 5};
                            sheet = workbook.getSheet("MVMH5");
                            mvmh = true;
                            break;
                    }
                } else if (sheet.getSheetName().equals("Wirtschaftlichkeitsrechnung")){
                    for (int i = 0; i < 3; i++) {
                        Paragraph emptyParagraph = new Paragraph(" ");
                        emptyParagraph.setSpacingAfter(12f); // Setze den Abstand nach dem Absatz
                        document.add(emptyParagraph);
                    }
                    table = new PdfPTable(8);
                    rowsAndColumns = new int[]{0, 27, 0, 7};
                }
//                else if (!sheet.getSheetName().equals("Mittelverwendung - Mittelherkun")){
//                    //wird nicht geprinted.
//                    return;
//                }

                else {
                    for (int i = 0; i < 5; i++) {
                        Paragraph emptyParagraph = new Paragraph(" ");
                        emptyParagraph.setSpacingAfter(8f); // Setze den Abstand nach dem Absatz
                        document.add(emptyParagraph);
                    }
                    table = new PdfPTable(7);
                }


                Font normal = new Font(FontFamily.HELVETICA, 14);
                boolean title = true;

                printSheet(document, sheet, rows, dataFormatter, table, title, normal, rowsAndColumns);

                document.add(new Paragraph(" "));

                if(mvmh) {
                    if (sheet.getSheetName().equals("MVMH1") ||
                            sheet.getSheetName().equals("MVMH2")||
                            sheet.getSheetName().equals("MVMH3")||
                            sheet.getSheetName().equals("MVMH4")||
                            sheet.getSheetName().equals("MVMH5")) {
                        sheet = workbook.getSheet("Mittelverwendung - Mittelherkun");
                    }
                }
                float[] columnWidths;
                if (sheet.getSheetName().equals("Executive Summary")) {
                    columnWidths = new float[]{30f};
                }
                else if  (sheet.getSheetName().equals("Lage")){
                    columnWidths = new float[]{20f, 10f, 10f, 10f, 10f};
                    table.setWidths(columnWidths);
                    table.setTotalWidth(290);
                    table.setHorizontalAlignment(Element.ALIGN_JUSTIFIED);
                    table.setLockedWidth(true);
                }
                else if  (sheet.getSheetName().equals("Widmung")){
                    columnWidths = new float[]{40f};
                }
                else if (sheet.getSheetName().equals("Gesamtinvestitionskosten")) {
                    columnWidths = new float[]{9f, 3f, 3f, 3f, 3f, 3f};
                }
                else if (sheet.getSheetName().equals("Mittelverwendung - Mittelherkun")){
                    columnWidths = new float[]{10f, 5f, 5f, 10f, 5f, 5f};
                }
                else if (sheet.getSheetName().equals("Wirtschaftlichkeitsrechnung")){
                    columnWidths = new float[]{10f, 6f, 3f, 4f, 6f, 6f, 6f, 6f};
                    table.setWidths(columnWidths);
                    table.setTotalWidth(800);
                    table.setLockedWidth(true);
                }
                else {
                    columnWidths = new float[]{5f, 0f, 35f, 7f, 7f, 5f, 15f};
                }

                if(!sheet.getSheetName().equals("Lage") && !sheet.getSheetName().equals("Wirtschaftlichkeitsrechnung")) {
                    table.setWidths(columnWidths);
                    table.setTotalWidth(650);
                    table.setLockedWidth(true);
                }
                document.add(table);

                for (Row row : sheet) {
                    if (row.getRowNum() > 154) {
                        for (Cell cell : row) {
                            String cellValue = dataFormatter.formatCellValue(cell);

                            p = new Paragraph(cellValue, title ? normal : normal);
                            p.setAlignment(Element.ALIGN_JUSTIFIED);
                            document.add(p);
                        }
                    }
                }
            }
        }
    }

    private int[] calculateRatio(String imagePath, int maxDim) {
        try {

            Image image = Image.getInstance(imagePath);

            float originalHeight = image.getHeight();
            float originalWidth = image.getWidth();
            float maxDimension = maxDim;
            float newWidth, newHeight;
            if (originalWidth > originalHeight) {
                if(imagePath.contains("standort.png"))
                    maxDimension = 490;
                newWidth = maxDimension;
                newHeight = (int) Math.round((double) originalHeight / originalWidth * maxDimension);
            } else {
                newWidth = (int) Math.round((double) originalWidth / originalHeight * maxDimension);
                newHeight = maxDimension;
            }
            return new int[] {(int) newWidth, (int) newHeight};
        } catch (IOException | DocumentException e) {
            e.printStackTrace();
        }
        return null;
    }

    private void printSheet(Document document, Sheet sheet, int rows, DataFormatter dataFormatter, PdfPTable table, boolean title, Font normal, int[] rowsAndColumns) throws DocumentException {
        Paragraph p;
        int i = 0;
        int n = 0;
        List list = new List();
        list.setSymbolIndent(12);
        list.setListSymbol("\u2022 ");
        list.setIndentationLeft(10f);
        for (Row row : sheet) {
            if(rowsAndColumns[0] <= i && i++ <= rowsAndColumns[1]) {
                if(sheet.getSheetName().equals("Widmung") || sheet.getSheetName().equals("Executive Summary")  || (sheet.getSheetName().equals("Lage") && i < 10)) {
                    for (Cell cell : row) {

                        String[] res = getStringFormattedCell(dataFormatter, cell);
                        String cellValue = res[0];
                        p = new Paragraph(cellValue, title ? new Font(FontFamily.HELVETICA, 18, Font.BOLD) : normal);
                        title = false;
                        p.setAlignment(Element.ALIGN_JUSTIFIED);
                        if(n++ == 0) {
                            document.add(p);
                            Paragraph p2 = new Paragraph(" ");
                            p2.setSpacingAfter(5f);
                            document.add(p2);
                        } else {
                            if(sheet.getSheetName().equals("Lage")) {
                                list.setIndentationLeft(0);
                                String[] arrayStrings = cellValue.split(" ");
                                StringBuilder sBuffer = new StringBuilder();
                                String tempString = "";
                                String tempStringEarlier = "";
                                for(String eachWord : arrayStrings){
                                    tempStringEarlier = tempString;
                                    tempString = tempString + eachWord + " ";
                                    if(tempString.length() >= 45) {
                                        sBuffer.append(tempStringEarlier+"\n");
                                        tempString = eachWord + " ";
                                        tempStringEarlier = "";
                                    }
                                }
                                sBuffer.append(tempString);
                                cellValue = sBuffer.toString().trim();
                            }
                            if(!cellValue.trim().isEmpty()) {
                                if(sheet.getSheetName().equals("Widmung")) {
                                    String[] arrayStrings = cellValue.split(" ");
                                    StringBuilder sBuffer = new StringBuilder();
                                    String tempString = "";
                                    String tempStringEarlier = "";
                                    for (String eachWord : arrayStrings) {
                                        tempStringEarlier = tempString;
                                        tempString = tempString + eachWord + " ";
                                        if (tempString.length() >= 50) {
                                            sBuffer.append(tempStringEarlier + "\n");
                                            tempString = eachWord + " ";
                                            tempStringEarlier = "";
                                        }
                                    }
                                    sBuffer.append(tempString);
                                    cellValue = sBuffer.toString().trim();
                                }
                                list.add(new ListItem(cellValue, normal));
                            }
                        }
                    }
                } else if (row.getRowNum() >= 1 && row.getRowNum() <= rows) {
                    int j = 0;
                    for (Cell cell : row) {
                        if (rowsAndColumns[2] <= j && j++ <= rowsAndColumns[3]) {
                            String[] res = getStringFormattedCell(dataFormatter, cell);
                            String cellValue = res[0];
                            if (sheet.getSheetName().equals("Basisinformation") && !cellValue.isEmpty() && table != null) {
                                table.addCell(cellValue);
                            } else if (sheet.getSheetName().equals("Lage")) {
                                if(cellValue.equals("zu Fuß")) {
                                    try {
                                        String imagePath = "src/main/resources/com/example/financingtool/icons/fuss.jpg";

                                        insertIcon(table, imagePath, 18);
                                    } catch (IOException e) {
                                        throw new RuntimeException(e);
                                    }
                                } else if(cellValue.equals("mit dem Fahrrad")) {
                                    try {
                                        String imagePath = "src/main/resources/com/example/financingtool/icons/rad.png";

                                        insertIcon(table, imagePath, 30);
                                    } catch (IOException e) {
                                        throw new RuntimeException(e);
                                    }
                                } else {
                                    if(isNumeric(cellValue)) {
                                        table.addCell(cellValue + " Min.");
                                    }else if(cellValue.equals("Schulen")) {
                                        try {
                                            String imagePath = "src/main/resources/com/example/financingtool/icons/schule.jpg";

                                            insertIcon(table, imagePath, 25);
                                        } catch (IOException e) {
                                            throw new RuntimeException(e);
                                        }
                                    }else if(cellValue.equals("Restaurants")) {
                                        try {
                                            String imagePath = "src/main/resources/com/example/financingtool/icons/essen.png";

                                            insertIcon(table, imagePath, 22);
                                        } catch (IOException e) {
                                            throw new RuntimeException(e);
                                        }
                                    }else if(cellValue.contains("Verkehr")) {
                                        try {
                                            String imagePath = "src/main/resources/com/example/financingtool/icons/zug.png";

                                            insertIcon(table, imagePath, 25);
                                        } catch (IOException e) {
                                            throw new RuntimeException(e);
                                        }
                                    }else if(cellValue.contains("handel")) {
                                        try {
                                            String imagePath = "src/main/resources/com/example/financingtool/icons/einkauf.jpg";

                                            insertIcon(table, imagePath, 25);
                                        } catch (IOException e) {
                                            throw new RuntimeException(e);
                                        }
                                    } else {
                                        PdfPCell pcell = new PdfPCell(new Phrase(cellValue));
                                        pcell.setBorder(0);
                                        pcell.setFixedHeight(30);
                                        pcell.setHorizontalAlignment(Element.ALIGN_CENTER);
                                        table.addCell(pcell);
                                    }
                                }
                            }else if (!sheet.getSheetName().equals("Basisinformation")) {
                                PdfPCell pcell = new PdfPCell(new Phrase(cellValue));
                                if(i == 2 || (sheet.getSheetName().equals("Wirtschaftlichkeitsrechnung") && (i == 10 || i == 20))) {
                                    pcell.setBorder(0);
                                    pcell.setHorizontalAlignment(Element.ALIGN_CENTER);
                                }
                                if(res[1].equals("true")) {
                                    pcell.setBorder(0);
                                } else if(sheet.getSheetName().contains("MVMH")) {
                                    if(isNumericDouble(cellValue)) {
                                        pcell.setHorizontalAlignment(Element.ALIGN_RIGHT);
                                    } else {
                                        pcell.setHorizontalAlignment(Element.ALIGN_JUSTIFIED);
                                    }
                                } else
                                    pcell.setHorizontalAlignment(Element.ALIGN_RIGHT);
                                table.addCell(pcell);
                            }


                            //  System.out.println(cell.getRowIndex() + " " + cell.getColumnIndex());
                            //  System.out.print(cellValue + "\t");
                        }
                    }

                } else if (row.getRowNum() < 2) {
                    int j = 0;
                    for (Cell cell : row) {
                        if (j++ < 1) {

                            String cellValue = dataFormatter.formatCellValue(cell);

                            p = new Paragraph(cellValue, title ? new Font(FontFamily.HELVETICA, 18, Font.BOLD) : normal);
                            title = false;
                            p.setAlignment(Element.ALIGN_JUSTIFIED);
                            document.add(p);
                        }
                    }
                }
            }
        }
        document.add(list);
    }

    private String[] getStringFormattedCell(DataFormatter dataFormatter, Cell cell) {
        String cellValue;
        String check = "false";
        String[] res = new String[2];
        if (cell.getCellType() == CellType.FORMULA) {
            if (cell.getCachedFormulaResultType() == CellType.ERROR) {
                cellValue = "";
            } else {
                if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
                    double numericValue = cell.getNumericCellValue();
                    if (numericValue == (int) numericValue) {
                        // It's an integer
                        DecimalFormat formatter = new DecimalFormat("#,###");
                        cellValue = formatter.format(numericValue);
                    } else {
                        // It's a double
                        DecimalFormat formatter = new DecimalFormat("#,##0.00");
                        cellValue = formatter.format(numericValue);
                    }
                    //cellValue = String.valueOf(cell.getNumericCellValue());
                } else {
                    cellValue = cell.getStringCellValue();
                }
            }
        } else {
            cellValue = dataFormatter.formatCellValue(cell);
            if(!cellValue.trim().isBlank())
                check = "true";
        }
        res[0] = cellValue;
        res[1] = check;
        return res;
    }

    private void insertIcon(PdfPTable table, String imagePath, int maxDim) throws BadElementException, IOException {
        Image image = Image.getInstance(imagePath);
        int[] widHei = calculateRatio(imagePath, maxDim);

        image.scaleAbsolute(widHei[0], widHei[1]);
        PdfPCell imageCell = new PdfPCell(image);
        imageCell.setBorder(0);
        imageCell.setHorizontalAlignment(Element.ALIGN_CENTER);
        table.addCell(imageCell);
    }

    private boolean isNumeric(String strNum) {
        if (strNum == null) {
            return false;
        }
        try {
            int d = Integer.parseInt(strNum);
        } catch (NumberFormatException nfe) {
            return false;
        }
        return true;
    }

    private boolean isNumericDouble(String strNum) {
        if (strNum == null) {
            return false;
        }
        try {
            double d = Double.parseDouble(strNum);
        } catch (NumberFormatException nfe) {
            return false;
        }
        return true;
    }

    private void insertLogo(Document document, String imagePath) {
        try {

            int[] widHei = calculateRatio(imagePath, 100);

            // Füge das Bild oben rechts auf die Seite ein
            Image image = Image.getInstance(imagePath);
            image.scaleAbsolute(widHei[0], widHei[1]);

            image.setAbsolutePosition(680, 450);
            document.add(image);


        } catch (IOException | DocumentException e) {
            e.printStackTrace();
        }
    }


    private void insertImage(Document document, String imagePath, int X, int Y, int maxDim) {
        try {
            int[] widHei = calculateRatio(imagePath, maxDim);

            Image image = Image.getInstance(imagePath);
            image.scaleAbsolute(widHei[0], widHei[1]);

            image.setAbsolutePosition(document.right()-widHei[0]+20, document.bottom()+400-widHei[1]);
            document.add(image);


        } catch (IOException | DocumentException e) {
            e.printStackTrace();
        }
    }


    /**
     *
     * @param pdfpath
     * @param excelpath
     * @throws DocumentException
     * @throws FileNotFoundException
     * @throws IOException
     */
    public void createpdf(String pdfpath, String excelpath) throws DocumentException, FileNotFoundException, IOException {

        Document document = new Document();
        PdfWriter.getInstance(document, new FileOutputStream(pdfpath));
        document.open();
        readNwrite(document, excelpath);
        document.close();
        String file1 = "src/main/resources/com/example/financingtool/Stammblattimg.pdf";
        if (Files.exists(Paths.get(file1))) {
            combinePdfDuo();
        } else {
            combinePdfSolo();
        }
    }

    //pdf mit bilder kombinieren


    public void combinePdfDuo() {

        String file1 = "src/main/resources/com/example/financingtool/Stammblattimg.pdf";
        String file2 = "src/main/resources/com/example/financingtool/tester.pdf";
        String outputFile = "src/main/resources/com/example/financingtool/Financingtool.pdf";
        String outputFile2 = "../Financing Tool.pdf";


        //zweiPDF kombinieren
        try {
            // Laden der ersten PDF-Datei
            PDDocument pdfDocument1 = PDDocument.load(new File(file1));

            // Laden der zweiten PDF-Datei
            PDDocument pdfDocument2 = PDDocument.load(new File(file2));

            // Kopieren aller Seiten von der ersten PDF-Datei zur Ausgabedatei
            for (int i = 0; i < pdfDocument1.getNumberOfPages(); i++) {
                PDPage page = pdfDocument1.getPage(i);
                pdfDocument2.addPage(page);
            }

            // Speichern des Ergebnisses
            pdfDocument2.save(outputFile);
            pdfDocument2.save(outputFile2);
            System.out.println("Erfolgreiche Kombination der pdf's");

            // Schließen der geöffneten Dokumente
            pdfDocument1.close();
            pdfDocument2.close();
            try {
                File file = new File(outputFile2);
                HostServices hostServices = getHostServices();
                hostServices.showDocument(file.getAbsolutePath());
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void combinePdfSolo() {

        String file2 = "src/main/resources/com/example/financingtool/tester.pdf";
        String outputFile = "src/main/resources/com/example/financingtool/Financingtool.pdf";
        String outputFile2 = "../Financing Tool.pdf";

        //zweiPDF kombinieren
        try {

            // Laden der zweiten PDF-Datei
            PDDocument pdfDocument2 = PDDocument.load(new File(file2));

            // Speichern des Ergebnisses
            pdfDocument2.save(outputFile);
            pdfDocument2.save(outputFile2);
            System.out.println("Erfolgreiche Kombination der PDFs");

            // Schließen der geöffneten Dokumente
            pdfDocument2.close();

            try {
                File file = new File(outputFile2);
                HostServices hostServices = getHostServices();
                hostServices.showDocument(file.getAbsolutePath());
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
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

    @Override
    public void start(Stage primaryStage) throws Exception {
        // muss wegen Application eingefügt sein
    }
}