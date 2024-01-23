package com.example.financingtool;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.Font.FontFamily;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfPageEventHelper;
import com.itextpdf.text.pdf.PdfWriter;

import java.io.*;
import java.text.DecimalFormat;
import java.util.Iterator;
//import java.util.Scanner;
//import javax.swing.JFileChooser;
//import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelToPDF {
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

                String logoPath = "src/main/resources/com/example/financingtool/logo.jpg";
                File jpeg = new File(logoPath);
                if (jpeg.exists()) {
                    insertLogo(document, logoPath);
                } else {
                    insertLogo(document, "src/main/resources/com/example/financingtool/logo.png");
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
                    String imagePath = "src/main/resources/com/example/financingtool/standort.png";

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
                    String imagePath = "src/main/resources/com/example/financingtool/adresse.png";

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
                    rowsAndColumns = new int[]{0, 22, 0, 7};
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
                }
                else if  (sheet.getSheetName().equals("Widmung")){
                    columnWidths = new float[]{40f};
                }
                else if (sheet.getSheetName().equals("Gesamtinvestitionskosten")) {
                    columnWidths = new float[]{10f, 3f, 3f, 3f, 3f, 3f};
                }
                else if (sheet.getSheetName().equals("Mittelverwendung - Mittelherkun")){
                    columnWidths = new float[]{10f, 5f, 5f, 10f, 5f, 5f};
                }
                else if (sheet.getSheetName().equals("Wirtschaftlichkeitsrechnung")){
                    columnWidths = new float[]{8f, 5f, 6f, 10f, 6f, 6f, 6f, 6f};
                }
                else {
                    columnWidths = new float[]{5f, 0f, 35f, 7f, 7f, 5f, 15f};
                }

                if(sheet.getSheetName().equals("Lage")) {
                    table.setWidths(columnWidths);
                    table.setTotalWidth(250);
                    table.setHorizontalAlignment(Element.ALIGN_JUSTIFIED);
                    table.setLockedWidth(true);
                } else {
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

                        String cellValue = getStringFormattedCell(dataFormatter, cell);
                        p = new Paragraph(cellValue, title ? new Font(FontFamily.HELVETICA, 18) : normal);
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
                            list.add(new ListItem(cellValue, normal));
                        }
                    }
                } else if (row.getRowNum() >= 1 && row.getRowNum() <= rows) {
                    int j = 0;
                    for (Cell cell : row) {
                        if (rowsAndColumns[2] <= j && j++ <= rowsAndColumns[3]) {
                            String cellValue = getStringFormattedCell(dataFormatter, cell);

                            if (sheet.getSheetName().equals("Basisinformation") && !cellValue.isEmpty() && table != null) {
                                table.addCell(cellValue);
                            } else if (sheet.getSheetName().equals("Lage")) {
                                if(cellValue.equals("zu Fuß")) {
                                    try {
                                        String imagePath = "src/main/resources/com/example/financingtool/fuss.jpg";

                                        insertIcon(table, imagePath, 18);
                                    } catch (IOException e) {
                                        throw new RuntimeException(e);
                                    }
                                } else if(cellValue.equals("mit dem Fahrrad")) {
                                    try {
                                        String imagePath = "src/main/resources/com/example/financingtool/rad.png";

                                        insertIcon(table, imagePath, 30);
                                    } catch (IOException e) {
                                        throw new RuntimeException(e);
                                    }
                                } else {
                                    if(isNumeric(cellValue)) {
                                        table.addCell(cellValue + " Min.");
                                    }else if(cellValue.equals("Schulen")) {
                                        try {
                                            String imagePath = "src/main/resources/com/example/financingtool/schule.jpg";

                                            insertIcon(table, imagePath, 25);
                                        } catch (IOException e) {
                                            throw new RuntimeException(e);
                                        }
                                    }else if(cellValue.equals("Restaurants")) {
                                        try {
                                            String imagePath = "src/main/resources/com/example/financingtool/essen.png";

                                            insertIcon(table, imagePath, 22);
                                        } catch (IOException e) {
                                            throw new RuntimeException(e);
                                        }
                                    }else if(cellValue.contains("Verkehr")) {
                                        try {
                                            String imagePath = "src/main/resources/com/example/financingtool/zug.png";

                                            insertIcon(table, imagePath, 25);
                                        } catch (IOException e) {
                                            throw new RuntimeException(e);
                                        }
                                    }else if(cellValue.contains("handel")) {
                                        try {
                                            String imagePath = "src/main/resources/com/example/financingtool/einkauf.jpg";

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
                                table.addCell(cellValue);
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

                            p = new Paragraph(cellValue, title ? new Font(FontFamily.HELVETICA, 18) : normal);
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

    private String getStringFormattedCell(DataFormatter dataFormatter, Cell cell) {
        String cellValue;
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
        }
        return cellValue;
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

    /**
     *
     * @return
     */
//    public String[] choosefile() {
//        //Choose File to Read
//        JFileChooser fileChooser = new JFileChooser();
//        fileChooser.setDialogTitle("Excel To PDF");
//        //only choose excel file format
//        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel", "xls", "xlsx");
//        fileChooser.setFileFilter(filter);
//        fileChooser.showOpenDialog(null);
//        File selectedfile = fileChooser.getCurrentDirectory();
//        String[] ret = new String[2];
//        ret[0] = fileChooser.getSelectedFile().getPath();
//        ret[1] = selectedfile.getPath();
//        return ret;
//
//    }
}
/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

//import com.itextpdf.text.Document;
//import com.itextpdf.text.DocumentException;
//import com.itextpdf.text.Element;
//import com.itextpdf.text.Font;
//import com.itextpdf.text.Font.FontFamily;
//import com.itextpdf.text.Paragraph;
//import com.itextpdf.text.pdf.PdfPTable;
//import com.itextpdf.text.pdf.PdfWriter;
//import java.io.File;
//import java.io.FileNotFoundException;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.util.Iterator;
//import java.util.Scanner;
////import javax.swing.JFileChooser;
////import javax.swing.filechooser.FileNameExtensionFilter;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.DataFormatter;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.ss.usermodel.WorkbookFactory;
//
///**
// *
// * @author User
// */
//public class ExcelToPDF {
//
//    /**
//     *
//     * @param args
//     * @throws FileNotFoundException
//     * @throws DocumentException
//     * @throws IOException
//     */
//    public static void main(String[] args) throws FileNotFoundException, DocumentException, IOException {
//        String excelpath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
//        ExcelToPDF excel = new ExcelToPDF();
//        String pdfpath = "src/main/resources/com/example/financingtool/tester.pdf";
//
//        excel.createpdf(pdfpath, excelpath);
//        System.out.print("Data succesfully stored at: ");
//        System.out.println(pdfpath);
//    }
//
//    /**
//     *
//     * @param document
//     * @param excelpath
//     * @throws IOException
//     * @throws DocumentException
//     */
//    public void readNwrite(Document document, String excelpath) throws IOException, DocumentException {
//        try (Workbook workbook = WorkbookFactory.create(new File(excelpath))) {
//            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
//            int sheetnum = 0;
//            while (sheetIterator.hasNext()) {
//                Sheet sheet = sheetIterator.next();
//                workbook.getSheetAt(sheetnum);
//                sheetnum++;
//
//                DataFormatter dataFormatter = new DataFormatter();
//
//                PdfPTable table = new PdfPTable(7);
//                Paragraph p;
//                Font normal = new Font(FontFamily.TIMES_ROMAN, 12);
//                boolean title = true;
//
//                for (Row row : sheet) {
//                    if (row.getRowNum() >= 4 && row.getRowNum() <= 155) {
//                        for (Cell cell : row) {
//
//                            String cellValue = dataFormatter.formatCellValue(cell);
//                            table.addCell(cellValue);
//                            System.out.println(cell.getRowIndex() + " " + cell.getColumnIndex());
//                            System.out.print(cellValue + "\t");
//
//                        }
//
//                    } else if (row.getRowNum() < 4) {
//                        for (Cell cell : row) {
//
//                            String cellValue = dataFormatter.formatCellValue(cell);
//
//                            p = new Paragraph(cellValue, title ? normal : normal);
//                            p.setAlignment(Element.ALIGN_JUSTIFIED);
//                            document.add(p);
//                        }
//
//                    }
//                }
//                document.add(new Paragraph(" "));
//                float[] columnWidths = new float[]{5f, 0f, 35f, 7f, 7f, 5f, 15f};
//                table.setWidths(columnWidths);
//                table.setTotalWidth(550);
//                table.setLockedWidth(true);
//                document.add(table);
//                for (Row row : sheet) {
//                    if (row.getRowNum() >154) {
//                        for (Cell cell : row) {
//
//                            String cellValue = dataFormatter.formatCellValue(cell);
//
//                            p = new Paragraph(cellValue, title ? normal : normal);
//                            p.setAlignment(Element.ALIGN_JUSTIFIED);
//                            document.add(p);
//                        }
//
//                    }
//                }
//            }
//        }
//    }
//
//    /**
//     *
//     * @param pdfpath
//     * @param excelpath
//     * @throws DocumentException
//     * @throws FileNotFoundException
//     * @throws IOException
//     */
//    public void createpdf(String pdfpath, String excelpath) throws DocumentException, FileNotFoundException, IOException {
//
//        Document document = new Document();
//        PdfWriter.getInstance(document, new FileOutputStream(pdfpath));
//        document.open();
//        readNwrite(document, excelpath);
//        document.close();
//    }
//
//    /**
//     *
//     * @return
//     */
////    public String[] choosefile() {
////        //Choose File to Read
////        JFileChooser fileChooser = new JFileChooser();
////        fileChooser.setDialogTitle("Excel To PDF");
////        //only choose excel file format
////        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel", "xls", "xlsx");
////        fileChooser.setFileFilter(filter);
////        fileChooser.showOpenDialog(null);
////        File selectedfile = fileChooser.getCurrentDirectory();
////        String[] ret = new String[2];
////        ret[0] = fileChooser.getSelectedFile().getPath();
////        ret[1] = selectedfile.getPath();
////        return ret;
////    }
//}

//import com.spire.xls.FileFormat;
//import com.spire.xls.Workbook;
//
//public class ExcelToPDF {
//    public static void main(String[] args) {
//
//        //Create a Workbook instance
//        Workbook workbook = new Workbook();
//        //Load an Excel file
//        workbook.loadFromFile("src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx");
//
//        //Fit all worksheets on one page (optional)
//        workbook.getConverterSetting().setSheetFitToPage(true);
//
//        //Save the workbook to PDF
//        workbook.saveToFile("src/main/resources/com/example/financingtool/iceTester.pdf", FileFormat.PDF);
//    }
//}

//package com.example.financingtool;
//
//import org.apache.pdfbox.pdmodel.PDDocument;
//import org.apache.pdfbox.pdmodel.PDPage;
//import org.apache.pdfbox.pdmodel.PDPageContentStream;
//import org.apache.pdfbox.pdmodel.common.PDRectangle;
//import org.apache.pdfbox.pdmodel.graphics.image.LosslessFactory;
//import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.apache.poi.xwpf.usermodel.*;
//import org.apache.fop.apps.*;
//import java.io.*;
//import java.util.List;
//import javax.xml.transform.Result;
//import javax.xml.transform.Source;
//import javax.xml.transform.Transformer;
//import javax.xml.transform.TransformerFactory;
//import javax.xml.transform.sax.SAXResult;
//import javax.xml.transform.stream.StreamSource;
//
//public class ExcelToPDF {
//
//    public static void main(String[] args) {
//        try {
//            // Convert Excel to DOCX
//            convertExcelToDocx("src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx", "src/main/resources/com/example/financingtool/tester.docx");
//
//            // Convert DOCX to PDF
//            convertDocxToPdf("src/main/resources/com/example/financingtool/tester.docx", "src/main/resources/com/example/financingtool/tester.pdf");
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }
//
//    private static void convertExcelToDocx(String excelFilePath, String docxFilePath) throws Exception {
//        Workbook workbook = WorkbookFactory.create(new File(excelFilePath));
//        XWPFDocument document = new XWPFDocument();
//
//        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
//            Sheet sheet = workbook.getSheetAt(i);
//            for (Row row : sheet) {
//                XWPFParagraph paragraph = document.createParagraph();
//                for (Cell cell : row) {
//                    XWPFRun run = paragraph.createRun();
//                    if(cell.getCellType() == CellType.NUMERIC)
//                        run.setText(String.valueOf(cell.getNumericCellValue()));
//                    else if (cell.getCellType() == CellType.STRING)
//                        run.setText(cell.getStringCellValue());
//                }
//            }
//        }
//
//        try (FileOutputStream out = new FileOutputStream(new File(docxFilePath))) {
//            document.write(out);
//        }
//    }
//    private static void convertDocxToPdf(String wordFilePath, String pdfFilePath) throws Exception {
//        try (FileInputStream fis = new FileInputStream(wordFilePath);
//             XWPFDocument document = new XWPFDocument(fis);
//             PDDocument pdfDocument = new PDDocument()) {
//
//            for (XWPFPictureData picture : document.getAllPictures()) {
//                XWPFPictureData pictureData = picture;
//                PDPage pdfPage = new PDPage(PDRectangle.A4);
//                pdfDocument.addPage(pdfPage);
//                try (PDPageContentStream contentStream = new PDPageContentStream(pdfDocument, pdfPage)) {
//                    contentStream.drawImage(LosslessFactory.createFromImage(pdfDocument, pictureData.getData()), 50, 600);
//                }
//            }
//
//            pdfDocument.save(new FileOutputStream(pdfFilePath));
//        }
//    }

//    private static void convertDocxToPdf(String docxFilePath, String pdfFilePath) throws Exception {
//        // Configure FOP
//        FopFactory fopFactory = FopFactory.newInstance(new File(".").toURI());
//        FOUserAgent foUserAgent = fopFactory.newFOUserAgent();
//        foUserAgent.getRendererOptions().put("pdf-a-mode", "PDF/A-1b");
//
//        // Create output stream for PDF
//        OutputStream out = new BufferedOutputStream(new FileOutputStream(new File(pdfFilePath)));
//
//        // Construct FOP with desired output format
//        Fop fop = fopFactory.newFop(MimeConstants.MIME_PDF, foUserAgent, out);
//
//        // Load DOCX file and apply XSL-FO transformation
//        FileInputStream docxInputStream = new FileInputStream(new File(docxFilePath));
//        TransformerFactory transformerFactory = TransformerFactory.newInstance();
//        Transformer transformer = transformerFactory.newTransformer();
//        transformer.setParameter("versionParam", "1.0");
//        Source xslt = new StreamSource(new ByteArrayInputStream(
//                ("<xsl:stylesheet version=\"1.0\" xmlns:xsl=\"http://www.w3.org/1999/XSL/Transform\">"
//                        + "<xsl:template match=\"/\">"
//                        + "  <fo:root xmlns:fo=\"http://www.w3.org/1999/XSL/Format\">"
//                        + "    <fo:layout-master-set>"
//                        + "      <fo:simple-page-master master-name=\"A4\" margin=\"2cm\">"
//                        + "        <fo:region-body/>"
//                        + "      </fo:simple-page-master>"
//                        + "    </fo:layout-master-set>"
//                        + "    <fo:page-sequence master-reference=\"A4\">"
//                        + "      <fo:flow flow-name=\"xsl-region-body\">"
//                        + "        <fo:block>"
//                        + "          <fo:external-graphic src=\"docx:" + docxFilePath + "\" content-type=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document\"/>"
//                        + "        </fo:block>"
//                        + "      </fo:flow>"
//                        + "    </fo:page-sequence>"
//                        + "  </fo:root>"
//                        + "</xsl:template>"
//                        + "</xsl:stylesheet>").getBytes()));
//
//        // Transform DOCX to PDF
//        Result res = new SAXResult(fop.getDefaultHandler());
//        transformer.transform(new StreamSource(docxInputStream), res);
//
//        // Close output streams
//        out.close();
//        docxInputStream.close();
//    }
//}