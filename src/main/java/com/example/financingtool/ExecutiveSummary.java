
package com.example.financingtool;


import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.PDResources;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.graphics.form.PDFormXObject;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

public class ExecutiveSummary {

    private static String wordFilePath = "src/main/resources/com/example/financingtool/ExecutiveSummary.docx";
    private static String staticText = "Executive Summary \n" +
            "Ankauf der Liegenschaft “Braumüllergasse 21” (EZ 2169, KG 01401) in Form eines Asset Deals mit\n" +
            "einer eigens gegründeten Projektgesellschaft\n" +
            "Kaufpreis: EUR ${kaufpreis},-\n" +
            "Bekanntgabe Bebauungsbestimmungen; sind angefordert; Widmung: ${wio} \n" +
            "Grundstücksgroße ${grundstuecksgroesse} m²\n" +
            "Erzielbare Wohnnutzfläche laut Stararchitekt: 800 m² WNFL zzgl. gew. Außenflächen von 160 m²\n" +
            "Nutzung: Wohnen – ${wohneinheiten} Wohneinheiten mit ${garagenstellplaetze} Garagenstellplätzen\n" +
            "Einzelabverkauf nach BTVG\n" +
            "GIK: EUR ${gik},- (gerundet)\n" +
            "Prognostizierter Verkaufserlös: EUR ${verkaufserloes},- ø Verkaufspreis EUR 10.000,- siehe Marktanalyse\n" +
            "Gewinn: ${gewinn} (gerundet) ROI ${roi}%\n" +
            "Ziel-Baubeginn: ${zielbaubeginn} \n" +
            "Ziel-Fertigstellung: ${zielfertigstellung} ";

    private String kaufpreis;
    private String grundstuecksgroesse;
    private String wohneinheiten;
    private String garagenstellplaetze;
    private String gik;
    private String verkaufserloes;
    private String gewinn;
    private String zielbaubeginn;
    private String zielfertigstellung;
    private String wio;
    private static XWPFDocument document;
    int countFilled=0;
    private String roi;


    public static void initializeDocument() {
        document = new XWPFDocument();


    }
   /* public void setDatenausBas(String kaufpreis, String grundstuecksgroesse, String wohneinheiten, String garagenstellplaetze, String zielbaubeginn, String zielfertigstellung){
        this.kaufpreis=kaufpreis;
        this.grundstuecksgroesse=grundstuecksgroesse;
        this.wohneinheiten=wohneinheiten;
        this.garagenstellplaetze=garagenstellplaetze;
        this.zielbaubeginn=zielbaubeginn;
        this.zielfertigstellung=zielfertigstellung;
        countFilled++;
        System.out.println(countFilled);
        System.out.println("GIK: "+ this.gik+" WIO: "+this.wio+" Basisiinformation: "+ this.kaufpreis);
    }
    public void setDatenausWidmung(String wio){
        this.wio=wio;
        System.out.println(this.wio);
        countFilled++;

        System.out.println(countFilled);
        System.out.println("GIK: "+ this.gik+" WIO: "+this.wio+" Basisiinformation: "+ this.kaufpreis);
    }
    public void setDatenausGIK(String gik){
        this.gik=gik;
        System.out.println(this.gik);
        countFilled++;
        System.out.println(countFilled);
        System.out.println("GIK: "+ this.gik+" WIO: "+this.wio+" Basisiinformation: "+ this.kaufpreis);

    }*/

    public void getGikData() {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Gesamtinvestitionskosten";
            int rowIdx = 13;
            int colIdx = 4;

            // FileInputStream und Workbook hier erstellen
            try (FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
                 Workbook workbook = new XSSFWorkbook(fileInputStream)) {

                Sheet sheet = workbook.getSheet(sheetName);

                Row row = sheet.getRow(rowIdx);
                Cell cell = row.getCell(colIdx);
                String gikCell = Double.toString(cell.getNumericCellValue());
                System.out.println(gikCell);
                this.gik=gikCell;
                //fk.setText(fkCell);
            } catch (NumberFormatException | IOException e) {
                e.printStackTrace();
                //resultLabel.setText("Fehler bei der Aktualisierung.");
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public void getWireData() {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Wirtschaftlichkeitsrechnung";
            int rowIdxGewinn = 20;
            int rowIdxROI=21;
            int rowIdxVerkaufserloes=18;
            int colIdx = 4;

            // FileInputStream und Workbook hier erstellen
            try (FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
                 Workbook workbook = new XSSFWorkbook(fileInputStream)) {

                Sheet sheet = workbook.getSheet(sheetName);

                Row rowGewinn = sheet.getRow(rowIdxGewinn);
                Cell cellGewinn = rowGewinn.getCell(colIdx);
                String gewinnCell =Double.toString(cellGewinn.getNumericCellValue());

                Row rowROI=sheet.getRow(rowIdxROI);
                Cell cellROI=rowROI.getCell(colIdx);
                String roiCell=Double.toString(cellROI.getNumericCellValue());

                Row rowVe=sheet.getRow(rowIdxVerkaufserloes);
                Cell cellVe=rowVe.getCell(colIdx);
                String veCell=Double.toString(cellVe.getNumericCellValue());

                System.out.println(gewinnCell+" "+roiCell+" "+veCell);
                this.gewinn=gewinnCell;
                this.roi=roiCell;
                this.verkaufserloes=veCell;
                //fk.setText(fkCell);
            } catch (NumberFormatException | IOException e) {
                e.printStackTrace();
                //resultLabel.setText("Fehler bei der Aktualisierung.");
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public void getWIODaten(){
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Basisinformation";
            int rowIdx = 3;
            int colIdx = 15;

            // FileInputStream und Workbook hier erstellen
            try (FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
                 Workbook workbook = new XSSFWorkbook(fileInputStream)) {

                Sheet sheet = workbook.getSheet(sheetName);

                Row row = sheet.getRow(rowIdx);
                Cell cell = row.getCell(colIdx);
                String wioCell = cell.getStringCellValue();
                this.wio=wioCell;
                System.out.println(wioCell);

                //fk.setText(fkCell);
            } catch (NumberFormatException | IOException e) {
                e.printStackTrace();
                //resultLabel.setText("Fehler bei der Aktualisierung.");
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
    public void getBasDaten(){
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Basisinformation";
            int rowIdxKaufpreis = 1;
            int colIdx = 1;
            int rowIdxGrundstuecksgroesse=2;
            int rowIDxWohneinheiten=4;
            int rowIDxBaubeginn=9;
            int rowIDxFertigstellung=10;



            // FileInputStream und Workbook hier erstellen
            try (FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
                 Workbook workbook = new XSSFWorkbook(fileInputStream)) {

                Sheet sheet = workbook.getSheet(sheetName);

                Row rowKp = sheet.getRow(rowIdxKaufpreis);
                Row rowGg = sheet.getRow(rowIdxGrundstuecksgroesse);
                Row rowWe = sheet.getRow(rowIDxWohneinheiten);
                Row rowBb = sheet.getRow(rowIDxBaubeginn);
                Row rowFf=sheet.getRow(rowIDxFertigstellung);

                Cell cellKp = rowKp.getCell(colIdx);
                String kaufpreisCell = cellKp.getStringCellValue();

                Cell cellGg = rowGg.getCell(colIdx);
                String ggCell = cellGg.getStringCellValue();

                Cell cellWe=rowWe.getCell(colIdx);
                String weCell=cellWe.getStringCellValue();

                Cell cellBb=rowBb.getCell(colIdx);
                String bbCell=cellBb.getStringCellValue();

                Cell cellFf=rowFf.getCell(colIdx);
                String ffCell=cellFf.getStringCellValue();


                System.out.println(kaufpreisCell);
                System.out.println(ggCell);
                System.out.println(weCell);
                System.out.println(bbCell);
                System.out.println(ffCell);

                this.kaufpreis=kaufpreisCell;
                this.garagenstellplaetze=ggCell;
                this.wohneinheiten=weCell;
                this.zielbaubeginn=bbCell;
                this.zielfertigstellung=ffCell;


                //fk.setText(fkCell);
            } catch (NumberFormatException | IOException e) {
                e.printStackTrace();
                //resultLabel.setText("Fehler bei der Aktualisierung.");
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public void setDaten() {
        Map<String, String> dynamicValues = new HashMap<>();
        dynamicValues.put("kaufpreis", Objects.requireNonNullElse(this.kaufpreis, "kaufpreis"));
        dynamicValues.put("grundstuecksgroesse", Objects.requireNonNullElse(this.grundstuecksgroesse, "grundstuecksgroesse"));
        dynamicValues.put("wohneinheiten", Objects.requireNonNullElse(this.wohneinheiten, "wohneinheiten"));
        dynamicValues.put("garagenstellplaetze", Objects.requireNonNullElse(this.garagenstellplaetze, "garagenstellplaetze"));
        dynamicValues.put("zielbaubeginn", Objects.requireNonNullElse(this.zielbaubeginn, "zielbaubeginn"));
        dynamicValues.put("zielfertigstellung", Objects.requireNonNullElse(this.zielfertigstellung, "zielfertigstellung"));
        dynamicValues.put("wio", Objects.requireNonNullElse(this.wio, "wio"));
        dynamicValues.put("gik", Objects.requireNonNullElse(this.gik, "gik"));
        dynamicValues.put("gewinn", Objects.requireNonNullElse(this.gewinn, "gewinn"));
        dynamicValues.put("roi", Objects.requireNonNullElse(this.roi, "roi"));
        dynamicValues.put("verkaufserloes",Objects.requireNonNullElse(this.verkaufserloes,"verkaufserloes"));


        createDocument(dynamicValues);
    }
    private void createDocument(Map<String,String> dynamicValues){
        // Erstelle die Word-Datei, wenn sie nicht existiert
        createWordFile();

        // Lade das vorhandene oder neu erstellte Word-Dokument
        XWPFDocument document = loadWordFile();

        // Füge Text zur Word-Datei hinzu
        addTextToWord(document, dynamicValues);

        // Speichere das aktualisierte Dokument
        saveWordFile(document);

        // Konvertiere Word zu PDF
        convertWordToPDF();
    }


    private void createWordFile() {
        File wordFile = new File(wordFilePath);
        if (document == null) {
            initializeDocument();
        }

        // Falls die Word-Datei existiert, lösche sie
        if (wordFile.exists()) {
            wordFile.delete();
        }

        try {
            // Erstelle ein neues Word-Dokument
            XWPFDocument newDocument = new XWPFDocument();
            FileOutputStream out = new FileOutputStream(wordFile);
            newDocument.write(out);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private XWPFDocument loadWordFile() {
        try {
            FileInputStream fis = new FileInputStream(wordFilePath);
            return new XWPFDocument(fis);
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }


    private void addTextToWord(XWPFDocument document, Map<String, String> dynamicValues) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();

        for (int i = 0; i < staticText.length(); i++) {
            char c = staticText.charAt(i);
            if (c == '$' && i < staticText.length() - 1 && staticText.charAt(i + 1) == '{') {
                // Wenn ein Platzhalter gefunden wird...
                int placeholderStart = i + 2;  // Der Index nach '{'
                int placeholderEnd = staticText.indexOf('}', placeholderStart);

                if (placeholderEnd != -1) {
                    String placeholder = staticText.substring(placeholderStart, placeholderEnd);

                    // Überprüfen, ob der Platzhalter in dynamicValues vorhanden ist
                    if (dynamicValues.containsKey(placeholder)) {
                        String replacement = dynamicValues.get(placeholder);

                        // Füge den Ersatztext zum Run hinzu
                        run.setText(replacement);

                        // Überspringe den Rest des Platzhaltertextes
                        i = placeholderEnd + 1;
                        continue;
                    }
                }
            } else if (c == '\n') {
                // Wenn ein Zeilenumbruchzeichen gefunden wird, füge einen Absatz hinzu
                paragraph = document.createParagraph();
                run = paragraph.createRun();
                continue;  // Springe zum nächsten Schleifeniteration
            }

            // Andernfalls füge normalen Text hinzu
            run.setText(String.valueOf(c));
        }

    }

    private void saveWordFile(XWPFDocument document) {
        try {
            FileOutputStream out = new FileOutputStream(wordFilePath);
            document.write(out);
            out.close();

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
            pdfDocument.addPage(pdfPage);

            PDPageContentStream contentStream = new PDPageContentStream(pdfDocument, pdfPage);

            try {
                List<XWPFParagraph> paragraphs = document.getParagraphs();
                contentStream.setFont(PDType1Font.HELVETICA_BOLD, 12);
                float yPosition = pdfPage.getMediaBox().getHeight() - 20;

                for (XWPFParagraph paragraph : paragraphs) {
                    String text = paragraph.getText();
                    contentStream.beginText();
                    contentStream.newLineAtOffset(20, yPosition);
                    contentStream.showText(text);
                    contentStream.newLine();
                    contentStream.endText();
                    yPosition -= 12; // Abstand zwischen den Zeilen anpassen
                }

            } finally {
                contentStream.close();
                pdfDocument.save("try2.pdf");
                pdfDocument.close();
                in.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}






