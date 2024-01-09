package com.example.financingtool;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.StackPane;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

public class StammblattController implements IAllExcelRegisterCards {

    static String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
    static String wordFilePath = "src/main/resources/com/example/financingtool/SEPJ_Stammblatt.docx";
    static String pdfFilePath = "src/main/resources/com/example/financingtool/SEPJ_Stammblatt.pdf";
    //wird das mit dem pdf ein problem werden, wenn es noch nicht existiert?

    //Maria M
    @FXML
    private Button weiterButton;

    //Label für Error Text??
    @FXML
    private Label resultLabel;
    //User inputs
    @FXML
    private TextField firmenname;
    @FXML
    private TextField strasse;

    @FXML
    private TextField plz;

    @FXML
    private TextField ort;

    @FXML
    private TextField schule;

    @FXML
    private TextField oeffi;

    @FXML
    private TextField lage;

    private static final String FILE_NAME = "SEPJ-Rechnungen.xlsx";
    private static final String SHEET_NAME = "Basisinformationen";

    public void onHelloButtonClick(ActionEvent actionEvent) {
    }

    //firmennamen in das excel eintragen
    public void submit() {
        //Submit nur möglich, wenn alle Felder befüllt.

        //leere Eingabe
        if (firmenname.getText().isEmpty() || strasse.getText().isEmpty() || plz.getText().isEmpty()
                || ort.getText().isEmpty() || lage.getText().isEmpty() || schule.getText().isEmpty() || oeffi.getText().isEmpty()) {
            resultLabel.setText("Daten unvollständig");
            return;
        }
        //ungültige eingabe
        else if (IAllExcelRegisterCards.isNumericStr(firmenname.getText()) ||
                IAllExcelRegisterCards.isNumericStr(strasse.getText()) ||
                IAllExcelRegisterCards.isNumericStr(ort.getText()) ||
                IAllExcelRegisterCards.isNumericStr(lage.getText()) ||
                IAllExcelRegisterCards.isNumericStr(oeffi.getText()) ||
                !IAllExcelRegisterCards.isNumericStr(plz.getText()) ||
                !IAllExcelRegisterCards.isNumericStr(schule.getText())
        ) {
            resultLabel.setText("Achtung, bitte geben Sie gültige Daten an");
            return;
        }
        //richtige Werte
        else {
            //    boolean str = IAllExcelRegisterCards.isNumericStr(firmenname.getText());

            String[] newValue = new String[7];
            newValue[0] = firmenname.getText();
            newValue[1] = strasse.getText();
            newValue[2] = plz.getText();
            newValue[3] = ort.getText();
            newValue[4] = schule.getText();
            newValue[5] = lage.getText();
            newValue[6] = oeffi.getText();
            writeToExcel(newValue);
            resultLabel.setText("Werte erfolgreich eingefügt.");
        }
    }

    public void writeToExcel(String[] newValue) {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Basisinformation";

            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheet(sheetName);

            // Überprüfen Sie, ob die Zeichenkette nicht leer ist und nicht null ist, bevor Sie sie parsen
            if (newValue != null && !newValue[0].isEmpty()) {
                updateCellValue(sheet, 1, 8, newValue[0]); //Reihenfolge: firmenname, strasse, plz, ort, schule, oeffi
                updateCellValue(sheet, 2, 8, newValue[1]);
                updateCellValue(sheet, 3, 8, newValue[2]);
                updateCellValue(sheet, 4, 8, newValue[3]);
                updateCellValue(sheet, 5, 8, newValue[4]);
                updateCellValue(sheet, 6, 8, newValue[5]);
                updateCellValue(sheet, 8, 8, newValue[6]);
                //kommt in die Zelle 7, 7
            } else {

                System.out.println("Daten unvollständig");
            }

            // Automatische Auswertung der Formeln im gesamten Arbeitsblatt
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            fileInputStream.close();

            FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();

            //   resultLabel.setText("Zelle erfolgreich aktualisiert.");
            System.out.println("Zelle erfolgreich eingefügt.");

        } catch (NumberFormatException | IOException e) {
            e.printStackTrace();
            //  resultLabel.setText("Fehler bei der Aktualisierung von Zelle D" + 6);
        }
    }

    //Kommentar Maria M: updateCellValue in Updateklasse hier aufrufen
    private static void updateCellValue(Sheet sheet, int rowIdx, int colIdx, String newValue) {
        Row row = sheet.getRow(rowIdx);
        Cell cell = row.getCell(colIdx);
        cell.setCellValue(newValue);
    }

    // Kommentar Maria M: updateRangeofCells in UpdateKlasse hier stattdessen idealerweise aufrufen
    //Werte aktualisieren
    public void update(ActionEvent actionEvent) throws IOException {
        String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
        String sheetName = "Basisinformation";

        try (FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                sheet = workbook.createSheet(sheetName);
                System.out.println("Sheet == NULL");
            }

            //leere daten
            if (firmenname.getText().isEmpty() && strasse.getText().isEmpty() && plz.getText().isEmpty() && ort.getText().isEmpty() &&
                    schule.getText().isEmpty() && oeffi.getText().isEmpty()) {
                resultLabel.setText("Daten erforderlich zum Aktualisieren");
                return;
            }
            //Reihenfolge: firmenname, strasse, plz, ort, schule, oeffi
            if (!firmenname.getText().isEmpty() && !IAllExcelRegisterCards.isNumericStr(firmenname.getText())) {
                updateCellValue(sheet, 1, 8, firmenname.getText());
            } else if (!firmenname.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(firmenname.getText())) {
                resultLabel.setText("Gültige Daten erforderlich");
                System.out.println("Firma ungültig" + firmenname.getText());
                return;
            }
            if (!strasse.getText().isEmpty() && !IAllExcelRegisterCards.isNumericStr(strasse.getText())) {
                updateCellValue(sheet, 2, 8, strasse.getText());
            } else if (!strasse.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(strasse.getText())) {
                resultLabel.setText("Gültige Daten erforderlich");
                System.out.println("Strasse");
                return;
            }
            if (!plz.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(plz.getText())) {
                updateCellValue(sheet, 3, 8, plz.getText());
            } else if (!plz.getText().isEmpty() && !IAllExcelRegisterCards.isNumericStr(plz.getText())) {
                resultLabel.setText("Gültige Daten erforderlich");
                System.out.println("plz");
                return;
            }
            if (!ort.getText().isEmpty() && !IAllExcelRegisterCards.isNumericStr(ort.getText())) {
                updateCellValue(sheet, 4, 8, ort.getText());
            } else if (!ort.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(ort.getText())) {
                resultLabel.setText("Gültige Daten erforderlich");
                System.out.println("ort");
                return;
            }
            if (!schule.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(firmenname.getText())) {
                updateCellValue(sheet, 5, 8, schule.getText());
            } else if (!schule.getText().isEmpty() && !IAllExcelRegisterCards.isNumericStr(schule.getText())) {
                resultLabel.setText("Gültige Daten erforderlich");
                return;
            }
            if (!oeffi.getText().isEmpty() && !IAllExcelRegisterCards.isNumericStr(firmenname.getText())) {
                updateCellValue(sheet, 8, 8, oeffi.getText());
            } else if (!oeffi.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(oeffi.getText())) {
                resultLabel.setText("Gültige Daten erforderlich");
                return;
            }

            resultLabel.setText("Daten erfolgreich geändert. ");
            // Automatische Auswertung der Formeln im gesamten Arbeitsblatt
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            fileInputStream.close();

            FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace(); // Handle or log the exception as needed
        }
    }

    //Maria M
    public void weiter(ActionEvent actionEvent) {

        BasisinformationApplication basisinformationApplication = new BasisinformationApplication();
        Weiter.weiter(weiterButton, BasisinformationApplication.class);
    }


    //-n Wenn ein Bild hochgeladen wird, wird eine pdf nur mit bildern erstellt. die pdfs werden beim weiter
    //klick zusammengefügt.
    public void uploadImage(ActionEvent actionEvent) {
        // Erstelle eine Instanz von FileChooser
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Bild auswählen");

        // Füge eine Filteroption für Bilddateien hinzu (optional)
        FileChooser.ExtensionFilter imageFilter = new FileChooser.ExtensionFilter("Bilder", "*.png", "*.jpg", "*.gif");
        fileChooser.getExtensionFilters().add(imageFilter);

        // Zeige den FileChooser und erhalte das ausgewählte Bild
        File selectedFile = fileChooser.showOpenDialog(null);

        if (selectedFile != null) {
            // Lade das ausgewählte Bild
            Image image = new Image(selectedFile.toURI().toString());

            // Erstelle ein ImageView und setze das Bild
            ImageView imageView = new ImageView(image);

            // Setze das ImageView in einen StackPane (oder einen anderen Container deiner Wahl)
            StackPane stackPane = new StackPane();
            stackPane.getChildren().add(imageView);

            // Erstelle die Szene und füge das StackPane hinzu
            Scene scene = new Scene(stackPane, 600, 400);

            // Setze die Szene für die Bühne (Stage)
            Stage stage = new Stage();
            stage.setScene(scene);

            // Setze den Titel und zeige die Bühne
            stage.setTitle("Hochgeladenes Bild");
            stage.show();

            // Hier kannst du die generatepdf-Methode aufrufen und den Bildpfad übergeben
            generatePdf(selectedFile.getAbsolutePath());
        }
    }

    //-nn pdf generieren nur für bilder
    public static void generatePdf(String imagePath) {
        try {
            String existingPdfPath = "src/main/resources/com/example/financingtool/Stammblattimg.pdf";
            String outputPdfPath = "src/main/resources/com/example/financingtool/Stammblattimg.pdf";
            String logoPath = "src/main/resources/com/example/financingtool/logo.jpg";

            PDDocument document;

            // Erstelle ein neues Dokument, wenn es nicht existiert
            if (Files.exists(Paths.get(existingPdfPath))) {
                // Lade die vorhandene PDF
                document = PDDocument.load(new File(existingPdfPath));
            } else {
                document = new PDDocument();
            }

            try {
                // Füge eine neue Seite hinzu
                PDPage page = new PDPage();
                document.addPage(page);

                // Lade das Bild
                PDImageXObject image = PDImageXObject.createFromFile(imagePath, document);
                PDImageXObject logo = PDImageXObject.createFromFile(logoPath, document);
                // Überprüfe, ob das Bild erfolgreich geladen wurde
                if (image != null) {
                    // Füge das Bild auf der Seite hinzu
                    try (PDPageContentStream contentStream = new PDPageContentStream(document, page, PDPageContentStream.AppendMode.APPEND, true)) {
                        // image.getHeight(), image.getWidth()
                        contentStream.drawImage(image, 50, 100, 500, 500);
                        contentStream.drawImage(logo, 430, 630, 100, 100);
                    }

                    // Speichere das aktualisierte PDF-Dokument
                    document.save(outputPdfPath);
                    System.out.println("Bild erfolgreich zu vorhandener/neuer PDF hinzugefügt.");
                } else {
                    System.out.println("Fehler beim Laden des Bildes.");
                }

            } finally {
                // Schließe das Dokument
                document.close();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }




    //excel to word
    public static void ExceltoWord() {

        try {
            FileInputStream excelFile = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheet("Stammdaten");
            XWPFDocument document = new XWPFDocument();

            // Create a paragraph and run to add content
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            System.out.println(15840 / 2);
            double x = 15840 / 27.94;
            System.out.println("x= " + x);
            System.out.println(x * 29.7);
            // Set Word document in landscape orientation
            document.getDocument().getBody().addNewSectPr().addNewPgSz().setW(x * 29.7);
            document.getDocument().getBody().addNewSectPr().addNewPgSz().setH(x * 21);

            FileOutputStream out = new FileOutputStream(wordFilePath);

            XWPFTable table = document.createTable();

            String[][] headers = new String[7][1];
            headers[0][0] = "Firmenname";
            headers[1][0] = "Strasse";
            headers[2][0] = "Postleitzahl";
            headers[3][0] = "Ort";
            headers[4][0] = "Entfernung zur Schule";
            headers[5][0] = "Lagebeschreibung";
            headers[6][0] = "Öffentliche Verkehrsmittel";



            XWPFTableRow headerRow = table.getRow(0);
           /* for (int i = 0; i < headers.length; i++) {
                XWPFTableCell cell = headerRow.getCell(i);
                if (cell == null) {
                    cell = headerRow.createCell();
                }
                cell.setText(headers[i][0]);
            }

            */

            for (int rowIdx = 1; rowIdx < 7; rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                XWPFTableRow dataRow = table.createRow();

                for (int colIdx = 0; colIdx < 2; colIdx++) {
                    XWPFTableCell cell = dataRow.getCell(colIdx);
                   if (cell == null) { //errorhandling
                        cell = dataRow.createCell();
                    }


                    if (row.getCell(colIdx) != null) {

                        cell.setText(row.getCell(colIdx).toString());
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
