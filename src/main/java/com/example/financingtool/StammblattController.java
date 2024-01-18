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
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class StammblattController implements IAllExcelRegisterCards {

    static String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
    static String wordFilePath = "src/main/resources/com/example/financingtool/SEPJ_Stammblatt.docx";
    static String pdfFilePath = "src/main/resources/com/example/financingtool/SEPJ_Stammblatt.pdf";

    //wird das mit dem pdf ein problem werden, wenn es noch nicht existiert?

    //Maria M
    @FXML
    private Button weiterButton;
    @FXML
    private Label resultLabelBasisinformation;
    @FXML
    private Label resultLabelStammdaten;
    //stammdaten userinput
    @FXML
    private TextField firmenname;
    @FXML
    private TextField strasse;
    @FXML
    private TextField plz;
    @FXML
    private TextField ort;
    @FXML
    private TextField oeffi;
    @FXML
    private TextField lage;

    //basisinformation userinput
    @FXML
    private TextField kaufpreis;
    @FXML
    private TextField groesse;

    @FXML
    private TextField nutzflaeche;
    @FXML
    private TextField wohneinheiten;
    @FXML
    private TextField garage;
    @FXML
    private TextField aussenflaeche;
    @FXML
    private TextField verkaufserloes;
    @FXML
    private TextField gewinn;
    @FXML
    private TextField beginn;
    @FXML
    private TextField ende;
    @FXML
    private TextField roi;

    //pfade
    private static final String FILE_NAME = "SEPJ-Rechnungen.xlsx";
    private static final String SHEET_NAME = "Basisinformationen";

    //basisinformation
    static ExecutiveSummary executiveSummary = new ExecutiveSummary();

    static WidmungController widmungController=new WidmungController();

    public static void setWidmungController(WidmungController widmungController){
        StammblattController.widmungController=widmungController;
    }

    public void onHelloButtonClick(ActionEvent actionEvent) {
    }

    //firmennamen in das excel eintragen
    public void submit() {
        //Submit nur möglich, wenn alle Felder befüllt.

        submitBasisInformation();

        //leere Eingabe
        if (firmenname.getText().isEmpty() || strasse.getText().isEmpty() || plz.getText().isEmpty()
                || ort.getText().isEmpty() || lage.getText().isEmpty()  || oeffi.getText().isEmpty()) {
            resultLabelStammdaten.setText("Stammdaten unvollständig");
            return;
        }
        //ungültige eingabe
        else if (IAllExcelRegisterCards.isNumericStr(firmenname.getText()) ||
                IAllExcelRegisterCards.isNumericStr(strasse.getText()) ||
                IAllExcelRegisterCards.isNumericStr(ort.getText()) ||
                IAllExcelRegisterCards.isNumericStr(lage.getText()) ||
                IAllExcelRegisterCards.isNumericStr(oeffi.getText()) ||
                !IAllExcelRegisterCards.isNumericStr(plz.getText())
        ) {
            resultLabelStammdaten.setText("Achtung, bitte geben Sie gültige Daten an");
            return;
        }
        //richtige Werte
        else {
            //    boolean str = IAllExcelRegisterCards.isNumericStr(firmenname.getText());

            String[] newValue = new String[6];
            newValue[0] = firmenname.getText();
            newValue[1] = strasse.getText();
            newValue[2] = plz.getText();
            newValue[3] = ort.getText();
            newValue[4] = lage.getText();
            newValue[5] = oeffi.getText();
            String strasseValue=strasse.getText();
            setWidmungController(widmungController);
            EventBus.getInstance().publish("updateAddress",strasseValue);
            writeToExcel(newValue);
            resultLabelStammdaten.setText("Stammdaten erfolgreich eingefügt");
        }

    }

    private void submitBasisInformation() {

        //leere werte
        if(kaufpreis.getText().isEmpty() || groesse.getText().isEmpty() || nutzflaeche.getText().isEmpty()
                || wohneinheiten.getText().isEmpty() || garage.getText().isEmpty() || aussenflaeche.getText().isEmpty()
                || verkaufserloes.getText().isEmpty() || gewinn.getText().isEmpty() || beginn.getText().isEmpty()
                ||ende.getText().isEmpty() || roi.getText().isEmpty() ){
            resultLabelBasisinformation.setText("Basisinformation unvollständig");
            // System.out.println("Daten unvollständig");
        }
        //ungültige werte
        else if (!IAllExcelRegisterCards.isNumericStr(kaufpreis.getText()) ||
                !IAllExcelRegisterCards.isNumericStr(groesse.getText()) ||
                !IAllExcelRegisterCards.isNumericStr(nutzflaeche.getText()) ||
                !IAllExcelRegisterCards.isNumericStr(wohneinheiten.getText()) ||
                !IAllExcelRegisterCards.isNumericStr(garage.getText()) ||
                !IAllExcelRegisterCards.isNumericStr(aussenflaeche.getText()) ||
                !IAllExcelRegisterCards.isNumericStr(verkaufserloes.getText()) ||
                !IAllExcelRegisterCards.isNumericStr(verkaufserloes.getText())
        ){
            resultLabelBasisinformation.setText("Achtung, bitte geben Sie gültige Daten an");
            return;
        }
        //gültige werte
        else {

            String[] newValue = new String[11];
            newValue[0] = kaufpreis.getText();
            newValue[1] = groesse.getText();
            newValue[2] = nutzflaeche.getText();
            newValue[3] = wohneinheiten.getText();
            newValue[4] = garage.getText();
            newValue[5] = aussenflaeche.getText();
            newValue[6] = verkaufserloes.getText();
            newValue[7] = gewinn.getText();
            newValue[8] = beginn.getText();
            newValue[9] = ende.getText();
            newValue[10] = roi.getText();
            setExecutiveSummary(executiveSummary);
            System.out.println("Daten aus Basisinformation gesendet gesendet: ");
            executiveSummary.setDatenausBas(kaufpreis.getText(),groesse.getText(),wohneinheiten.getText(),garage.getText(),beginn.getText(),ende.getText());

            writeToBasisInformationExcel(newValue);
        }
    }

    public static void setExecutiveSummary(ExecutiveSummary executiveSummary) {
        StammblattController.executiveSummary=executiveSummary;
    }

    private void writeToBasisInformationExcel(String[] newValue) {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Basisinformation";

            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheet(sheetName);

            // Überprüfen Sie, ob die Zeichenkette nicht leer ist und nicht null ist, bevor Sie sie parsen
            if (newValue != null && !newValue[0].isEmpty()) {
                updateCellValue(sheet, 1, 1, newValue[0]); //Reihenfolge:
                updateCellValue(sheet,2, 1, newValue[1]);
                updateCellValue(sheet, 3, 1, newValue[2]);
                updateCellValue(sheet, 4,1, newValue[3]);
                updateCellValue(sheet, 5,1, newValue[4]);
                updateCellValue(sheet, 6,1, newValue[5]);
                updateCellValue(sheet, 7,1, newValue[6]);
                updateCellValue(sheet, 8,1, newValue[7]);
                updateCellValue(sheet, 9,1, newValue[8]);
                updateCellValue(sheet, 10,1, newValue[9]);
                updateCellValue(sheet, 11,1, newValue[10]);
                //kommt in die Zelle 7, 7
            }
            else{
                //eig sollte das eh nicht vorkommen, weil es davor schon ausgeschlossen ist.
                resultLabelBasisinformation.setText("Basisinformationen unvollständig");
            }

            // Automatische Auswertung der Formeln im gesamten Arbeitsblatt
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            fileInputStream.close();

            FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();

            resultLabelBasisinformation.setText("Basisinformation erfolgreich eingefügt");


        } catch (NumberFormatException | IOException e) {
            e.printStackTrace();
            //  resultLabel.setText("Fehler bei der Aktualisierung von Zelle D" + 6);
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
                updateCellValue(sheet, 1, 8, newValue[0]); //Reihenfolge: firmenname, strasse, plz, ort, oeffi
                updateCellValue(sheet, 2, 8, newValue[1]);
                updateCellValue(sheet, 3, 8, newValue[2]);
                updateCellValue(sheet, 4, 8, newValue[3]);
                updateCellValue(sheet, 5, 8, newValue[4]);
                updateCellValue(sheet, 7, 8, newValue[5]);
                //kommt in die Zelle 7, 7
            } else {

                System.out.println("Stammdaten unvollständig");
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
        updateBasisInformation();
        String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
        String sheetName = "Basisinformation";

        try (FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                sheet = workbook.createSheet(sheetName);
                System.out.println("Sheet == NULL");
            }



            //stammdaten anfang
            if (firmenname.getText().isEmpty() && strasse.getText().isEmpty() && plz.getText().isEmpty() && ort.getText().isEmpty() &&
                    oeffi.getText().isEmpty() && lage.getText().isEmpty()) {
                resultLabelStammdaten.setText("Stammdaten leer");
                return;
            }
            //Reihenfolge: firmenname, strasse, plz, ort, oeffi
            if (!firmenname.getText().isEmpty() && !IAllExcelRegisterCards.isNumericStr(firmenname.getText())) {
                updateCellValue(sheet, 1, 8, firmenname.getText());
            } else if (!firmenname.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(firmenname.getText())) {
                resultLabelStammdaten.setText("Gültige Stammdaten erforderlich");
                System.out.println("Firma ungültig" + firmenname.getText());
                return;
            }
            if (!strasse.getText().isEmpty() && !IAllExcelRegisterCards.isNumericStr(strasse.getText())) {
                updateCellValue(sheet, 2, 8, strasse.getText());
            } else if (!strasse.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(strasse.getText())) {
                resultLabelStammdaten.setText("Gültige Daten erforderlich");
                System.out.println("Strasse");
                return;
            }
            if (!plz.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(plz.getText())) {
                updateCellValue(sheet, 3, 8, plz.getText());
            } else if (!plz.getText().isEmpty() && !IAllExcelRegisterCards.isNumericStr(plz.getText())) {
                resultLabelStammdaten.setText("Gültige Daten erforderlich");
                System.out.println("plz");
                return;
            }
            if (!ort.getText().isEmpty() && !IAllExcelRegisterCards.isNumericStr(ort.getText())) {
                updateCellValue(sheet, 4, 8, ort.getText());
            } else if (!ort.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(ort.getText())) {
                resultLabelStammdaten.setText("Gültige Daten erforderlich");
                System.out.println("ort");
                return;
            }

            if (!lage.getText().isEmpty() && !IAllExcelRegisterCards.isNumericStr(lage.getText())) {
                updateCellValue(sheet, 5, 8, lage.getText());
            } else if (!lage.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(lage.getText())) {
                resultLabelStammdaten.setText("Gültige Daten erforderlich");
                System.out.println("lage");
                return;
            }

            //schule row: 5, coll: 8
            if (!oeffi.getText().isEmpty() && !IAllExcelRegisterCards.isNumericStr(firmenname.getText())) {
                updateCellValue(sheet, 7, 8, oeffi.getText());
            } else if (!oeffi.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(oeffi.getText())) {
                resultLabelStammdaten.setText("Gültige Daten erforderlich");
                return;
            }
            //stammdaten ende


            resultLabelStammdaten.setText("Stammdaten erfolgreich geändert. ");
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

    private void updateBasisInformation() {
        String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
        String sheetName = "Basisinformation";

        try (FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                sheet = workbook.createSheet(sheetName);
                System.out.println("Sheet == NULL");
            }

            //BASISINFORMATION ANFANG
            //überprüfen, ob basisinformation daten leer sind
            if (kaufpreis.getText().isEmpty() && groesse.getText().isEmpty() && nutzflaeche.getText().isEmpty() &&
                    wohneinheiten.getText().isEmpty() && garage.getText().isEmpty() && aussenflaeche.getText().isEmpty()
                    && verkaufserloes.getText().isEmpty() && gewinn.getText().isEmpty() && beginn.getText().isEmpty()
                    && ende.getText().isEmpty() && roi.getText().isEmpty()) {
                resultLabelBasisinformation.setText("Basisinformation leer");

                return;
            }

            if (!kaufpreis.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(kaufpreis.getText())) {
                updateCellValue(sheet, 1, 1, kaufpreis.getText());


            } else if (!IAllExcelRegisterCards.isNumericStr(kaufpreis.getText())) {
                resultLabelBasisinformation.setText("Gültige Basisinformationen erforderlich");
                return;
            }

            if (!groesse.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(groesse.getText())) {
                updateCellValue(sheet, 2, 1, groesse.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(groesse.getText())) {
                resultLabelBasisinformation.setText("Gültige Basisinformationen erforderlich");
                return;
            }

            if (!nutzflaeche.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(nutzflaeche.getText())) {
                updateCellValue(sheet, 3, 1, nutzflaeche.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(nutzflaeche.getText())) {
                resultLabelBasisinformation.setText("Gültige Basisinformationen erforderlich");
                return;
            }

            if (!wohneinheiten.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(wohneinheiten.getText())) {
                updateCellValue(sheet, 4, 1, wohneinheiten.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(wohneinheiten.getText())) {
                resultLabelBasisinformation.setText("Gültige Basisinformationen erforderlich");
                return;
            }

            if (!garage.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(garage.getText())) {
                updateCellValue(sheet, 5, 1, garage.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(garage.getText())) {
                resultLabelBasisinformation.setText("Gültige Basisinformationen erforderlich");
                return;
            }

            if (!aussenflaeche.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(aussenflaeche.getText())) {
                updateCellValue(sheet, 6, 1, aussenflaeche.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(aussenflaeche.getText())) {
                resultLabelBasisinformation.setText("Gültige Basisinformationen erforderlich");
                return;
            }

            if (!verkaufserloes.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(verkaufserloes.getText())) {
                updateCellValue(sheet, 7, 1, verkaufserloes.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(verkaufserloes.getText())) {
                resultLabelBasisinformation.setText("Gültige Basisinformationen erforderlich");
                return;
            }

            if (!gewinn.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(gewinn.getText())) {
                updateCellValue(sheet, 8, 1, gewinn.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(gewinn.getText())) {
                resultLabelBasisinformation.setText("Gültige Basisinformationen erforderlich");
                return;
            }


            if (!beginn.getText().isEmpty()) { // Datum kann String + Double sein
                updateCellValue(sheet, 9, 1, beginn.getText());
            }

            if (!ende.getText().isEmpty()) { // Datum kann String + Double sein
                updateCellValue(sheet, 10, 1, ende.getText());
            }

            if (!roi.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(roi.getText())) {
                updateCellValue(sheet, 11, 1, roi.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(roi.getText())) {
                resultLabelBasisinformation.setText("Gültige Basisinformationen erforderlich");
                return;
            }

            // Automatische Auswertung der Formeln im gesamten Arbeitsblatt
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            fileInputStream.close();

            FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();
            resultLabelBasisinformation.setText("Basisinformation erfolgreich aktualisiert");


        } catch (IOException e) {
            e.printStackTrace(); // Handle or log the exception as needed
        }

    }

    //Maria M
    /*public void weiter(ActionEvent actionEvent) {

        BasisinformationApplication basisinformationApplication = new BasisinformationApplication();
        Weiter.weiter(weiterButton, BasisinformationApplication.class);
    }*/


    //-n Wenn ein Bild hochgeladen wird, wird eine pdf nur mit bildern erstellt. die pdfs werden beim weiter
    //klick zusammengefügt.
    public void uploadImage(ActionEvent actionEvent) {
        // Erstelle eine Instanz von FileChooser
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Bild auswählen");

        // Füge eine Filteroption für Bilddateien hinzu (optional)
        FileChooser.ExtensionFilter imageFilter = new FileChooser.ExtensionFilter("Bilder", ".png", ".jpg", "*.gif");
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
                // Füge eine neue Seite hinzu im Querformat hinzu
                PDPage page = new PDPage(new PDRectangle(PDRectangle.A4.getHeight(), PDRectangle.A4.getWidth()));
                // PDPage page = new PDPage();
                document.addPage(page);

                // Lade das Bild
                PDImageXObject image = PDImageXObject.createFromFile(imagePath, document);
                PDImageXObject logo = PDImageXObject.createFromFile(logoPath, document);
                // Überprüfe, ob das Bild erfolgreich geladen wurde
                if (image != null) {
                    // Füge das Bild auf der Seite hinzu
                    try (PDPageContentStream contentStream = new PDPageContentStream(document, page, PDPageContentStream.AppendMode.APPEND, true)) {
                        // image.getHeight(), image.getWidth()
                        contentStream.drawImage(image, 50, 50, 500, 500);
                        contentStream.drawImage(logo, 680, 450, 100, 100);
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

            String[][] headers = new String[6][1];
            headers[0][0] = "Firmenname";
            headers[1][0] = "Strasse";
            headers[2][0] = "Postleitzahl";
            headers[3][0] = "Ort";
            headers[4][0] = "Lagebeschreibung";
            headers[5][0] = "Öffentliche Verkehrsmittel";


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

    public void delete() {
        //Bilder löschen
        String filePath = "src/main/resources/com/example/financingtool/Stammblattimg.pdf";
        try {
            // Pfad zum PDF-Datei erstellen
            Path path = Paths.get(filePath);

            // Datei löschen
            Files.deleteIfExists(path);

            // System.out.println("Die Datei wurde erfolgreich gelöscht: " + filePath);
            resultLabelStammdaten.setText("Bilder erfolgreich gelöscht");
        } catch (IOException e) {
            // Fehler behandeln, wenn das Löschen fehlschlägt
            //  System.err.println("Fehler beim Löschen der Datei: " + e.getMessage());
            resultLabelStammdaten.setText("Fehler beim Löschen der Datei");
        }
    }
}