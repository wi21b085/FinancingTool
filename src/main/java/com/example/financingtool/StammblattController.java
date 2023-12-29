package com.example.financingtool;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.text.Text;

import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class StammblattController {


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

        if(firmenname.getText().isEmpty() || strasse.getText().isEmpty() || plz.getText().isEmpty()
        || ort.getText().isEmpty() || lage.getText().isEmpty() || schule.getText().isEmpty() || oeffi.getText().isEmpty()){
            resultLabel.setText("Daten unvollständig");
           // System.out.println("Daten unvollständig");
        }
        else if (!(firmenname.getText() instanceof String) || !(strasse.getText() instanceof String) || plz.getText() instanceof String
                || !(ort.getText() instanceof String) || !(lage.getText() instanceof String) || schule.getText() instanceof String
                || !(oeffi.getText() instanceof String)){
            resultLabel.setText("Ein/Mehrere Werte sind ungültig. Bitte versuchen Sie es erneut.");
            return;
        }
        else {

            String[] newValue = new String[7];
            newValue[0] = firmenname.getText();
            newValue[1] = strasse.getText();
            newValue[2] = plz.getText();
            newValue[3] = ort.getText();
            newValue[4] = schule.getText();
            newValue[5] = lage.getText();
            newValue[6] = oeffi.getText();
            writeToExcel(newValue);
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
                updateCellValue(sheet,2, 8, newValue[1]);
                updateCellValue(sheet, 3, 8, newValue[2]);
                updateCellValue(sheet, 4,8,newValue[3]);
                updateCellValue(sheet, 5,8, newValue[4]);
                updateCellValue(sheet, 6,8, newValue[5]);
                updateCellValue(sheet, 8,8, newValue[6]);
                //kommt in die Zelle 7, 7
            }
            else if (!(firmenname.getText() instanceof String) || !(strasse.getText() instanceof String) || plz.getText() instanceof String
                    || !(ort.getText() instanceof String) || !(lage.getText() instanceof String) || schule.getText() instanceof String
                    || !(oeffi.getText() instanceof String)){
                resultLabel.setText("Ein/Mehrere Werte sind ungültig. Bitte versuchen Sie es erneut.");
                return;
            }
            else{

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
                //Reihenfolge: firmenname, strasse, plz, ort, schule, oeffi
                if (!firmenname.getText().isEmpty()) {
                    updateCellValue(sheet, 1, 8, firmenname.getText());
                }
                if (!strasse.getText().isEmpty()) {
                    updateCellValue(sheet, 2, 8, strasse.getText());
                }
                if (!plz.getText().isEmpty()) {
                    updateCellValue(sheet, 3, 8, plz.getText());
                }
                if (!ort.getText().isEmpty()) {
                    updateCellValue(sheet, 4, 8, ort.getText());
                }
                if (!schule.getText().isEmpty()) {
                    updateCellValue(sheet, 5, 8, schule.getText());
                }
                if (!oeffi.getText().isEmpty()) {
                    updateCellValue(sheet, 8, 8, oeffi.getText());
                }
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
       Weiter.weiter(weiterButton, BasisinformationApplication.class);
    }





}
