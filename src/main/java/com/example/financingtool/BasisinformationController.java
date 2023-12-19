package com.example.financingtool;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.text.Text;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class BasisinformationController {
    @FXML
    private  Label resultLabel;
    //User inputs
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
    private TextField gik;

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

    private static final String FILE_NAME = "SEPJ-Rechnungen.xlsx";
    private static final String SHEET_NAME = "Basisinformationen";

    public void onHelloButtonClick(ActionEvent actionEvent) {
    }

    //firmennamen in das excel eintragen
    public void submit() {
        //Submit nur möglich, wenn alle Felder befüllt.

        //Parse the values.

        if(kaufpreis.getText().isEmpty() || groesse.getText().isEmpty() || nutzflaeche.getText().isEmpty()
                || wohneinheiten.getText().isEmpty() || garage.getText().isEmpty() || gik.getText().isEmpty()
                || verkaufserloes.getText().isEmpty() || gewinn.getText().isEmpty() || beginn.getText().isEmpty()
                ||ende.getText().isEmpty() || roi.getText().isEmpty() ){
            resultLabel.setText("Daten unvollständig");
            // System.out.println("Daten unvollständig");
        }
        else if (kaufpreis.getText() instanceof String || groesse.getText() instanceof String || nutzflaeche.getText() instanceof String
                || wohneinheiten.getText() instanceof String || garage.getText() instanceof String || gik.getText() instanceof String
                || verkaufserloes.getText() instanceof String || gewinn.getText() instanceof String
                 || roi.getText() instanceof String){
            resultLabel.setText("Ein/Mehrere Werte sind ungültig. Bitte versuchen Sie es erneut.");

        }
        else {

            String[] newValue = new String[11];
            newValue[0] = kaufpreis.getText();
            newValue[1] = groesse.getText();
            newValue[2] = nutzflaeche.getText();
            newValue[3] = wohneinheiten.getText();
            newValue[4] = garage.getText();
            newValue[5] = gik.getText();
            newValue[6] = verkaufserloes.getText();
            newValue[7] = gewinn.getText();
            newValue[8] = beginn.getText();
            newValue[9] = ende.getText();
            newValue[10] = roi.getText();

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
                resultLabel.setText("Daten unvollständig");
            }

            // Automatische Auswertung der Formeln im gesamten Arbeitsblatt
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            fileInputStream.close();

            FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();

            resultLabel.setText("Basisinformation erfolgreich gesendet.");


        } catch (NumberFormatException | IOException e) {
            e.printStackTrace();
            //  resultLabel.setText("Fehler bei der Aktualisierung von Zelle D" + 6);
        }
    }

    private static void updateCellValue(Sheet sheet, int rowIdx, int colIdx, String newValue) {
        Row row = sheet.getRow(rowIdx);
        Cell cell = row.getCell(colIdx);
        cell.setCellValue(newValue);
    }


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
            //Überprüfen, ob die Datentypen stimmen:
            else if (kaufpreis.getText() instanceof String || groesse.getText() instanceof String || nutzflaeche.getText() instanceof String
                    || wohneinheiten.getText() instanceof String || garage.getText() instanceof String || gik.getText() instanceof String
                    || verkaufserloes.getText() instanceof String || gewinn.getText() instanceof String
                    || roi.getText() instanceof String  ){
                resultLabel.setText("Ein/Mehrere Werte sind ungültig. Bitte versuchen Sie es erneut.");
                return;
            }


            if (!kaufpreis.getText().isEmpty()) {
                updateCellValue(sheet, 1, 1, kaufpreis.getText());
            }
            if (!groesse.getText().isEmpty()) {
                updateCellValue(sheet, 2, 1, groesse.getText());
            }
            if (!nutzflaeche.getText().isEmpty()) {
                updateCellValue(sheet, 3, 1, nutzflaeche.getText());
            }
            if (!wohneinheiten.getText().isEmpty()) {
                updateCellValue(sheet, 4, 1, wohneinheiten.getText());
            }
            if (!garage.getText().isEmpty()) {
                updateCellValue(sheet, 5, 1, garage.getText());
            }
            if (!gik.getText().isEmpty()) {
                updateCellValue(sheet, 6, 1, gik.getText());
            }
            if (!verkaufserloes.getText().isEmpty()) {
                updateCellValue(sheet, 7, 1, verkaufserloes.getText());
            }
            if (!gewinn.getText().isEmpty()) {
                updateCellValue(sheet, 8, 1, gewinn.getText());
            }
            if (!beginn.getText().isEmpty()) {
                updateCellValue(sheet, 9, 1, beginn.getText());
            }
            if (!ende.getText().isEmpty()) {
                updateCellValue(sheet, 10, 1, ende.getText());
            }
            if (!roi.getText().isEmpty()) {
                updateCellValue(sheet, 11, 1, ende.getText());
            }

            // Automatische Auswertung der Formeln im gesamten Arbeitsblatt
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            fileInputStream.close();

            FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();
            resultLabel.setText("Basisinformation erfolgreich aktualisiert");


        } catch (IOException e) {
            e.printStackTrace(); // Handle or log the exception as needed
        }


    }




}
