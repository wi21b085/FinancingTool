package com.example.financingtool;

//import com.itextpdf.kernel.color.Lab;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.Pane;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class WireController implements IAllExcelRegisterCards {

    @FXML
    private Pane pane;
    @FXML
    private TextField ev;
    @FXML
    private TextField av;
    @FXML
    private TextField wepr;
    @FXML
    private TextField wapr;
    @FXML
    private TextField eplz;
    @FXML
    private TextField aplz;
    @FXML
    private TextField pepr;
    @FXML
    private TextField papr;
    @FXML
    private Label resultLabel;

    @FXML
    protected void continueClick() {
        try {
            String evText = ev.getText();
            String eplzText = eplz.getText();
            String avText = av.getText();
            String aplzText = aplz.getText();
            String weprText = wepr.getText();
            String waprText = wapr.getText();
            String peprText = pepr.getText();
            String paprText = papr.getText();

            if (!isValidInput(evText, true) || !isValidInput(eplzText, false) || !isValidInput(avText, true)
                    || !isValidInput(aplzText, false) || !isValidInput(weprText, false) || !isValidInput(waprText, false)
                    || !isValidInput(peprText, false) || !isValidInput(paprText, false)) {
                resultLabel.setText("Ungültige Eingabe! Stellen Sie sicher, dass alle Felder vollständig ausgefüllt sind und gültige Daten enthalten!");
            }

            // Werte werden in ein Array gespeichert
            String[] value = new String[8];
            value [0] = evText;
            value [1] = eplzText;
            value [2] = avText;
            value [3] = aplzText;
            value [4] = weprText;
            value [5] = waprText;
            value [6] = peprText;
            value [7] = paprText;

            // Excel-Datei laden
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Wirtschaftlichkeitsrechnung";
            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            fileInputStream.close();

            // Werte ins Excel eintragen
            Sheet sheet = workbook.getSheet(sheetName);
            sheet.getRow(4).getCell(1).setCellValue(Double.parseDouble(value[0])); // ev - B4
            sheet.getRow(5).getCell(1).setCellValue(Double.parseDouble(value[2])); // av - B5
            sheet.getRow(3).getCell(5).setCellValue(Double.parseDouble(value[4])); // wepr - F3
            sheet.getRow(3).getCell(7).setCellValue(Double.parseDouble(value[5])); // wapr - H3
            sheet.getRow(10).getCell(2).setCellValue(Double.parseDouble(value[1])); // eplz - C10
            sheet.getRow(11).getCell(2).setCellValue(Double.parseDouble(value[3])); // aplz - C11
            sheet.getRow(9).getCell(5).setCellValue(Double.parseDouble(value[6])); // pepr - F9
            sheet.getRow(9).getCell(7).setCellValue(Double.parseDouble(value[7])); // papr - H9

            // Formeln aktualisieren
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            // Excel-Datei speichern
            FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();

            resultLabel.setText("Werte wurden erfolgreich aktualisiert");


        } catch (Exception e) {
            e.printStackTrace();
            Alert alert = new Alert(Alert.AlertType.ERROR);
            alert.setTitle("Fehler");
            alert.setHeaderText("Ein Fehler ist aufgetreten");
            alert.setContentText("Fehlermeldung: " + e.getMessage());
            alert.showAndWait();
        }

    }

    private boolean isValidInput(String input, boolean isPercentage) {
        // Überprüfen, ob der Input leer ist oder eine Zahl bzw. ein Prozentsatz ist
        if (input.trim().isEmpty()) {
            return false;
        }

        if (isPercentage) {
            boolean isPercentageInRange = IAllExcelRegisterCards.testPercentageRange(input);
            return isPercentageInRange;
        } else {
            return IAllExcelRegisterCards.isNumericStr(input);
        }
    }



}
