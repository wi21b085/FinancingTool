package com.example.financingtool;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class HelloApplication extends Application {

    private Label resultLabel;

    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(HelloApplication.class.getResource("hello-view.fxml"));
        ScrollPane scrollPane = fxmlLoader.load();
        VBox root = (VBox) scrollPane.getContent();

        Scene scene = new Scene(scrollPane, 1280, 720);
        stage.setTitle("Hello!");

        resultLabel = new Label("Aktueller Wert: ");

        // Erstellen Sie 10 Textfelder für die Benutzereingabe
        TextFieldWithValidation[] userInputFields = new TextFieldWithValidation[10];
        for (int i = 0; i < 10; i++) {
            userInputFields[i] = new TextFieldWithValidation();
            if (i < 3) {
                userInputFields[i].setPromptText("KB0" + i + " netto eingeben");
            } else {
                userInputFields[i].setPromptText("KB0" + (i + 3) + " netto eingeben");
            }

        }
        // Erstellen Sie Textfelder für die Benutzereingabe für Spalte D
        TextFieldWithValidation userInputFieldD2 = new TextFieldWithValidation();
        userInputFieldD2.setPromptText("UST Grund");

        TextFieldWithValidation userInputFieldD3to9 = new TextFieldWithValidation();
        userInputFieldD3to9.setPromptText("Genereller UST");

        TextFieldWithValidation userInputFieldD10 = new TextFieldWithValidation();
        userInputFieldD10.setPromptText("UST Finanzierung");

        // Button hinzufügen, um den Bereich von B3 bis B10 zu aktualisieren
        javafx.scene.control.Button updateRangeButton = new javafx.scene.control.Button("Bereich aktualisieren");
        updateRangeButton.setOnAction(e -> {

            String[] newValues = new String[10];
            int countNonNumeric=0;
            String nonNumericValue=new String();
            for (int i = 0; i < 9; i++) {
                if (userInputFields[i].isNumeric() || userInputFields[i].getText().trim().isEmpty()) {
                    newValues[i] = userInputFields[i].getText();
                    updateRangeOfCells(newValues);
                }else{
                    countNonNumeric++;
                    nonNumericValue=userInputFields[i].getText();
                    System.out.println(nonNumericValue);
                }

            }
            if(countNonNumeric>0){
                System.out.println("Achtung: Die Werte müssen numerisch sein. Fehler bei  "+ nonNumericValue);
                resultLabel.setText("Achtung: Die Werte müssen numerisch sein. Fehler bei"+ nonNumericValue);
            }

        });
        javafx.scene.control.Button updateButtonD = new javafx.scene.control.Button("UST aktualisieren");
        updateButtonD.setOnAction(e -> {

            String newValueD2 = parsePercentageValue(userInputFieldD2.getText());
            System.out.println(newValueD2);
            String newValueD3to9 = parsePercentageValue(userInputFieldD3to9.getText());
            System.out.println(newValueD3to9);
            String newValueD10 = parsePercentageValue(userInputFieldD10.getText());
            System.out.println(newValueD10);

            boolean checkstrNewValueD2=isNumericStr(newValueD2);
            boolean checkstrNewValueD3to9=isNumericStr(newValueD3to9);
            boolean checkstrNewValueD10=isNumericStr(newValueD10);


            if(checkstrNewValueD2|| userInputFieldD2.getText().trim().isEmpty()) {
                if(testPercentageRange(newValueD2)){
                    updateCellD(20, 1, newValueD2);
                    updateCellD(1, 3, newValueD2);
                }else{
                    System.out.println("Achtung die Werte müssen zwischen 0%-100% bzw zwischen 0.0-1.0 betragen.");
                    resultLabel.setText("Achtung die Werte müssen zwischen 0%-100% bzw zwischen 0.0-1.0 betragen.");
                }



            }else{
                System.out.println("Achtung die Werte müssen numerisch sein");
                resultLabel.setText("Achtung die Werte müssen numerisch sein");

            }if(checkstrNewValueD3to9|| userInputFieldD3to9.getText().trim().isEmpty()) {
                if(testPercentageRange(newValueD3to9)) {
                    updateCellD(19, 1, newValueD3to9);
                    for (int i = 2; i < 9; i++) {
                        updateCellD(i, 3, newValueD3to9);
                    }
                }else{
                    System.out.println("Achtung die Werte müssen zwischen 0%-100% bzw zwischen 0.0-1.0 betragen.");
                    resultLabel.setText("Achtung die Werte müssen zwischen 0%-100% bzw zwischen 0.0-1.0 betragen.");
                }
            }else{
                    System.out.println("Achtung die Werte müssen numerisch sein");
                    resultLabel.setText("Achtung die Werte müssen numerisch sein");

                }

            if(checkstrNewValueD10|| userInputFieldD10.getText().trim().isEmpty()) {
                if(testPercentageRange(newValueD10)) {
                    updateCellD(9,3, newValueD10);

                }
            }else{
                System.out.println("Achtung die Werte müssen numerisch sein");
                resultLabel.setText("Achtung die Werte müssen numerisch sein");

            }



        });

        // Füge alle UI-Elemente zum Root-VBox hinzu
        for (int i = 0; i < 10; i++) {
            root.getChildren().add(userInputFields[i]);
        }
        root.getChildren().addAll(resultLabel, updateRangeButton, userInputFieldD2, userInputFieldD3to9, userInputFieldD10, updateButtonD);

        // Setze die Szene und zeige die Bühne
        stage.setScene(scene);
        stage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }

    public static boolean isNumericStr(String str) {
        try {
            double numericValue = Double.parseDouble(str);
            // Wenn die Konvertierung erfolgreich ist, ist der String numerisch
            return true;
        } catch (NumberFormatException e) {
            // Wenn eine NumberFormatException auftritt, ist der String nicht numerisch
            return false;
        }
    }



    private void updateRangeOfCells(String[] newValues) {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Gesamtinvestitionskosten";
            int startRow = 1;
            int endRow = 9;
            int colIdx = 1;

            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheet(sheetName);

            // Annahme: newValues enthält die neuen Werte für B3 bis B10
            for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
                String newValueStr = newValues[rowIdx - startRow];

                // Überprüfen Sie, ob die Zeichenkette nicht leer ist und nicht null ist, bevor Sie sie parsen
                if (newValueStr != null && !newValueStr.isEmpty()) {
                    double newValue = Double.parseDouble(newValueStr);
                    updateCellValue(sheet, rowIdx, colIdx, newValue);
                }
            }


            // Automatische Auswertung der Formeln im gesamten Arbeitsblatt
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            fileInputStream.close();

            FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();

            resultLabel.setText("Bereich von B2 bis B10 erfolgreich aktualisiert.");
            System.out.println("Zellen erfolgreich aktualisiert.");

        } catch (NumberFormatException | IOException e) {
            e.printStackTrace();
            resultLabel.setText("Fehler bei der Aktualisierung.");
        }
    }

    private void updateCellD(int rowIdx, int colIdx, String newValue) {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Gesamtinvestitionskosten";

            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheet(sheetName);

            // Überprüfen Sie, ob die Zeichenkette nicht leer ist und nicht null ist, bevor Sie sie parsen
            if (newValue != null && !newValue.isEmpty()) {
                double newCellValue = Double.parseDouble(newValue);
                updateCellValue(sheet, rowIdx, colIdx, newCellValue);
            }

            // Automatische Auswertung der Formeln im gesamten Arbeitsblatt
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            fileInputStream.close();

            FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();

            resultLabel.setText("Zelle erfolgreich aktualisiert.");
            System.out.println("Zelle erfolgreich aktualisiert.");

        } catch (NumberFormatException | IOException e) {
            e.printStackTrace();
            resultLabel.setText("Fehler bei der Aktualisierung von Zelle D" + rowIdx);
        }
    }



    private void updateCellValue(Sheet sheet, int rowIdx, int colIdx, double newValue) {
        Row row = sheet.getRow(rowIdx);
        Cell cell = row.getCell(colIdx);
        cell.setCellValue(newValue);
    }

    private class TextFieldWithValidation extends TextField {
        public boolean isNumeric() {
            try {
                Double.parseDouble(getText());
                return true;
            } catch (NumberFormatException e) {
                return false;
            }
        }
    }
    private String parsePercentageValue(String value) {
        value = value.trim(); // Entferne führende und abschließende Leerzeichen

        if (value.endsWith("%")) {
            try {
                // Extrahiere den Prozentanteil und konvertiere ihn in einen Dezimalwert
                double percentage = Double.parseDouble(value.substring(0, value.length() - 1));
                // Teile durch 100, um den Wert in das Dezimalformat zu konvertieren
                return String.valueOf(percentage / 100.0);
            } catch (NumberFormatException e) {
                // Fehler beim Parsen der Zahl
                e.printStackTrace();
                return "0.0"; // Standardwert oder Fehlerbehandlung nach Bedarf
            }
        }
        return value;
    }


    private boolean testPercentageRange(String value){
        if (!value.isEmpty()) {
        if(Double.parseDouble(value)<=1&&Double.parseDouble(value)>=0) {
            return true;
        }else{
            return false;
        }
        }
            return true
                    ;

    }


}
