package com.example.financingtool;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class HelloApplication extends Application {

    private Label resultLabel;

    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(HelloApplication.class.getResource("hello-view.fxml"));
        VBox root = fxmlLoader.load();
        Scene scene = new Scene(root, 1280, 720);
        stage.setTitle("Hello!");

        resultLabel = new Label("Aktueller Wert: ");

        // Erstellen Sie 10 Textfelder für die Benutzereingabe
        TextField[] userInputFields = new TextField[10];
        for (int i = 0; i < 10; i++) {
            userInputFields[i] = new TextField();
            if (i<3) {
                userInputFields[i].setPromptText("KB0" + i + " netto eingeben");
            }else{
                userInputFields[i].setPromptText("KB0" + (i+3) + " netto eingeben");
            }

        }

        // Button hinzufügen, um den Bereich von B3 bis B10 zu aktualisieren
        javafx.scene.control.Button updateRangeButton = new javafx.scene.control.Button("Bereich aktualisieren");
        updateRangeButton.setOnAction(e -> {
            String[] newValues = new String[10];
            for (int i = 0; i < 9; i++) {
                newValues[i] = userInputFields[i].getText();
            }
            updateRangeOfCells(newValues);
        });

        // Füge alle UI-Elemente zum Root-VBox hinzu
        for (int i = 0; i < 9; i++) {
            root.getChildren().add(userInputFields[i]);
        }
        root.getChildren().addAll(resultLabel, updateRangeButton);

        // Setze die Szene und zeige die Bühne
        stage.setScene(scene);
        stage.show();
    }

    public static void main(String[] args) {
        launch(args);
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


    private void updateCellValue(Sheet sheet, int rowIdx, int colIdx, double newValue) {
        Row row = sheet.getRow(rowIdx);
        Cell cell = row.getCell(colIdx);
        cell.setCellValue(newValue);
    }
}
