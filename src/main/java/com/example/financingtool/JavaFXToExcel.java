package com.example.financingtool;
import javafx.application.Application;
import javafx.beans.property.SimpleStringProperty;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class JavaFXToExcel extends Application {

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        // Erstelle eine einfache JavaFX-Anwendung mit einer TableView, Eingabefeldern und einem Button
        TableView<String[]> tableView = new TableView<>();
        tableView.setEditable(true);

        TableColumn<String[], String> nameCol = new TableColumn<>("Name");
        TableColumn<String[], String> ageCol = new TableColumn<>("Alter");
        TableColumn<String[], String> cityCol = new TableColumn<>("Stadt");

        nameCol.setCellValueFactory(cellData -> new SimpleStringProperty(cellData.getValue()[0]));
        ageCol.setCellValueFactory(cellData -> new SimpleStringProperty(cellData.getValue()[1]));
        cityCol.setCellValueFactory(cellData -> new SimpleStringProperty(cellData.getValue()[2]));

        tableView.getColumns().addAll(nameCol, ageCol, cityCol);

        TextField nameField = new TextField();
        TextField ageField = new TextField();
        TextField cityField = new TextField();
        Button addButton = new Button("Hinzufügen");
        Button exportButton = new Button("Daten exportieren");

        HBox inputBox = new HBox(nameField, ageField, cityField, addButton);
        VBox root = new VBox(tableView, inputBox, exportButton);
        Scene scene = new Scene(root, 600, 400);

        primaryStage.setTitle("Rechnungen_V1");
        primaryStage.setScene(scene);
        primaryStage.show();

        // Event Handler für den Hinzufügen-Button
        addButton.setOnAction(event -> {
            String[] newData = new String[]{nameField.getText(), ageField.getText(), cityField.getText()};
            tableView.getItems().add(newData);
            // Leere die Eingabefelder nach dem Hinzufügen
            nameField.clear();
            ageField.clear();
            cityField.clear();
        });

        // Event Handler für den Export-Button
        exportButton.setOnAction(event -> exportToExcel(tableView));
    }

    private void exportToExcel(TableView<String[]> tableView) {
        try {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("ExportedData");

            // Erstellen Sie die Überschriftenzeile
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Alter");
            headerRow.createCell(2).setCellValue("Stadt");

            // Iterieren Sie durch die TableView-Daten und schreiben Sie sie in die Tabelle
            int rowNum = 1; // Starte bei Zeile 1, da Zeile 0 die Überschriften sind
            for (String[] rowData : tableView.getItems()) {
                Row dataRow = sheet.createRow(rowNum++);
                for (int i = 0; i < rowData.length; i++) {
                    Cell cell = dataRow.createCell(i);
                    cell.setCellValue(rowData[i]);
                }
            }

            // Schreiben Sie das Workbook in eine Excel-Datei
            try (FileOutputStream outputStream = new FileOutputStream("ExportedData.xlsx")) {
                workbook.write(outputStream);
            }

            // Schließen Sie das Workbook
            workbook.close();

            System.out.println("Daten wurden erfolgreich nach Excel exportiert.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}