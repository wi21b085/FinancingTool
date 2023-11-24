package com.example.financingtool;

import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TableView;
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
        // Erstelle eine einfache JavaFX-Anwendung mit einer TableView und einem Button
        TableView<String[]> tableView = new TableView<>();
        Button exportButton = new Button("Daten exportieren");

        VBox root = new VBox(tableView, exportButton);
        Scene scene = new Scene(root, 400, 300);

        primaryStage.setTitle("JavaFX to Excel");
        primaryStage.setScene(scene);
        primaryStage.show();

        // Beispiel-Daten für die TableView
        tableView.getItems().addAll(
                new String[]{"Name", "Alter", "Stadt"},
                new String[]{"John Doe", "25", "New York"},
                new String[]{"Jane Doe", "30", "London"}
        );

        // Event Handler für den Export-Button
        exportButton.setOnAction(event -> exportToExcel(tableView));
    }

    private void exportToExcel(TableView<String[]> tableView) {
        try {
            // Erstelle ein neues Workbook und eine Tabelle
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("ExportedData");

            // Iteriere durch die TableView-Daten und schreibe sie in die Tabelle
            int rowNum = 0;
            for (String[] rowData : tableView.getItems()) {
                Row excelRow = sheet.createRow(rowNum++);
                int colNum = 0;
                for (String cellData : rowData) {
                    Cell cell = excelRow.createCell(colNum++);
                    cell.setCellValue(cellData);
                }
            }

            // Schreibe das Workbook in eine Excel-Datei
            try (FileOutputStream outputStream = new FileOutputStream("ExportedData.xlsx")) {
                workbook.write(outputStream);
            }

            // Schließe das Workbook
            workbook.close();

            System.out.println("Daten wurden erfolgreich nach Excel exportiert.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
