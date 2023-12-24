package com.example.financingtool;

import javafx.application.Application;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.value.ObservableValue;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.scene.Scene;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.stage.Stage;
import javafx.util.Callback;

public class Wirtschaftlichkeitsrechnung extends Application {

    @Override
    public void start(Stage primaryStage) {
        // Spalten- und Zeilennamen
        String[] columnNames = {"Berechnung Anleger/ Eigennutzer", "Verhältnis", "Wohnfläche", "Netto (10%)", "Eigennutzerpreis", "Anlegerpreis"};
        String[] rowNames = {"Wohnungen", "Preise", "Eigennutzer", "Anleger", "Summe", "", "Parkplätze", "Preise", "Eigennutzer", "Anleger", "Summe"};

        ObservableList<TableRowData> data = FXCollections.observableArrayList();

        TableView<TableRowData> tableView = new TableView<>(data);

        for (String columnName : columnNames) {
            TableColumn<TableRowData, String> column = createColumn(columnName);
            tableView.getColumns().add(column);
        }

        for (String rowName : rowNames) {
            data.add(new TableRowData(rowName, columnNames.length));
        }

        tableView.setEditable(true);


        Scene scene = new Scene(tableView, 600, 300);
        primaryStage.setScene(scene);
        primaryStage.setTitle("Wirtschaftlichkeitsrechnung");
        primaryStage.show();
    }

    private TableColumn<TableRowData, String> createColumn(String columnName) {
        TableColumn<TableRowData, String> column = new TableColumn<>(columnName);

        column.setCellValueFactory(new Callback<TableColumn.CellDataFeatures<TableRowData, String>, ObservableValue<String>>() {
            @Override
            public ObservableValue<String> call(TableColumn.CellDataFeatures<TableRowData, String> param) {
                return param.getValue().getValue(columnName);
            }
        });

        // Bearbeitung von Textfeldern
        column.setCellFactory(TextFieldTableCell.forTableColumn());

        // Datenmodell aktualisieren, wenn der Benutzer Werte eingibt
        column.setOnEditCommit(event -> {
            event.getTableView().getItems().get(event.getTablePosition().getRow()).setValue(columnName, event.getNewValue());
        });

        return column;
    }

    public static void main(String[] args) {
        launch(args);
    }

    // Datenmodellklasse für eine Tabellenzeile
    public static class TableRowData {
        private final ObservableList<SimpleStringProperty> values;

        public TableRowData(String rowName, int numColumns) {
            values = FXCollections.observableArrayList();
            values.add(new SimpleStringProperty(rowName));
            for (int i = 1; i < numColumns; i++) {
                values.add(new SimpleStringProperty(""));
            }
        }

        public ObservableValue<String> getValue(String columnName) {
            int columnIndex = getIndexFromColumnName(columnName);
            return values.get(columnIndex);
        }

        public void setValue(String columnName, String value) {
            int columnIndex = getIndexFromColumnName(columnName);
            values.get(columnIndex).set(value);
        }

        private int getIndexFromColumnName(String columnName) {
            for (int i = 0; i < values.size(); i++) {
                if (columnName.equals("Berechnung Anleger/ Eigennutzer")) {
                    return 0;
                } else if (columnName.equals("Verhältnis")) {
                    return 1;
                } else if (columnName.equals("Wohnfläche")) {
                    return 2;
                } else if (columnName.equals("Netto (10%)")) {
                    return 3;
                } else if (columnName.equals("Eigennutzerpreis")) {
                    return 4;
                } else if (columnName.equals("Anlegerpreis")) {
                    return 5;
                }
            }
            return -1;
        }
    }
}
