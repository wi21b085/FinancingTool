package com.example.financingtool;

import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.Cell;

public interface IAllExcelRegisterCards {

    // Maria P, Maria M

    static public boolean testPercentageRange(String value) {
        if (value.isEmpty()) {
            return true;
        }

        try {
            double doubleValue = Double.parseDouble(value);
            return doubleValue >= 0 && doubleValue <= 1;
        } catch (NumberFormatException e) {
            // Handle the case where parsing to double fails
            return false;
        }
    }


// Maria P, Maria M

    static public String parsePercentageValue(String value) {
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

    // Maria M, Maria P, Hadi
    static boolean isNumericStr(String str) {
        try {
            double numericValue = Double.parseDouble(str);

            // Wenn die Konvertierung erfolgreich ist, ist der String numerisch
            return true;
        } catch (NumberFormatException e) {
            // Wenn eine NumberFormatException auftritt, ist der String nicht numerisch
            return false;
        }
    }

    // letztes
    static public void openNewJavaFXWindow() {
        Stage newStage = new Stage();

        // Button für die Konvertierung in Word im neuen Fenster
        javafx.scene.control.Button convertToWordButton = new javafx.scene.control.Button("Konvertierung in eine PDF");
        convertToWordButton.setOnAction(e -> ExcelToWordConverter.convertWordToPDF());

        // Layout für das neue Fenster
        VBox newRoot = new VBox(10);
        newRoot.setAlignment(Pos.CENTER);
        newRoot.getChildren().add(convertToWordButton);

        Scene newScene = new Scene(newRoot, 300, 200);
        newStage.setTitle("PDF Konvertierung");
        newStage.setScene(newScene);
        newStage.show();
    }

    static public boolean emptyCell(Cell cell) {
        if (cell == null) {
            return true;
        } else {
            return false;
        }


    }
}