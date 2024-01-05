package com.example.financingtool;

import com.itextpdf.kernel.color.Lab;
import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javafx.scene.paint.Color;

public class GIKtoExcel extends Application  implements IAllExcelRegisterCards{

    public Label resultLabel=new Label("Aktueller Wert: ");
    private TextField[] userInputFields = new TextField[9];
    private TextField userInputFieldD2 = new TextField();
    private TextField userInputFieldD3to9 = new TextField();
    private TextField userInputFieldD10 = new TextField();

    private Label[] errorLabels = new Label[9];
    private Label errorLabelD2 = new Label();
    private Label errorLabelD3to9 = new Label();
    private Label errorLabelD10 = new Label();

    private String notInRange="Achtung die Werte müssen zwischen 0%-100% bzw zwischen 0.0-1.0 betragen.";
    private String errorNotNumeric="Achtung, die Werte müssen numerisch sein";
    private boolean updateD2=true;
    private boolean updateD3to9=true;
    private boolean updateD10=true;
    private boolean valid=true;


    private String sheetName = "Gesamtinvestitionskosten";
    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(GIKtoExcel.class.getResource("gik.fxml"));
        ScrollPane scrollPane = fxmlLoader.load();
        VBox root = (VBox) scrollPane.getContent();

        Scene scene = new Scene(scrollPane, 1280, 720);
        stage.setTitle("GIKtoExcel");


        // Erstellen Sie 10 Textfelder für die Benutzereingabe
        for (int i = 0; i < 9; i++) {
            userInputFields[i] = new TextField();
            errorLabels[i] = new Label(); // Initialisiere das Fehlerlabel
            if (i < 3) {
                userInputFields[i].setPromptText("KB0" + i + " netto eingeben");
            } else {
                userInputFields[i].setPromptText("KB0" + (i + 3) + " netto eingeben");
            }
        }

        // Button hinzufügen, um den Bereich von B3 bis B10 zu aktualisieren
        javafx.scene.control.Button updateRangeButton = new javafx.scene.control.Button("Bereich aktualisieren");
        updateRangeButton.setOnAction(e -> {
            String[] newValues = new String[9];
            int countNonNumeric = 0;
            String nonNumericValue = "";
            for (int i = 0; i < 9; i++) {
                if (IAllExcelRegisterCards.isNumericStr(userInputFields[i].getText()) || userInputFields[i].getText().trim().isEmpty()) {
                    newValues[i] = userInputFields[i].getText();
                    // Setze das Fehlerlabel auf leer, da keine Fehlermeldung vorliegt
                    errorLabels[i].setText("");
                    valid=true;
                } else {

                    countNonNumeric++;
                    nonNumericValue = userInputFields[i].getText();
                    errorLabels[i].setText("Achtung: Die Werte müssen numerisch sein. Fehler bei " + nonNumericValue);
                }
            }
            if (countNonNumeric > 0) {
                resultLabel.setText("Ein oder mehrere Werte sind ungültig.");
                valid=false;
            } else {
                Update.updateRangeOfCells(newValues,1,9,1, resultLabel, sheetName);
            }
        });

        // Button hinzufügen, um Zellen D2, D3 bis D9 und D10 zu aktualisieren
        javafx.scene.control.Button updateButtonD = new javafx.scene.control.Button("UST aktualisieren");
        userInputFieldD2.setPromptText("UST Grund");
        errorLabelD2.setTextFill(Color.RED); // Setze die Textfarbe auf Rot
        userInputFieldD3to9.setPromptText("Genereller UST");
        errorLabelD3to9.setTextFill(Color.RED);
        userInputFieldD10.setPromptText("UST Finanzierung");
        errorLabelD10.setTextFill(Color.RED);
        updateButtonD.setOnAction(e -> {
         updateD2 = validateAndUpdate(userInputFieldD2.getText(), 20, 1, errorLabelD2) && validateAndUpdate(userInputFieldD2.getText(), 1, 3, errorLabelD2);
         updateD3to9 = validateAndUpdate(userInputFieldD3to9.getText(), 19, 1, errorLabelD3to9);
         updateD10 = validateAndUpdate(userInputFieldD10.getText(), 9, 3, errorLabelD10);



            if (updateD2 && updateD3to9 && updateD10) {
                resultLabel.setText("Daten erfolgreich aktualisiert.");
            }
        });



        // Füge alle UI-Elemente zum Root-VBox hinzu
        for (int i = 0; i < 9; i++) {
            root.getChildren().add(userInputFields[i]);
        }

        javafx.scene.control.Button weiterButton = new javafx.scene.control.Button("Weiter");
        System.out.println(updateD2+" "+updateD10+" "+updateD3to9+" "+valid);
        if (updateD2 && updateD3to9 && updateD10 && valid) {
            weiterButton.setOnAction(e ->

            {
                Weiter.weiter(weiterButton, MV_MH.class);
                //ExcelToWordConverter.exportExcelToWord("Gesamtinvestitionskosten");
                //ExcelToWordConverter.openNewJavaFXWindow();
            });



        }else {
            System.out.println("Test");
            FileInputStream fileInputStream = new FileInputStream(new File("src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx"));
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheet(sheetName);
            Cell [] cells=new Cell[9];
            for(int i=0;i<cells.length;i++) {
              cells[i]=sheet.getRow(i+1).getCell(1);

                if (IAllExcelRegisterCards.emptyCell(cells[i])) {


                    System.out.println("Achtung, Sie haben nichts eingegeben und es ist kein Wert vorhanden.");
                }
            }
            Cell[] cellsforUst= new Cell[3];
            cellsforUst[0]=sheet.getRow(19).getCell(1);
            cellsforUst[1]=sheet.getRow(20).getCell(1);
            cellsforUst[2]=sheet.getRow(9).getCell(3);
            for(int i=0;i<cellsforUst.length;i++){
                System.out.println("Achtung, Sie haben nichts eingegeben und es ist kein Wert vorhanden.");
            }

            weiterButton.setOnAction(e ->resultLabel.setText("Achtung, es kann keine Konvertieurng ausgeführt werden, solange die Daten nicht valide sind."));
        }
        root.getChildren().addAll(resultLabel, updateRangeButton, userInputFieldD2, userInputFieldD3to9, userInputFieldD10, updateButtonD, weiterButton);
        // Setze die Szene und zeige die Bühne
        stage.setScene(scene);
        stage.show();
    }


    public static void main(String[] args) {
        launch(args);

    }
    private void updateCellD(int rowIdx, int colIdx, String newValue) {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
           // String sheetName = "Gesamtinvestitionskosten";

            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheet(sheetName);

            // Überprüfen Sie, ob die Zeichenkette nicht leer ist und nicht null ist, bevor Sie sie parsen
            if (newValue != null && !newValue.isEmpty()) {
                double newCellValue = Double.parseDouble(newValue);
               Update.updateCellValue(sheet, rowIdx, colIdx, newCellValue);
            }

            // Automatische Auswertung der Formeln im gesamten Arbeitsblatt
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();

            fileInputStream.close();

            FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();

            // resultLabel.setText("Zelle erfolgreich aktualisiert.");
            // System.out.println("Zelle erfolgreich aktualisiert.");

        } catch (NumberFormatException | IOException e) {
            e.printStackTrace();
            resultLabel.setText("Fehler bei der Aktualisierung von Zelle D" + rowIdx);
        }
    }
    private boolean validateAndUpdate(String userInput, int rowIdx, int colIdx, Label errorLabel) {
        try{
            String parsedValue = IAllExcelRegisterCards.parsePercentageValue(userInput);
            System.out.println(parsedValue);


            boolean isNumeric = IAllExcelRegisterCards.isNumericStr(parsedValue);
            boolean isValidRange = IAllExcelRegisterCards.testPercentageRange(parsedValue);

            if (isNumeric || userInput.trim().isEmpty()) {
                System.out.println("Ist numerisch");
                if (isValidRange) {
                    System.out.println("Ist zwischen 0 und 1");
                    errorLabel.setText(""); // Set the error label to empty since there is no error message
                    updateCellD(rowIdx, colIdx, parsedValue);
                    return true;
                } else {

                    System.out.println("Ist nicht in der Range");
                    errorLabel.setText(notInRange);
                    resultLabel.setText(notInRange);
                    System.out.println(notInRange);
                    return false;
                }
            } else {

                errorLabel.setText(errorNotNumeric);
                resultLabel.setText(errorNotNumeric);
                System.out.println(errorNotNumeric);
                return false;
            }
        }catch (NumberFormatException e) {
            // Handle the case where parsing to double fails
            errorLabel.setText("Achtung: Die Eingabe ist keine gültige Zahl.");
            resultLabel.setText("Achtung: Die Eingabe ist keine gültige Zahl.");
            System.out.println("Achtung: Die Eingabe ist keine gültige Zahl.");
        }
        return false;
    }
}
