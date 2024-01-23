package com.example.financingtool;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class GIKController implements IAllExcelRegisterCards{
    private boolean valid=true;
    private String notInRange="Achtung die Werte müssen zwischen 0%-100% bzw zwischen 0.0-1.0 betragen.";
    @FXML
    private Label resultLabel;
    private Stage stage;
    @FXML
    private Label welcomeText;

    private boolean updateD2=true;
    private boolean updateD3to9=true;
    private boolean updateD10=true;
    @FXML
    private Button weiterButton;

    @FXML
    TextField userInputField1;

    @FXML
    TextField userInputField2;

    @FXML
    TextField userInputField3;

    @FXML
    TextField userInputField4;

    @FXML
    TextField userInputField5;

    @FXML
    TextField userInputField6;

    @FXML
    TextField userInputField7;

    @FXML
    TextField userInputField8;

    @FXML
    TextField userInputField9;
    private TextField[] userInputFields = new TextField[9];

    @FXML
    private Button updateRangeButton;

    @FXML
    private TextField userInputFieldD2;

    @FXML
    private TextField userInputFieldD3to9;

    @FXML
    private TextField userInputFieldD10;

    @FXML
    private Button updateButtonD;
    private String errorNotNumeric="Achtung, die Werte müssen numerisch sein";

    private String sheetName = "Gesamtinvestitionskosten";

    static MV_MH_Controller mvMhController = new MV_MH_Controller();
    static ExecutiveSummary executiveSummary = new ExecutiveSummary();


    public static void setMV_MH_Controller(MV_MH_Controller mvMhController) {
        GIKController.mvMhController = mvMhController;
    }
    public static void setExecutiveSummary(ExecutiveSummary executiveSummary) {
       GIKController.executiveSummary=executiveSummary;
    }

    @FXML
    public void initialize() {
        // Initialisiere das Array in der initialize-Methode
        userInputFields = new TextField[]{userInputField1, userInputField2, userInputField3,
                userInputField4, userInputField5, userInputField6,
                userInputField7, userInputField8, userInputField9};
    }
    /*public void weiter(ActionEvent actionEvent) {


    }*/

    public void updateRange(ActionEvent actionEvent) {

        String[] newValues = new String[9];
        int countNonNumeric = 0;
        String nonNumericValue = "";
        for (int i = 0; i < 9; i++) {
            if (IAllExcelRegisterCards.isNumericStr(userInputFields[i].getText()) || userInputFields[i].getText().trim().isEmpty()) {
                newValues[i] = userInputFields[i].getText();
                // Setze das Fehlerlabel auf leer, da keine Fehlermeldung vorliegt

                valid=true;
            } else {

                countNonNumeric++;
                nonNumericValue = userInputFields[i].getText();

            }
        }
        if (countNonNumeric > 0) {
            resultLabel.setText("Ein oder mehrere Werte sind ungültig.");
            valid=false;
        } else {
            Update.updateRangeOfCells(newValues,2,10,1, resultLabel, sheetName);
        }

        // Hole den Wert aus userInputField8
        String newValue = userInputField9.getText();
        System.out.println(userInputField9.getText());
        setMV_MH_Controller(mvMhController);
        EventBus.getInstance().publish("updateFK", newValue);

        updateD2 = validateAndUpdate(userInputFieldD2.getText(), 21, 1, resultLabel) && validateAndUpdate(userInputFieldD2.getText(), 2, 3,resultLabel);
        updateD3to9 = validateAndUpdate(userInputFieldD3to9.getText(), 20, 1, resultLabel);
        updateD10 = validateAndUpdate(userInputFieldD10.getText(), 10, 3, resultLabel);


        if (updateD2 && updateD3to9 && updateD10&valid==true) {
            resultLabel.setText("Daten erfolgreich aktualisiert.");
          getCellData();

         //   executiveSummary.setDatenausGIK(gik);

        }


    }
    private void getCellData() {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Gesamtinvestitionskosten";
            int rowIdx = 13;
            int colIdx = 4;

            // FileInputStream und Workbook hier erstellen
            try (FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
                 Workbook workbook = new XSSFWorkbook(fileInputStream)) {

                Sheet sheet = workbook.getSheet(sheetName);

                Row row = sheet.getRow(rowIdx);
                Cell cell = row.getCell(colIdx);
                String gikCell = Double.toString(cell.getNumericCellValue());
                System.out.println(gikCell);
                //fk.setText(fkCell);
             // executiveSummary.setDatenausGIK(gikCell);
              //  executiveSummary.setDaten();
            } catch (NumberFormatException | IOException e) {
                e.printStackTrace();
                //resultLabel.setText("Fehler bei der Aktualisierung.");
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
//Funktion von hier nach oben hinzugefügt zusätzlich
    public void updateD(ActionEvent actionEvent) {

         updateD2 = validateAndUpdate(userInputFieldD2.getText(), 21, 1, resultLabel) && validateAndUpdate(userInputFieldD2.getText(), 2, 3,resultLabel);
         updateD3to9 = validateAndUpdate(userInputFieldD3to9.getText(), 20, 1, resultLabel);
         updateD10 = validateAndUpdate(userInputFieldD10.getText(), 10, 3, resultLabel);


        if (updateD2 && updateD3to9 && updateD10) {
            resultLabel.setText("Daten erfolgreich aktualisiert.");
        }
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
            // Handle the case where parsing to double fail
            errorLabel.setText("Achtung: Die Eingabe ist keine gültige Zahl.");
            resultLabel.setText("Achtung: Die Eingabe ist keine gültige Zahl.");
            System.out.println("Achtung: Die Eingabe ist keine gültige Zahl.");
        }
        return false;
    }

    public void weiter(ActionEvent actionEvent) {
    }
}
