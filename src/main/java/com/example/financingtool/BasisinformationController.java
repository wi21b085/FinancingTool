package com.example.financingtool;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;

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

    //Maria M
    @FXML
    private Button weiterButton;

    private static final String FILE_NAME = "SEPJ-Rechnungen.xlsx";
    private static final String SHEET_NAME = "Basisinformationen";

    static ExecutiveSummary executiveSummary = new ExecutiveSummary();


    public static void setExecutiveSummary(ExecutiveSummary executiveSummary) {
       BasisinformationController.executiveSummary=executiveSummary;
    }

    public void onHelloButtonClick(ActionEvent actionEvent) {
    }

    //firmennamen in das excel eintragen
    public void submit() {
        //leere werte
        if(kaufpreis.getText().isEmpty() || groesse.getText().isEmpty() || nutzflaeche.getText().isEmpty()
                || wohneinheiten.getText().isEmpty() || garage.getText().isEmpty() || gik.getText().isEmpty()
                || verkaufserloes.getText().isEmpty() || gewinn.getText().isEmpty() || beginn.getText().isEmpty()
                ||ende.getText().isEmpty() || roi.getText().isEmpty() ){
            resultLabel.setText("Daten unvollständig");
            // System.out.println("Daten unvollständig");
        }
        //ungültige werte
        else if (!IAllExcelRegisterCards.isNumericStr(kaufpreis.getText()) ||
                 !IAllExcelRegisterCards.isNumericStr(groesse.getText()) ||
                 !IAllExcelRegisterCards.isNumericStr(nutzflaeche.getText()) ||
                 !IAllExcelRegisterCards.isNumericStr(wohneinheiten.getText()) ||
                 !IAllExcelRegisterCards.isNumericStr(garage.getText()) ||
                 !IAllExcelRegisterCards.isNumericStr(gik.getText()) ||
                 !IAllExcelRegisterCards.isNumericStr(verkaufserloes.getText()) ||
                 !IAllExcelRegisterCards.isNumericStr(verkaufserloes.getText())
        ){
            resultLabel.setText("Achtung, bitte geben Sie gültige Daten an");
            return;
        }
        //gültige werte
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
            setExecutiveSummary(executiveSummary);
            System.out.println("Daten aus Basisinformation gesendet gesendet: ");
            executiveSummary.setDatenausBas(kaufpreis.getText(),groesse.getText(),wohneinheiten.getText(),garage.getText(),beginn.getText(),ende.getText());

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

            //alle daten leer
            if (kaufpreis.getText().isEmpty() && groesse.getText().isEmpty() && nutzflaeche.getText().isEmpty() &&
                    wohneinheiten.getText().isEmpty() && garage.getText().isEmpty() && gik.getText().isEmpty()
                    && verkaufserloes.getText().isEmpty() && gewinn.getText().isEmpty() && beginn.getText().isEmpty()
                    && ende.getText().isEmpty() && roi.getText().isEmpty()) {
                resultLabel.setText("Daten erforderlich zum Aktualisieren");
                System.out.println("Alle Felder leer");
                return;
            }

            // Überprüfen, ob leer + Datentyp stimmt
            if (!kaufpreis.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(kaufpreis.getText())) {
                updateCellValue(sheet, 1, 1, kaufpreis.getText());


            } else if (!IAllExcelRegisterCards.isNumericStr(kaufpreis.getText())) {
                resultLabel.setText("Gültige Daten erforderlich");
                return;
            }

            if (!groesse.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(groesse.getText())) {
                updateCellValue(sheet, 2, 1, groesse.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(groesse.getText())) {
                resultLabel.setText("Gültige Daten erforderlich");
                return;
            }

            if (!nutzflaeche.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(nutzflaeche.getText())) {
                updateCellValue(sheet, 3, 1, nutzflaeche.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(nutzflaeche.getText())) {
                resultLabel.setText("Gültige Daten erforderlich");
                return;
            }

            if (!wohneinheiten.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(wohneinheiten.getText())) {
                updateCellValue(sheet, 4, 1, wohneinheiten.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(wohneinheiten.getText())) {
                resultLabel.setText("Gültige Daten erforderlich");
                return;
            }

            if (!garage.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(garage.getText())) {
                updateCellValue(sheet, 5, 1, garage.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(garage.getText())) {
                resultLabel.setText("Gültige Daten erforderlich");
                return;
            }

            if (!gik.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(gik.getText())) {
                updateCellValue(sheet, 6, 1, gik.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(gik.getText())) {
                resultLabel.setText("Gültige Daten erforderlich");
                return;
            }

            if (!verkaufserloes.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(verkaufserloes.getText())) {
                updateCellValue(sheet, 7, 1, verkaufserloes.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(verkaufserloes.getText())) {
                resultLabel.setText("Gültige Daten erforderlich");
                return;
            }

            if (!gewinn.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(gewinn.getText())) {
                updateCellValue(sheet, 8, 1, gewinn.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(gewinn.getText())) {
                resultLabel.setText("Gültige Daten erforderlich");
                return;
            }


            if (!beginn.getText().isEmpty()) { // Datum kann String + Double sein
                updateCellValue(sheet, 9, 1, beginn.getText());
            }

            if (!ende.getText().isEmpty()) { // Datum kann String + Double sein
                updateCellValue(sheet, 10, 1, ende.getText());
            }

            if (!roi.getText().isEmpty() && IAllExcelRegisterCards.isNumericStr(roi.getText())) {
                updateCellValue(sheet, 11, 1, roi.getText());
            } else if (!IAllExcelRegisterCards.isNumericStr(roi.getText())) {
                resultLabel.setText("Gültige Daten erforderlich");
                return;
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



  /*  public void weiter(ActionEvent actionEvent) {
        //ExcelToWordConverter.exportExcelToWord("Basisinformation");

        Weiter.weiter(weiterButton, GIKtoExcel.class);
    }*/
}
