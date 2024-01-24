package com.example.financingtool;

import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.input.KeyEvent;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;


public class StandortController implements IAllExcelRegisterCards {

    @FXML
    private TextField schFus;
    @FXML
    private TextField schRad;
    @FXML
    private TextField rstFus;
    @FXML
    private TextField rstRad;
    @FXML
    private TextField oefFus;
    @FXML
    private TextField oefRad;
    @FXML
    private TextField ezFus;
    @FXML
    private TextField ezRad;
    @FXML
    private TextArea L1;
    @FXML
    private Button maps;

    @FXML
    private Label dist;
    private String adCell;
    private int plzCell;
    private String ortCell;
    private String adresse;
    @FXML
    private Label resultLabel;

    @FXML
    protected void continueClick() {
        boolean fwb = check();

        //System.out.println(fwb);
    }

    @FXML
    protected void openMaps() {
        executePy(adresse);
    }

    private boolean check() {
        try {
            if (
                    schFus.getText().isEmpty() || schRad.getText().isEmpty() ||
                    rstFus.getText().isEmpty() || rstRad.getText().isEmpty() ||
                    oefFus.getText().isEmpty() || oefRad.getText().isEmpty() ||
                    ezFus.getText().isEmpty() || ezRad.getText().isEmpty()
            ) {
                resultLabel.setText("Distanzen unvollständig");
                return false;
            } else if (
                    IAllExcelRegisterCards.isNumericStr(schFus.getText()) &&
                    IAllExcelRegisterCards.isNumericStr(schRad.getText()) &&
                    IAllExcelRegisterCards.isNumericStr(rstFus.getText()) &&
                    IAllExcelRegisterCards.isNumericStr(rstRad.getText()) &&
                    IAllExcelRegisterCards.isNumericStr(oefFus.getText()) &&
                    IAllExcelRegisterCards.isNumericStr(oefRad.getText()) &&
                    IAllExcelRegisterCards.isNumericStr(ezFus.getText()) &&
                    IAllExcelRegisterCards.isNumericStr(ezRad.getText())
            ) {
                resultLabel.setText("Daten wurden übernommen");
                setCell(schFus, 2, 1);
                setCell(schRad, 2, 2);
                setCell(rstFus, 3, 1);
                setCell(rstRad, 3, 2);
                setCell(oefFus, 4, 1);
                setCell(oefRad, 4, 2);
                setCell(ezFus, 5, 1);
                setCell(ezRad, 5, 2);

                if(!L1.getText().isBlank() && L1.getText() != null) {
                    String text = L1.getText();
                    String[] textSplit = text.split("\n", 0);
                    //System.out.println(textSplit);
                    int n = 2;
                    for (String a : textSplit) {
                        if(a.isEmpty())
                            setCell("\n", n++, 5);
                        else
                            setCell(a, n++, 5);
                    }
                    if (textSplit.length < 6) {
                        emptyCells(n);
                    }
                }
                return true;
            } else {
                resultLabel.setText("Bitte nur numerische Distanzwerte!");
                return false;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return false;
    }

    private void emptyCells(int n) {
        for (int i = n; i < 8; i++) {
            //System.out.println(i);
            setCell("\n", i, 5);
        }
    }

    public void executePy(String name) {
        String s = null;
        try {
            //Runtime.getRuntime().exec("src\\main\\resources\\com\ \example\\financingtool\\script.bat");
            //String[] cmd = { "python", "src\\main\\resources\\com\\example\\financingtool\\widmung.py", name};
            //Process p = Runtime.getRuntime().exec(cmd);
            Process p = Runtime.getRuntime().exec("cmd /C start src\\main\\resources\\com\\example\\financingtool\\script.bat standort.py \"" + name + "\"");

            BufferedReader stdInput = new BufferedReader(new
                    InputStreamReader(p.getInputStream()));

            BufferedReader stdError = new BufferedReader(new
                    InputStreamReader(p.getErrorStream()));

            // read the output from the command
            System.out.println("Here is the standard output of the command:\n");
            while ((s = stdInput.readLine()) != null) {
                System.out.println(s);
            }

            // read any errors from the attempted command
            System.out.println("Here is the standard error of the command (if any):\n");
            while ((s = stdError.readLine()) != null) {
                System.out.println(s);
            }

            //System.exit(0);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void initialize() {
        // Add a key pressed event handler
        L1.addEventFilter(KeyEvent.KEY_PRESSED, event -> {
            // Get the current number of lines
            int numberOfLines = L1.getParagraphs().size();

            // Allow Enter key
            if (event.getCode() == javafx.scene.input.KeyCode.ENTER) {
                // Check if the number of lines is already at the limit
                if (numberOfLines >= 6) {
                    event.consume(); // Consume the event to prevent further input
                }
            }
        });
        getAddress();
        EventBus.getInstance().subscribe("updateAddress", this::updateAddress);
        schFus.setPromptText("zu Fuß");
        schRad.setPromptText("mit dem Fahrrad");
        rstFus.setPromptText("zu Fuß");
        rstRad.setPromptText("mit dem Fahrrad");
        oefFus.setPromptText("zu Fuß");
        oefRad.setPromptText("mit dem Fahrrad");
        ezFus.setPromptText("zu Fuß");
        ezRad.setPromptText("mit dem Fahrrad");
        emptyCells(2);
        //resultLabel.setText("Testtext");
    }

    private void updateAddress(Object newValue) {
        System.out.println(newValue.toString());
        if(newValue.toString().isEmpty()){
            getAddress();
        }else {
            adCell = newValue.toString();
            getAddress();
            distSet();
        }
    }

    public void getAddress() {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Basisinformation";

            FileInputStream fileInputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheet(sheetName);

            Row row = sheet.getRow(2);
            Cell cell = row.getCell(8);
            adCell = cell.getStringCellValue();
            //System.out.println(adCell);

            row = sheet.getRow(3);
            cell = row.getCell(8);
            plzCell = Integer.parseInt(cell.getStringCellValue());

            row = sheet.getRow(4);
            cell = row.getCell(8);
            ortCell = cell.getStringCellValue();

            distSet();

            fileInputStream.close();

            workbook.close();


        } catch (NumberFormatException | IOException e) {
            e.printStackTrace();
            //resultLabel.setText("Fehler bei der Aktualisierung.");
        }
    }

    private void distSet() {
        dist.setText("Distanzen von '"+adCell+"' zu (Zahlen in Minuten):");

        adresse = adCell + ", " + plzCell + " " + ortCell;
    }

    public void setCell(String tf, int row, int coll) {
        String sheet = "Standort";
        String[] area = new String[1];
        area[0] = tf;
        Update.updateRangeOfCellsString(area, row, row, coll, new Label(), sheet);
    }

    public void setCell(TextField tf, int row, int coll) {
        String sheet = "Standort";
        String[] field = new String[1];
        field[0] = tf.getText();
        Update.updateRangeOfCells(field, row, row, coll, new Label(), sheet);
    }
}
