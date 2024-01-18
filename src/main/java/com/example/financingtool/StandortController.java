package com.example.financingtool;

import javafx.fxml.FXML;
import javafx.scene.control.ChoiceBox;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.text.Text;
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
    private TextField L1;
    @FXML
    private TextField L2;
    @FXML
    private TextField L3;
    @FXML
    private TextField L4;
    @FXML
    private TextField L5;
    @FXML
    private TextField L6;

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
        check();

        executePy(adresse);
    }

    private void check() {

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
        getAddress();
        schFus.setPromptText("zu Fuß");
        schRad.setPromptText("mit dem Fahrrad");
        rstFus.setPromptText("zu Fuß");
        rstRad.setPromptText("mit dem Fahrrad");
        oefFus.setPromptText("zu Fuß");
        oefRad.setPromptText("mit dem Fahrrad");
        ezFus.setPromptText("zu Fuß");
        ezRad.setPromptText("mit dem Fahrrad");
        //resultLabel.setText("Testtext");
    }

    public void getAddress() {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Basisinformation";

            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheet(sheetName);

            Row row = sheet.getRow(2);
            Cell cell = row.getCell(8);
            adCell = cell.getStringCellValue();
            //System.out.println(adCell);
            dist.setText("Distanzen von '"+adCell+"' zu (Zahlen in Minuten):");

            row = sheet.getRow(3);
            cell = row.getCell(8);
            plzCell = Integer.parseInt(cell.getStringCellValue());

            row = sheet.getRow(4);
            cell = row.getCell(8);
            ortCell = cell.getStringCellValue();

            adresse = adCell + ", " + plzCell + " " + ortCell;
            fileInputStream.close();

            workbook.close();


        } catch (NumberFormatException | IOException e) {
            e.printStackTrace();
            //resultLabel.setText("Fehler bei der Aktualisierung.");
        }
    }

    public void setCell(String tf, int row, int coll) {
        String sheet = "Basisinformation";
        String[] field = new String[1];
        field[0] = tf;
        Update.updateRangeOfCellsString(field, row, row, coll, new Label(), sheet);
    }
}
