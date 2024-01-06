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


public class WidmungController implements IAllExcelRegisterCards {

    @FXML
    private ChoiceBox<String> fw = new ChoiceBox<>();
    @FXML
    private ChoiceBox<String> bk = new ChoiceBox<>();
    @FXML
    private ChoiceBox<String> bw = new ChoiceBox<>();
    @FXML
    private ChoiceBox<String> bb1 = new ChoiceBox<>();
    @FXML
    private ChoiceBox<String> bb2 = new ChoiceBox<>();
    @FXML
    private TextField bs;

    @FXML
    private Text ad;
    private String adCell;

    @FXML
    protected void continueClick() {
        //check(fw);
        /*try {
            if(!tranche.getValue().isEmpty()){
                String sheet = "Mittelverwendung - Mittelherkun";
                //aktuell nur mit 1 oder 2 Tranchen machbar
                int tranchen = Integer.parseInt(tranche.getValue());

                String[] vals = new String[tranchen];
                String[] bez = new String[tranchen];

                boolean numericTest = true;
                for (Node node : pane.getChildren()) {
                    for (int i = 0; i < tranchen; i++) {
                        if(node instanceof TextField){
                            TextField text = (TextField) node;
                            if(text.getId().contains("text"+i)){
                                if (IAllExcelRegisterCards.isNumericStr(text.getText()) || text.getText().trim().isEmpty()) {
                                    vals[i] = text.getText();
                                } else {
                                    numericTest = false;
                                }
                            } else if(text.getId().contains("label"+i)){
                                if (!text.getText().trim().isEmpty()) {
                                    bez[i] = text.getText();
                                }
                            }
                        }
                    }
                }

                if(numericTest) {
                    switch (tranchen) {
                        case 1:
                            setCell(this.ik, 10, 1);
                            setCell(this.ek, 10, 4);
                            setCell(this.btvg, 12, 4);
                            Update.updateRangeOfCells(vals, 11, 11, 4, new Label(), sheet);
                            Update.updateRangeOfCellsString(bez, 11, 11, 3, new Label(), sheet);
                            break;
                        case 2:
                            setCell(this.ik, 1, 1);
                            setCell(this.ek, 1, 4);
                            setCell(this.btvg, 4, 4);
                            Update.updateRangeOfCells(vals, 2, 3, 4, new Label(), sheet);
                            Update.updateRangeOfCellsString(bez, 2, 3, 3, new Label(), sheet);
                            break;
                        case 3:
                            setCell(this.ik, 18, 1);
                            setCell(this.ek, 18, 4);
                            setCell(this.btvg, 22, 4);
                            Update.updateRangeOfCells(vals, 19, 21, 4, new Label(), sheet);
                            Update.updateRangeOfCellsString(bez, 19, 21, 3, new Label(), sheet);
                            break;
                        case 4:
                            setCell(this.ik, 28, 1);
                            setCell(this.ek, 28, 4);
                            setCell(this.btvg, 33, 4);
                            Update.updateRangeOfCells(vals, 29, 32, 4, new Label(), sheet);
                            Update.updateRangeOfCellsString(bez, 29, 32, 3, new Label(), sheet);
                            break;
                        case 5:
                            setCell(this.ik, 39, 1);
                            setCell(this.ek, 39, 4);
                            setCell(this.btvg, 45, 4);
                            Update.updateRangeOfCells(vals, 40, 44, 4, new Label(), sheet);
                            Update.updateRangeOfCellsString(bez, 40, 44, 3, new Label(), sheet);
                            break;
                    }
                    //Update.updateRangeOfCells(bez, 2, tranchen, 3, new Label(), sheet);
                }
            }
        } catch (Exception e) {
            System.out.println("Tranchen müssen ausgewählt sein!");;
        }*/

        executePy(adCell);
    }

    private boolean check(ChoiceBox<String> fw) {
        try {
            if(!fw.getValue().isEmpty()){
                String sheet = "Basisinformation";
                //aktuell nur mit 1 oder 2 Tranchen machbar
                String widmung = fw.getValue();
                System.out.println(widmung);

                String[] ins = new String[1];

                boolean numericTest = true;
//                for (Node node : pane.getChildren()) {
//                    for (int i = 0; i < tranchen; i++) {
//                        if(node instanceof TextField){
//                            TextField text = (TextField) node;
//                            if(text.getId().contains("text"+i)){
//                                if (IAllExcelRegisterCards.isNumericStr(text.getText()) || text.getText().trim().isEmpty()) {
//                                    vals[i] = text.getText();
//                                } else {
//                                    numericTest = false;
//                                }
//                            } else if(text.getId().contains("label"+i)){
//                                if (!text.getText().trim().isEmpty()) {
//                                    bez[i] = text.getText();
//                                }
//                            }
//                        }
//                    }
                return true;
                }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return false;
    }

    public void executePy(String name) {
        String s = null;
        try {
            //Runtime.getRuntime().exec("src\\main\\resources\\com\\example\\financingtool\\script.bat");
            //String[] cmd = { "python", "src\\main\\resources\\com\\example\\financingtool\\widmung.py", name};
            //Process p = Runtime.getRuntime().exec(cmd);
            Process p = Runtime.getRuntime().exec("cmd /C start src\\main\\resources\\com\\example\\financingtool\\script.bat \""+name+"\"");

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
        System.out.println(adCell);
    }

    public void getAddress() {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Basisinformation";
            int rowIdx = 2;
            int colIdx = 8;

            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheet(sheetName);

            Row row = sheet.getRow(rowIdx);
            Cell cell = row.getCell(colIdx);
            adCell = cell.getStringCellValue();
            System.out.println(adCell);
            ad.setText(adCell);
            fileInputStream.close();

            workbook.close();


        } catch (NumberFormatException | IOException e) {
            e.printStackTrace();
            //resultLabel.setText("Fehler bei der Aktualisierung.");
        }
    }

    public void setCell(TextField tf, int row, int coll) {
        String sheet = "Basisinformation";
        String[] field = new String[1];
        if (IAllExcelRegisterCards.isNumericStr(tf.getText()) || tf.getText().trim().isEmpty()) {
            field[0] = tf.getText();
            Update.updateRangeOfCells(field, row, row, coll, new Label(), sheet);
        }
    }
}
