package com.example.financingtool;

import javafx.application.Application;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.Node;
import javafx.scene.control.Button;
import javafx.scene.control.ChoiceBox;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.Pane;
import javafx.scene.text.Text;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class MV_MH_Controller implements IAllExcelRegisterCards {

    @FXML
    private Pane pane;
    @FXML
    private TextField ik;
    @FXML
    private TextField ek;
    @FXML
    private Text fk;

    @FXML
    private TextField btvg;
    @FXML
    private ChoiceBox<String> tranche = new ChoiceBox<>();

    //Maria M
    @FXML
    private Button weiterButton;



    @FXML
    protected void continueClick() {
        try {
            if(!tranche.getValue().isEmpty()){
                String sheet = "Mittelverwendung - Mittelherkun";
                //aktuell nur mit 1 oder 2 Tranchen machbar
                int tranchen = Integer.parseInt(tranche.getValue());

//                String[] ik = new String[1];
//                if (IAllExcelRegisterCards.isNumericStr(this.ik.getText()) || this.ik.getText().trim().isEmpty()) {
//                    ik[0] = this.ik.getText();
//                    Update.updateRangeOfCells(ik, 1, 1, 1, new Label(), sheet);
//                }

//                String[] ek = new String[1];
//                if (IAllExcelRegisterCards.isNumericStr(this.ek.getText()) || this.ek.getText().trim().isEmpty()) {
//                    ek[0] = this.ek.getText();
//                    Update.updateRangeOfCells(ek, 1, 1, 4, new Label(), sheet);
//                }

//                String[] btvg = new String[1];
//                if (IAllExcelRegisterCards.isNumericStr(this.btvg.getText()) || this.btvg.getText().trim().isEmpty()) {
//                    btvg[0] = this.btvg.getText();
//                    Update.updateRangeOfCells(btvg, tranchen+2, 4, 4, new Label(), sheet);
//                }

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
        }

    }

    @FXML
    public void selectChoice() {
        int val = Integer.parseInt(tranche.getValue());
        System.out.println(tranche.getValue());
        TextField[] lt = new TextField[val];
        TextField[] tt = new TextField[val];

        List<Node> nodes = new ArrayList<>();
        for (Node node : pane.getChildren()) {
            if(node instanceof TextField){
                TextField label = (TextField) node;
                if(label.getId().contains("label")){
                    nodes.add(node);
                }
            }
            if(node instanceof TextField){
                TextField text = (TextField) node;
                if(text.getId().contains("text")){
                    nodes.add(node);
                }
            }
        }
        pane.getChildren().removeAll(nodes);

        for (int i = 0; i < val; i++) {
            lt[i] = new TextField("Fremdkapital Tranche "+ (i+1));
            lt[i].setLayoutX(240);
            lt[i].setLayoutY(120+30*(i+1));
            lt[i].setId("label"+i);
            lt[i].setPromptText("Bezeichnung eingeben");

            tt[i] = new TextField();
            tt[i].setLayoutX(400.0);
            tt[i].setLayoutY(120+30*(i+1));
            tt[i].setId("text"+i);
            tt[i].setPromptText("Betrag eingeben");

            pane.getChildren().add(lt[i]);
            pane.getChildren().add(tt[i]);
        }
    }

    public void setCell(TextField tf, int row, int coll) {
        String sheet = "Mittelverwendung - Mittelherkun";
        String[] field = new String[1];
        if (IAllExcelRegisterCards.isNumericStr(tf.getText()) || tf.getText().trim().isEmpty()) {
            field[0] = tf.getText();
            Update.updateRangeOfCells(field, row, row, coll, new Label(), sheet);
        }
    }

    public void initialize() throws Exception {
        getFK();
    }

    public void getFK() throws Exception {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Gesamtinvestitionskosten";
            int rowIdx = 9;
            int colIdx = 1;

            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheet(sheetName);

            Row row = sheet.getRow(rowIdx);
            Cell cell = row.getCell(colIdx);
            String fkCell = Double.toString(cell.getNumericCellValue());
            System.out.println(fkCell);
            fk.setText(fkCell);
            fileInputStream.close();

            workbook.close();


        } catch (NumberFormatException | IOException e) {
            e.printStackTrace();
            //resultLabel.setText("Fehler bei der Aktualisierung.");
        }
    }

    public void weiterMariaM(ActionEvent actionEvent) {
        Weiter.weiter(weiterButton,Textfeld.class);
        ExcelToWordConverter.exportExcelToWord();
    }
}