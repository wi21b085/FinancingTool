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

public class MV_MH_Controller extends Application implements IAllExcelRegisterCards {

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
                String[] vals1 = new String[1];
                vals1[0] = ik.getText();
                Update.updateRangeOfCells(vals1, 1, 1, 1, new Label(), sheet);

                int tranchen = Integer.parseInt(tranche.getValue());
                String[] vals2 = new String[2+tranchen];
                vals2[0] = ek.getText();
                vals2[1] = btvg.getText();

                for (int i = 0; i < tranchen; i++) {
                    for (Node node : pane.getChildren()) {
                        if(node instanceof TextField){
                            TextField text = (TextField) node;
                            if(text.getId().contains("text"+i)){
                                vals2[i+2] = text.getText();
                            }
                        }
                    }
                }

                Update.updateRangeOfCells(vals2, 1, 4, 4, new Label(), sheet);

            }
        } catch (Exception e) {
            System.out.println("Tranchen müssen ausgewählt sein!");;
        }

    }

    @FXML
    public void selectChoice() {
        int val = Integer.parseInt(tranche.getValue());
        System.out.println(tranche.getValue());
        Label[] lt = new Label[val];
        TextField[] tt = new TextField[val];

        List<Node> nodes = new ArrayList<>();
        for (Node node : pane.getChildren()) {
            if(node instanceof Label){
                Label label = (Label) node;
                if(label.getText().contains("Tranche ")){
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
            lt[i] = new Label("Tranche "+ (i+1));
            lt[i].setLayoutX(300);
            lt[i].setLayoutY(128+30*(i+1));
            //lt[i].setId("label"+i);

            tt[i] = new TextField();
            tt[i].setLayoutX(400.0);
            tt[i].setLayoutY(120+30*(i+1));
            tt[i].setId("text"+i);

            pane.getChildren().add(lt[i]);
            pane.getChildren().add(tt[i]);
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

    @Override
    public void start(Stage stage) throws Exception {


    }




    public void weiterMariaM(ActionEvent actionEvent) {
        Weiter.weiter(weiterButton, MainApplication.class);
    }
}