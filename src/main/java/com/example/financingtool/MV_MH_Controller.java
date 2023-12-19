package com.example.financingtool;

import javafx.application.Application;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.Node;
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

    public MV_MH_Controller() throws Exception {
    }

    @FXML
    protected void continueClick() {
        if(tranche.getValue().isEmpty()){
            System.out.println("Tranchen müssen ausgewählt sein!");;
        }
        System.out.println(ik.getText());
        ek.getText();
        fk.getText();

        TextField[] userInputFields = new TextField[10];
        String[] newValues = new String[10];
        int countNonNumeric = 0;
        String nonNumericValue = "";
        for (int i = 0; i < 10; i++) {
            if (IAllExcelRegisterCards.isNumericStr(userInputFields[i].getText()) || userInputFields[i].getText().trim().isEmpty()) {
                newValues[i] = userInputFields[i].getText();
                // Setze das Fehlerlabel auf leer, da keine Fehlermeldung vorliegt
                //errorLabels[i].setText("");
                //valid=true;
            } else {

                countNonNumeric++;
                nonNumericValue = userInputFields[i].getText();
                //errorLabels[i].setText("Achtung: Die Werte müssen numerisch sein. Fehler bei " + nonNumericValue);
            }
        }
        if (countNonNumeric > 0) {
            //resultLabel.setText("Ein oder mehrere Werte sind ungültig.");
            //valid=false;
        } else {
            //updateRangeOfCells(newValues);
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
//        tranche.setValue("1");
//        if(tranche.getValue().equals("1")) {
//            Label lt0 = new Label("Tranche 1");
//            lt0.setLayoutX(300);
//            lt0.setLayoutY(128+30);
//            //lt[i].setId("label"+i);
//
//            TextField tt0 = new TextField();
//            tt0.setLayoutX(400.0);
//            tt0.setLayoutY(120+30);
//            tt0.setId("text0");
//
//            pane.getChildren().add(lt0);
//            pane.getChildren().add(tt0);
//        }


    }


}