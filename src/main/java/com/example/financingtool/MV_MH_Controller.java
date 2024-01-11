package com.example.financingtool;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.Node;
import javafx.scene.control.Button;
import javafx.scene.control.ChoiceBox;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.Pane;
import javafx.scene.text.Text;

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
    private Text fk=new Text();

    @FXML
    private TextField btvg;
    @FXML
    private ChoiceBox<String> tranche = new ChoiceBox<>();

    //Maria M
    @FXML
    private Button weiterButton;

    String newvalue="Kein FK";



    @FXML
    protected void continueClick() {
        try {
            if(!tranche.getValue().isEmpty()){
                String sheet = "Mittelverwendung - Mittelherkun";

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
                            setCell(this.ik, 1, 9);
                            setCell(this.ek, 1, 12);
                            setCell(this.btvg, 3, 12);
                            Update.updateRangeOfCells(vals, 2, 2, 12, new Label(), sheet);
                            Update.updateRangeOfCellsString(bez, 2, 2, 11, new Label(), sheet);
                            insertTranche(sheet);
                            break;
                        case 2:
                            setCell(this.ik, 1, 1);
                            setCell(this.ek, 1, 4);
                            setCell(this.btvg, 4, 4);
                            Update.updateRangeOfCells(vals, 2, 3, 4, new Label(), sheet);
                            Update.updateRangeOfCellsString(bez, 2, 3, 3, new Label(), sheet);
                            insertTranche(sheet);
                            break;
                        case 3:
                            setCell(this.ik, 1, 17);
                            setCell(this.ek, 1, 20);
                            setCell(this.btvg, 5, 20);
                            Update.updateRangeOfCells(vals, 2, 4, 20, new Label(), sheet);
                            Update.updateRangeOfCellsString(bez, 2, 4, 19, new Label(), sheet);
                            insertTranche(sheet);
                            break;
                        case 4:
                            setCell(this.ik, 1, 25);
                            setCell(this.ek, 1, 28);
                            setCell(this.btvg, 6, 28);
                            Update.updateRangeOfCells(vals, 2, 5, 28, new Label(), sheet);
                            Update.updateRangeOfCellsString(bez, 2, 5, 27, new Label(), sheet);
                            insertTranche(sheet);
                            break;
                        case 5:
                            setCell(this.ik, 1, 33);
                            setCell(this.ek, 1, 36);
                            setCell(this.btvg, 7, 36);
                            Update.updateRangeOfCells(vals, 2, 6, 36, new Label(), sheet);
                            Update.updateRangeOfCellsString(bez, 2, 6, 35, new Label(), sheet);
                            insertTranche(sheet);
                            break;
                    }
                }
            }
        } catch (Exception e) {
            System.out.println("Tranchen müssen ausgewählt sein!");;
        }

    }

    private void insertTranche(String sheet) {
        Update.updateRangeOfCells(new String[]{tranche.getValue()}, 10, 10, 7, new Label(), sheet);
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

    /*public void updateFKValue(String newValue) {
        System.out.println("In MV_MH " + newValue);
        this.fk.setText(newValue);
    }*/

    public void setFKValue(String fkValue) {
        updateFKValue(fkValue);
    }

    public void initialize() throws Exception {
        EventBus.getInstance().subscribe("updateFK", this::updateFKValue);
    }

    private void updateFKValue(Object newValue) {
        this.fk.setText(newValue.toString());

    }

/*    public void getFK() throws Exception {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Gesamtinvestitionskosten";
            int rowIdx = 9;
            int colIdx = 1;

            // FileInputStream und Workbook hier erstellen
            try (FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
                 Workbook workbook = new XSSFWorkbook(fileInputStream)) {

                Sheet sheet = workbook.getSheet(sheetName);

                Row row = sheet.getRow(rowIdx);
                Cell cell = row.getCell(colIdx);
                String fkCell = Double.toString(cell.getNumericCellValue());
                System.out.println(fkCell);
                fk.setText(fkCell);
            } catch (NumberFormatException | IOException e) {
                e.printStackTrace();
                //resultLabel.setText("Fehler bei der Aktualisierung.");
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }*/



    public void weiterMariaM(ActionEvent actionEvent) {
        //Weiter.weiter(weiterButton,Textfeld.class);
        ExcelToWordConverter.exportExcelToWord();
    }

 public void setCon() {
        GIKController.setMV_MH_Controller(this);
    }
}