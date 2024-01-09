package com.example.financingtool;

import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.Node;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.layout.AnchorPane;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

import java.io.IOException;
import java.net.URL;
import java.util.ResourceBundle;

public class MainWindowController implements Initializable {

    @FXML
    private AnchorPane mainAnchorPane;

    @FXML
    private VBox dynamicElementsContainer;

    @FXML
    private TabPane tabPane; // Füge diese Zeile hinzu
    @FXML
    private GIKController gikController;
    @FXML
    private MV_MH_Controller mvMhController;



    @Override
    public void initialize(URL url, ResourceBundle resourceBundle) {
        System.out.println("Initializing MainWindowController...");

        // Print statements or debugging code...

        System.out.println("Initialization complete.");
        gikController.setMV_MH_Controller(mvMhController);
    }

    public void convert(ActionEvent actionEvent) {
        ExcelToWordConverter.exportExcelToWord();
    }


    // Rest des Codes für den MainWindowController...
}