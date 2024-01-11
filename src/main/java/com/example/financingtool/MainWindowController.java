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



    @FXML
    private Tab stammblatt;

    @FXML
    private Tab basisinformation;

    @FXML
    private Tab gik;

    @FXML
    private Tab mvMh;

    @FXML
    private Tab wire;

    @FXML
    private Tab widmung;

    @FXML
    private Tab logoMaker;

    @FXML
    private Tab converter;



    @Override
    public void initialize(URL url, ResourceBundle resourceBundle) {
        System.out.println("Initializing MainWindowController...");

       /* stammblatt.setClosable(false);
        basisinformation.setClosable(false);
        gik.setClosable(false);
        mvMh.setClosable(false);
        wire.setClosable(false);
        widmung.setClosable(false);
        logoMaker.setClosable(false);
        converter.setClosable(false);
        // Print statements or debugging code...
        gikController.setMV_MH_Controller(mvMhController);*/
        System.out.println("Initialization complete.");


    }

    public void convert(ActionEvent actionEvent) {
        ExcelToWordConverter.exportExcelToWord();
    }


    // Rest des Codes für den MainWindowController...
}