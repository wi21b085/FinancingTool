package com.example.financingtool;

import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.layout.AnchorPane;
import javafx.stage.Stage;

import java.io.IOException;
import java.net.URL;
import java.util.ResourceBundle;

public class MainWindowController implements Initializable {

    @FXML
    private AnchorPane mainAnchorPane;

    @Override
    public void initialize(URL location, ResourceBundle resources) {
        // Erstellen Sie eine Instanz der GIKtoExcel-Klasse
        GIKtoExcel gikToExcel = new GIKtoExcel();

        // Rufen Sie die Methode start auf, um das GIKtoExcel-Fenster zu starten
        try {
            gikToExcel.start(new Stage());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Rest des Codes für den MainWindowController...
}