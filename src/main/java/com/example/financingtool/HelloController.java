package com.example.financingtool;

import javafx.fxml.FXML;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;

public class HelloController {

    private Stage stage;
    @FXML
    private Label welcomeText;

    @FXML
    private TextField kaufpreis;

    @FXML
    private Label output;
    @FXML
    protected void onHelloButtonClick() {
        //welcomeText.setText("Welcome to JavaFX Application!");
        // https://www.youtube.com/watch?v=Af-hwO19AMY
        FileChooser.ExtensionFilter ex1 = new FileChooser.ExtensionFilter("Grafikdateien", "*.png", "*.jpg", "*.jpeg");
        FileChooser.ExtensionFilter ex2 = new FileChooser.ExtensionFilter("Alle Dateien", "*.*");


        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().addAll(ex1, ex2);
        fileChooser.setTitle("Wähle eine Grafik aus");
        fileChooser.setInitialDirectory(new File("C:\\"));
        File selectedFile = fileChooser.showOpenDialog(stage);
        if (selectedFile != null) {
            System.out.println(selectedFile.getPath());
            welcomeText.setText(selectedFile.getName());
        }
    }

    @FXML
    protected void submit(){
        output.setText("€ " + kaufpreis.getText());
    }
}