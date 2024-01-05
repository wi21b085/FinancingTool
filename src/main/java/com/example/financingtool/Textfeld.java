package com.example.financingtool;

import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;

import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.stage.Stage;

import java.io.IOException;

public class Textfeld extends Application {
    @FXML
    private TextField textField;

    //Maria M
    @FXML
    private Button weiterButton;

    @FXML
    protected void senden(ActionEvent event) throws IOException {
        String text = textField.getText();
        System.out.println("Gesendeter Text: " + text);
        ExcelToWordConverter.addTextToFirstPage(text);
    }

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage stage) throws Exception {
        Parent root = FXMLLoader.load(getClass().getResource("Textfeld.fxml"));
        stage.setTitle("JavaFX Text to Word");
        Scene scene = new Scene(root, 300, 200);

        stage.setScene(scene);
        stage.show();
    }

    public void weiter(ActionEvent actionEvent) {
        Weiter.weiter(weiterButton, StammblattApplication.class);
    }
}

