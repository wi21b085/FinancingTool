package com.example.financingtool;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.stage.Stage;

public class Wire extends Application {

    @Override
    public void start(Stage stage) throws Exception {
        FXMLLoader fxmlLoader = new FXMLLoader(Wire.class.getResource("wire.fxml"));
        Scene scene = new Scene(fxmlLoader.load(), 700, 360);
        stage.setTitle("Wirtschaftlichkeitsrechnung");
        stage.setScene(scene);
        stage.show();
    }

    public static void main(String[] args) {
        launch();
    }
}
