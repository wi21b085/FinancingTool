package com.example.financingtool;
import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/*
import java.io.*;

public class BasisinformationApplication extends Application {
    //wozu brauch ich das?
    private Label resultLabel;
    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(BasisinformationApplication.class.getResource("basisinformation.fxml"));
        ScrollPane scrollPane = fxmlLoader.load();
        VBox root = (VBox) scrollPane.getContent();

        Scene scene = new Scene(scrollPane, 1280, 720);
        stage.setTitle("Hello!");

        //was mache ich damit
        resultLabel = new Label("Aktueller Wert: ");

        // Setze die Szene und zeige die Bühne
        stage.setScene(scene);
        stage.show();
    }
}*/
