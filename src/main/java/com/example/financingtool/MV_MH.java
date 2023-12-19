package com.example.financingtool;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.scene.text.Text;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class MV_MH extends Application {

    @Override
    public void start(Stage stage) throws Exception {
        FXMLLoader fxmlLoader = new FXMLLoader(MV_MH.class.getResource("MV_MH.fxml"));
        Scene scene = new Scene(fxmlLoader.load(), 700, 500);
        stage.setTitle("Test");
        stage.setScene(scene);
        stage.show();
//        FXMLLoader fxmlLoader = new FXMLLoader(HelloApplication.class.getResource("MV_MH.fxml"));
//        ScrollPane scrollPane = fxmlLoader.load();
//        VBox root = (VBox) scrollPane.getContent();
//
//        Label iv = new Label("Investitionskosten");
//        TextField iv_input = new TextField();
//        Label ek = new Label("Eigenkapital");
//        TextField ek_input = new TextField();
//
//        Button addButton = new Button("Hinzufügen");
//
//        HBox inputBox = new HBox(iv, iv_input, ek, ek_input, addButton);
//        //Scene scene = new Scene(root, 600, 400);
//
//        Scene scene = new Scene(scrollPane, 1280, 720);
//        stage.setTitle("Mittelverwendung & Mittelherkunft");
//
//        // Event Handler für den Hinzufügen-Button
//        addButton.setOnAction(event -> {
//
//            // Leere die Eingabefelder nach dem Hinzufügen
//            iv_input.clear();
//            ek_input.clear();
//        });
//        root.getChildren().addAll(inputBox);
//
//        // Setze die Szene und zeige die Bühne
//        stage.setScene(scene);
//        stage.show();


    }

    public static void main(String[] args) {
        launch();
    }
}
