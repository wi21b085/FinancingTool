package com.example.financingtool;

import com.itextpdf.kernel.color.Lab;
import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javafx.scene.paint.Color;

public class GIKtoExcel extends Application  implements IAllExcelRegisterCards{

    public Label resultLabel=new Label("Aktueller Wert: ");

    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(GIKtoExcel.class.getResource("gik.fxml"));
        ScrollPane scrollPane = fxmlLoader.load();
        VBox root = (VBox) scrollPane.getContent();

        Scene scene = new Scene(scrollPane, 1280, 720);
        stage.setTitle("GIKtoExcel");
        // Setze die Szene und zeige die Bühne
        stage.setScene(scene);
        stage.show();
    }
    public static void main(String[] args) {
        launch(args);

    }

}
