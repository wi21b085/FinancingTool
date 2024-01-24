package com.example.financingtool;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.IOException;

public class MainApplication extends Application {

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) throws IOException {
        // Lade das FXML-Layout
        FXMLLoader loader = new FXMLLoader(getClass().getResource("MainWindowController.fxml"));
        Parent root = loader.load();

        // Holen Sie den Controller, um auf das TabPane zuzugreifen
        MainWindowController controller = loader.getController();

       /* // Erstellen Sie den Tab mit dem Inhalt des Stammdaten-FXMLs
        Tab stammblattTab = new Tab("Stammdaten");
        stammblattTab.setContent(FXMLLoader.load(getClass().getResource("stammblatt.fxml")));
*/
        // Füge den Tab zum TabPane hinzu


        // Setze die Szene
        Scene scene = new Scene(root, 850, 700);
        primaryStage.setTitle("Financing Tool");
        primaryStage.setScene(scene);
        primaryStage.show();
    }
}