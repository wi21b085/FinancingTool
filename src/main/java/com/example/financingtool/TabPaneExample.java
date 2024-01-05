package com.example.financingtool;

import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.stage.Stage;

public class TabPaneExample extends Application {

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("TabPane Beispiel");

        TabPane tabPane = new TabPane();

        // Erstelle Tab 1
        Tab tab1 = new Tab();
        tab1.setText("Tab 1");
        // Füge die Inhalte für Tab 1 hinzu
        // Beispiel: tab1.setContent(new YourContentForTab1());

        // Erstelle Tab 2
        Tab tab2 = new Tab();
        tab2.setText("Tab 2");
        // Füge die Inhalte für Tab 2 hinzu
        // Beispiel: tab2.setContent(new YourContentForTab2());

        // Füge die Tabs zur TabPane hinzu
        tabPane.getTabs().addAll(tab1, tab2);

        Scene scene = new Scene(tabPane, 400, 300);
        primaryStage.setScene(scene);

        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}

