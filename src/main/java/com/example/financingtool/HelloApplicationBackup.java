package com.example.financingtool;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.IOException;

public class HelloApplicationBackup extends Application {
    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(HelloApplicationBackup.class.getResource("hello-view.fxml"));
        Scene scene = new Scene(fxmlLoader.load(), 1280, 720);
        stage.setTitle("Hello!");
        stage.setScene(scene);
        stage.show();
    }


    public static void main(String[] args) {

    }
}


//import javafx.application.Application;
//import javafx.scene.Scene;
//import javafx.scene.control.Button;
//import javafx.scene.layout.StackPane;
//import javafx.stage.Stage;
//
//public class HelloApplication extends Application {
//
//    public static void main(String[] args) {
//        launch(args);
//    }
//
//    @Override
//    public void start(Stage primaryStage) {
//        primaryStage.setTitle("Schritt 1");
//        Button nextPageButton = new Button("Weiter");
//        StackPane root = new StackPane(nextPageButton);
//        primaryStage.setScene(new Scene(root, 300, 200));
//        primaryStage.show();
//
//        nextPageButton.setOnAction(event -> {
//            primaryStage.close();
//            Stage stage2 = new Stage();
//            stage2.setTitle("Schritt 2");
//            Button nextPageButton2 = new Button("Weiter");
//            StackPane page2Root = new StackPane(nextPageButton2);
//            stage2.setScene(new Scene(page2Root, 300, 200));
//            stage2.show();
//
//            nextPageButton2.setOnAction(event2 -> {
//                Stage stage3 = new Stage();
//                stage3.setTitle("Schritt 3");
//                StackPane page3Root = new StackPane(new Button("Fertig"));
//                stage3.setScene(new Scene(page3Root, 300, 200));
//                stage3.show();
//            });
//        });
//    }
//}