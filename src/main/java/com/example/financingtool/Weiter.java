package com.example.financingtool;

import javafx.application.Application;
import javafx.scene.control.Button;
import javafx.stage.Stage;

public class Weiter {


    public static void weiter (Button weiterButton, Class<?> applicationClass){
        Stage currentStage = (Stage) weiterButton.getScene().getWindow();
        currentStage.close();

        // Starte eine neue Anwendung
        startNewApplication(applicationClass);
    }

     public static void startNewApplication(Class<?> applicationClass) {

        //BasisinformationApplication basisinformationApplication = new BasisinformationApplication();
        try {
            Application application = (Application) applicationClass.getDeclaredConstructor().newInstance();
            // Starte die Anwendung
            application.start(new Stage());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
