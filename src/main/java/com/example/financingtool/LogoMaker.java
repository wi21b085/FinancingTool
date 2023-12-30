package com.example.financingtool;

import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.ScrollPane;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;

import java.io.File;
import java.io.IOException;

public class LogoMaker extends Application {

    @Override
    public void start(Stage stage) throws Exception {
        FXMLLoader fxmlLoader = new FXMLLoader(PdfTest.class.getResource("logo.fxml"));
        ScrollPane scrollPane = fxmlLoader.load();
        VBox root = (VBox) scrollPane.getContent();

        Scene scene = new Scene(scrollPane, 1280, 720);
        stage.setTitle("Hello!");

        // Setze die Szene und zeige die Bühne
        stage.setScene(scene);
        stage.show();
    }
    public static void generatePdf(String imagePath) {
        try {
            // Pfad zur vorhandenen PDF-Datei
            String existingPdfPath = "src/main/resources/com/example/financingtool/empty.pdf";

            // Pfad zur Ausgabedatei
            String outputPdfPath = "src/main/resources/com/example/financingtool/logo.pdf";

            // Lade die vorhandene PDF
            PDDocument document = PDDocument.load(new File(existingPdfPath));

            // Füge eine neue Seite hinzu (optional, wenn du das Bild auf einer bestehenden Seite platzieren möchtest)
            PDPage page = new PDPage();
            document.addPage(page);

            // Lade das Bild
            PDImageXObject image = PDImageXObject.createFromFile(imagePath, document);

            // Füge das Bild auf der Seite hinzu
            PDPageContentStream contentStream = new PDPageContentStream(document, page, PDPageContentStream.AppendMode.APPEND, true);

            //image.getHeight(), image.getWidth()
            contentStream.drawImage(image, 430, 630, 100, 100);
            contentStream.close();

            // Speichere das aktualisierte PDF-Dokument
            document.save(outputPdfPath);
            document.close();

            System.out.println("Bild erfolgreich zu vorhandener PDF hinzugefügt.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }



    public void submit(ActionEvent actionEvent) {
        // Erstelle eine Instanz von FileChooser
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Bild auswählen");

        // Füge eine Filteroption für Bilddateien hinzu (optional)
        FileChooser.ExtensionFilter imageFilter = new FileChooser.ExtensionFilter("Bilder", "*.png", "*.jpg", "*.gif");
        fileChooser.getExtensionFilters().add(imageFilter);

        // Zeige den FileChooser und erhalte das ausgewählte Bild
        File selectedFile = fileChooser.showOpenDialog(null);

        if (selectedFile != null) {
            // Lade das ausgewählte Bild
            Image image = new Image(selectedFile.toURI().toString());

            // Erstelle ein ImageView und setze das Bild
            ImageView imageView = new ImageView(image);

            // Setze das ImageView in einen StackPane (oder einen anderen Container deiner Wahl)
            StackPane stackPane = new StackPane();
            stackPane.getChildren().add(imageView);

            // Erstelle die Szene und füge das StackPane hinzu
            Scene scene = new Scene(stackPane, 600, 400);

            // Setze die Szene für die Bühne (Stage)
            Stage stage = new Stage();
            stage.setScene(scene);

            // Setze den Titel und zeige die Bühne
            stage.setTitle("Hochgeladenes Bild");
            stage.show();

            // Hier kannst du die generatepdf-Methode aufrufen und den Bildpfad übergeben
            generatePdf(selectedFile.getAbsolutePath());
        }
}}
