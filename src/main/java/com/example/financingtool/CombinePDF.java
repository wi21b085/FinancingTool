package com.example.financingtool;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;

import java.io.IOException;

public class CombinePDF {
    static String file1 = "src/main/resources/com/example/financingtool/A.pdf";
    static String file2 = "src/main/resources/com/example/financingtool/B.pdf";
    static String outputFile ="src/main/resources/com/example/financingtool/C.pdf";

    public static void combinePdf() {
        //zweiPDF kombinieren
        try {
            // Laden der ersten PDF-Datei
            PDDocument pdfDocument1 = PDDocument.load(new java.io.File(file1));

            // Laden der zweiten PDF-Datei
            PDDocument pdfDocument2 = PDDocument.load(new java.io.File(file2));

            // Kopieren aller Seiten von der ersten PDF-Datei zur Ausgabedatei
            for (int i = 0; i < pdfDocument1.getNumberOfPages(); i++) {
                PDPage page = pdfDocument1.getPage(i);
                pdfDocument2.addPage(page);
            }

            // Speichern des Ergebnisses
            pdfDocument2.save(outputFile);
            System.out.println("Erfolgreiche Kombination der pdf's");

            // Schließen der geöffneten Dokumente
            pdfDocument1.close();
            pdfDocument2.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args){
        combinePdf();
    }
}
