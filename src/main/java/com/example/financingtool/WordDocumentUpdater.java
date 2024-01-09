package com.example.financingtool;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class WordDocumentUpdater {
    public static void main(String[] args) {
        try {
            updateWordDocument("src/main/resources/com/example/financingtool/SEPJ-Rechnungen.docx", "Neuer Text für das bestehende Dokument.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    static void updateOnFirstPage(String filePath, String newText) throws IOException{
        File file = new File(filePath);
        FileInputStream fis = new FileInputStream(file);
        XWPFDocument document = new XWPFDocument(fis);

        // Einfügen des neuen Absatzes am Anfang des Dokuments
        XWPFParagraph newParagraph = document.createParagraph();
        XWPFRun run = newParagraph.createRun();
        run.setText(newText);

        // Verschieben des vorhandenen Inhalts nach unten
        for (IBodyElement element : document.getBodyElements()) {
            if (element instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph) element;
                if (!paragraph.isEmpty()) {
                    // Füge den vorhandenen Absatz in das Dokument ein
                    XWPFParagraph clonedParagraph = document.createParagraph();
                    clonedParagraph.getCTP().set(paragraph.getCTP().copy());
                    document.setParagraph(clonedParagraph, document.getParagraphs().size() - 1);
                }
            }
        }

        fis.close();

        // Speichern Sie das aktualisierte Dokument
        FileOutputStream fos = new FileOutputStream(file);
        document.write(fos);
        fos.close();
    }

    static void updateWordDocument(String filePath, String newText) throws IOException {
        File file = new File(filePath);
        FileInputStream fis = new FileInputStream(file);
        XWPFDocument document = new XWPFDocument(fis);

        // Hinzufügen des neuen Textes
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(newText);

        fis.close();

        // Speichern Sie das aktualisierte Dokument
        FileOutputStream fos = new FileOutputStream(file);
        document.write(fos);
        fos.close();
    }
}
