module com.example.financingtool {
    requires javafx.controls;
    requires javafx.fxml;

    requires org.controlsfx.controls;
    requires com.dlsc.formsfx;
    requires org.kordamp.bootstrapfx.core;
    requires org.apache.poi.ooxml;
    requires kernel;
    requires org.apache.pdfbox;
    requires org.apache.poi.poi;
    requires com.aspose.words;
    requires java.desktop;


    opens com.example.financingtool to javafx.fxml;
    exports com.example.financingtool;

}