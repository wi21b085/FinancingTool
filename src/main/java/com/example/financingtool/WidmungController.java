package com.example.financingtool;

import javafx.fxml.FXML;
import javafx.scene.control.ChoiceBox;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.text.Text;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;


public class WidmungController implements IAllExcelRegisterCards {

    @FXML
    private ChoiceBox<String> fw = new ChoiceBox<>();
    @FXML
    private ChoiceBox<String> bk = new ChoiceBox<>();
    @FXML
    private ChoiceBox<String> bw = new ChoiceBox<>();
    @FXML
    private ChoiceBox<String> bb1 = new ChoiceBox<>();
    @FXML
    private ChoiceBox<String> bb2 = new ChoiceBox<>();
    @FXML
    private TextField bs;

    @FXML
    private Text adresse;
    private String adCell;
    @FXML
    private Label resultLabel;

    static ExecutiveSummary executiveSummary = new ExecutiveSummary();


    public static void setExecutiveSummary(ExecutiveSummary executiveSummary) {
        BasisinformationController.executiveSummary=executiveSummary;
    }

    @FXML
    protected void continueClick() {
        boolean fwb = check();

        System.out.println(fwb);
        if(fwb)
            executePy(adCell);
    }

    private boolean check() {
        try {
            if (fw.getValue() != null && bk.getValue() != null) {
                String val = fw.getValue() + " " + bk.getValue() + " = Bauklasse " + rome(bk.getValue());
                if (!bs.getText().isEmpty())
                    val += " beschränkt auf " + bs.getText();

                System.out.println(val);
                setCell(val, 1, 14);
              //  executiveSummary.setDatenausWidmung(val);
                resultLabel.setText("");
            } else {
                resultLabel.setText("Hinweis: Flächenwidmung oder Bauklasse nicht ausgewählt");
                return false;
            }

            if (bw.getValue() != null) {
                String bauweise = switch (bw.getValue()) {
                    case "o" -> "o = offene Bauweise";
                    case "gk" -> "gk = gekuppelte Bauweise";
                    case "g" -> "g = geschlossene Bauweise";
                    default -> null;
                };
                System.out.println(bauweise);
                setCell(bauweise, 2, 14);
                resultLabel.setText("Daten aktualisiert");
            } else {
                resultLabel.setText("Hinweis: Bauweise nicht ausgewählt");
                return false;
            }

            if (bb1.getValue() != null) {
                String bb = bbVal(bb1);
                setCell(bb, 3, 14);
            } else {
                setCell("\n", 3, 14);
            }

            if (bb2.getValue() != null) {
                String bb = bbVal(bb2);

                if(bb1.getValue() != null && !bb1.getValue().isEmpty()){
                    if(bb1.getValue().contains(bb2.getValue()))
                        setCell("\n", 4, 14);
                    else
                        setCell(bb, 4, 14);
                } else {
                    setCell(bb, 3, 14);
                    setCell("\n", 4, 14);
                }
            } else {
                setCell("\n", 4, 14);
            }
            return true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return false;
    }

    private String bbVal(ChoiceBox<String> bb2) {
        String bb = switch (bb2.getValue()) {
            case "Ekz" -> "Ekz = für Einkaufszentren bestimmt";
            case "ÖZ" -> "ÖZ = Grundflächen für öffentliche Zwecke (Enteignung möglich)";
            case "BB" -> "BB = Besondere Bestimmungen";
            case "G" -> "G = Gärtnerische Ausgestaltung";
            default -> "\n";
        };
        System.out.println(bb);
        return bb;
    }

    private String rome(String value) {
        String res = switch (value) {
            case "I" -> "1";
            case "II" -> "2";
            case "III" -> "3";
            case "IV" -> "4";
            case "V" -> "5";
            case "VI" -> "6";
            default -> null;
        };
        return res;
    }

    public void executePy(String name) {
        String s = null;
        try {
            //Runtime.getRuntime().exec("src\\main\\resources\\com\ \example\\financingtool\\script.bat");
            //String[] cmd = { "python", "src\\main\\resources\\com\\example\\financingtool\\widmung.py", name};
            //Process p = Runtime.getRuntime().exec(cmd);
            Process p = Runtime.getRuntime().exec("cmd /C start src\\main\\resources\\com\\example\\financingtool\\script.bat widmung.py \"" + name + "\"");

            BufferedReader stdInput = new BufferedReader(new
                    InputStreamReader(p.getInputStream()));

            BufferedReader stdError = new BufferedReader(new
                    InputStreamReader(p.getErrorStream()));

            // read the output from the command
            System.out.println("Here is the standard output of the command:\n");
            while ((s = stdInput.readLine()) != null) {
                System.out.println(s);
            }

            // read any errors from the attempted command
            System.out.println("Here is the standard error of the command (if any):\n");
            while ((s = stdError.readLine()) != null) {
                System.out.println(s);
            }

            //System.exit(0);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void initialize() {
        //getAddress();
        EventBus.getInstance().subscribe("updateAddress", this::updateAddress);
        bs.setPromptText("40% und/oder 45m");
    }

    private void updateAddress(Object newValue) {
        this.adresse.setText(newValue.toString());
    }

    public void getAddress() {
        try {
            String excelFilePath = "src/main/resources/com/example/financingtool/SEPJ-Rechnungen.xlsx";
            String sheetName = "Basisinformation";
            int rowIdx = 2;
            int colIdx = 8;

            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheet(sheetName);

            Row row = sheet.getRow(rowIdx);
            Cell cell = row.getCell(colIdx);
            adCell = cell.getStringCellValue();
            //System.out.println(adCell);
            adresse.setText(adCell);
            fileInputStream.close();

            workbook.close();


        } catch (NumberFormatException | IOException e) {
            e.printStackTrace();
            //resultLabel.setText("Fehler bei der Aktualisierung.");
        }
    }

    public void setCell(String tf, int row, int coll) {
        String sheet = "Basisinformation";
        String[] field = new String[1];
        field[0] = tf;
        Update.updateRangeOfCellsString(field, row, row, coll, new Label(), sheet);
    }
}
