package com.example.calcoloeffemeridi;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import static org.apache.poi.ss.usermodel.BuiltinFormats.getBuiltinFormat;


public class HelloController {


    @FXML
    private Button btnEsegui;

    @FXML
    private ComboBox<String> cmbProvincia;

    @FXML
    public void initialize() {

        cmbProvincia.getItems().addAll("Agrigento", "Alessandria", "Ancona", "Aosta", "Arezzo", "Ascoli Piceno", "Asti",
                "Avellino", "Bari", "Belluno", "Benevento", "Bergamo", "Biella", "Bologna", "Bolzano", "Brescia", "Brindisi",
                "Cagliari", "Caltanissetta", "Campobasso", "Caserta", "Catania", "Catanzaro", "Chieti", "Como", "Cosenza",
                "Cremona", "Cuneo", "Enna", "Ferrara", "Firenze", "Foggia", "Forli`", "Frosinone", "Genova", "Gorizia", "Grosseto",
                "Imperia", "Isernia", "La Spezia", "L\'aquila", "Latina", "Lecce", "Livorno", "Lucca", "Macerata", "Mantova", "Massa",
                "Matera", "Messina", "Milano", "Modena", "Napoli", "Novara", "Nuoro", "Oristano", "Padova", "Palermo", "Parma", "Pavia",
                "Perugia", "Pesaro", "Pescara", "Piacenza", "Pisa", "Pistoia", "Pordenone", "Potenza", "Prato", "Ragusa", "Ravenna",
                "Reggio Calabria", "Reggio Emilia", "Rieti", "Rimini", "Roma", "Rovigo", "Salerno", "Sassari", "Savona", "Siena",
                "Siracusa", "Sondrio", "Taranto", "Teramo", "Terni", "Torino", "Trapani", "Trento", "Treviso", "Trieste", "Udine",
                "Varese", "Venezia", "Verbania", "Vercelli", "Verona", "Vicenza", "Viterbo");
        
    }



    @FXML
    private TextField txtAnno;

    @FXML
    void onClickEsegui(ActionEvent event) {

        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle("Seleziona la posizione di salvataggio");
        FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Cartella di lavoro di Excel (*.xlsx)", "*.xlsx");
        fileChooser.getExtensionFilters().add(extFilter);
        File file = fileChooser.showSaveDialog(btnEsegui.getScene().getWindow());

        // Ora puoi utilizzare il percorso del file per scrivere i dati o fare altre operazioni
        if (file != null) {
            System.out.println("Percorso di salvataggio selezionato: " + file.getAbsolutePath());
            esporta(file.getAbsolutePath(), txtAnno.getText(), cmbProvincia.getValue());
        }
    }
    private void esporta(String path, String anno, String citta) {
        String[] mesi = {"Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"};


        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Dati-" + citta + "-" + anno);

        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Data");
        headerRow.createCell(1).setCellValue("Sorge");
        headerRow.createCell(2).setCellValue("Tramonta");
        headerRow.createCell(3).setCellValue("Delta");


        System.setProperty("webdriver.chrome.driver",
                "Y:\\lombardini\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");

        WebDriver driver = new ChromeDriver();
        try {
            driver.get("http://www.marcomenichelli.it/sole.asp");

            Select city = new Select(driver.findElement(By.name("MyCity")));
            city.selectByVisibleText(citta);

            WebElement year = driver.findElement(By.name("MyAnno"));
            year.sendKeys(Keys.CONTROL + "a");
            year.sendKeys(Keys.BACK_SPACE);
            year.sendKeys(anno);

            for (int t = 0; t < mesi.length; t++) {
                String mese = mesi[t];



//                Select city = new Select(driver.findElement(By.name("MyCity")));
//                city.selectByVisibleText(citta);

                Select month = new Select(driver.findElement(By.name("MyMese")));
                month.selectByVisibleText(mese);



                Thread.sleep(2000);

                java.util.List<WebElement> colonnaSorge =  driver.findElements(By.xpath("/html/body/center/center/table[5]/tbody/tr/td[2]"));
                java.util.List<WebElement> colonnaTramonta =  driver.findElements(By.xpath("/html/body/center/center/table[5]/tbody/tr/td[6]"));

                for (int i = 1; i < colonnaSorge.size(); i++) {
                    String sorge = colonnaSorge.get(i).getText();
                    String tramonta = colonnaTramonta.get(i).getText();
                    String formattedSorge = formatString(sorge);
                    String formattedTramonta = formatString(tramonta);

                    int numRiga = sheet.getLastRowNum()+1;
                    Row row = sheet.createRow(numRiga);
                    Cell data = row.createCell(0);
                    int m = t+1;
                    if (i<10) {
                        if(m<10) {
                            data.setCellValue("0"+ i +"/0" + m);
                        }
                        else {
                            data.setCellValue("0"+ i +"/" + m);
                        }

                    }
                    else {
                        if (m<10) {
                            data.setCellValue(""+ i +"/0" + m);
                        }
                        else {
                            data.setCellValue(""+ i +"/" + m);
                        }

                    }

                    row.createCell(1).setCellValue(formattedSorge);
                    row.createCell(2).setCellValue(formattedTramonta);
                    Cell formula = row.createCell(3);
                    formula.setCellFormula("1-C" + (++numRiga) + "+B" + numRiga);

                    CellStyle timeCellStyle = workbook.createCellStyle();
                    CellStyle dateCellStyle = workbook.createCellStyle();
                    timeCellStyle.setDataFormat((short)(getBuiltinFormat("h:mm")));
                    dateCellStyle.setDataFormat((short)(getBuiltinFormat("d-mmm")));

                    formula.setCellStyle(timeCellStyle);
                    data.setCellStyle(dateCellStyle);
                }


                workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
            }

            try {
//                FileOutputStream fileOutputStream = new FileOutputStream("Effemeridi-" + citta + "-" + anno +".xlsx");
                FileOutputStream fileOutputStream = new FileOutputStream(path);
                workbook.write(fileOutputStream);
            }catch (IOException e) {
                e.printStackTrace();
            } finally {
                try{
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }



        } catch (InterruptedException e) {
            e.printStackTrace();
        } finally {
            driver.quit();
        }




    }

    private static String formatString(String input) {
        String numeriPart = input.replaceAll("\\D", "");
        if (numeriPart.length() >= 4) {
            String ore = numeriPart.substring(0, 2);
            String minuti = numeriPart.substring(2, 4);
            return ore + ":" + minuti;
        } else {
            return input;
        }
    }
}









