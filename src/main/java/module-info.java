module com.example.calcoloeffemeridi {
    requires javafx.controls;
    requires javafx.fxml;


    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;
    requires org.apache.commons.compress;
    requires org.seleniumhq.selenium.chrome_driver;
    requires org.seleniumhq.selenium.support;
    requires dev.failsafe.core;

    opens com.example.calcoloeffemeridi to javafx.fxml;
    exports com.example.calcoloeffemeridi;

}