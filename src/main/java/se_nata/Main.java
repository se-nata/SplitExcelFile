package se_nata;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class Main {
    public static void main(String[] args) throws InvalidFormatException {

        new SplitExcelFile("./Таблицы домена RCO — техническое наполнение.xlsx");
    }
}