package se_nata;


import org.apache.logging.log4j.core.tools.picocli.CommandLine;
import org.apache.logging.log4j.core.util.Assert;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;

import static org.junit.jupiter.api.Assertions.assertEquals;

class SplitExcelFileTest {
    String file = "./test.xlsx";
    HashSet<String> values = new HashSet<String>();
    XSSFWorkbook workbook;

    @BeforeEach
    void setup() throws InvalidFormatException, IOException {
        OPCPackage pkg = OPCPackage.open(new File(file));
        workbook = new XSSFWorkbook(pkg);
        values.add("OAD");
        values.add("AOD");
    }

    @Test
    void test_getSetOfValues() {
        boolean success = false;
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(1);
        XSSFCell cell = row.getCell(0);
        XSSFRow row1 = sheet.getRow(2);
        XSSFCell cell1 = row1.getCell(0);
        assertEquals("OAD", cell1.getStringCellValue());
        XSSFRow row2 = sheet.getRow(3);
        XSSFCell cell2 = row2.getCell(0);
        assertEquals("AOD", cell2.getStringCellValue());
        assertEquals(2, values.size());
    }

    @Test
    void test_writeWorkBooks() {
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFWorkbook book;
        for (String s : values) {

             book = new XSSFWorkbook();
            XSSFSheet sheetnew = book.createSheet();
            int count = 0;
            for (int i = 0; i <= 3; i++) {


                if (sheet.getRow(i).getCell(0).getStringCellValue().equals(s)) {

                    XSSFRow oldrow = sheet.getRow(i);

                    XSSFRow rownew = sheetnew.createRow(count);
                    count++;
                    for (int j = 0; j <= 3; j++) {

                        XSSFCell oldcell = oldrow.getCell(j);

                        XSSFCell cellnew = rownew.createCell(j);

                        cellnew.setCellValue(oldcell.getStringCellValue());

                    }


                }

              
                String fileNewName = s.toString() + "_testnew_" + ".xlsx";
                try {
                    FileOutputStream out = new FileOutputStream(fileNewName);
                    book.write(out);
                    out.close();
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }


            }


        }
        //System.out.println("3 "+ book.getSheetAt(0).getRow(0).getPhysicalNumberOfCells());
       // assert book != null;
       // assertEquals(4, book.getSheetAt(0).getRow(0).getPhysicalNumberOfCells());
      //  assertEquals("ИОД", book.getSheetAt(0).getRow(0).getCell(2).getStringCellValue());
    }
}