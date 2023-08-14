package se_nata;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Iterator;


public class SplitExcelFile {

    private final String fileName;
    private HashSet<String> uniqueValues = new HashSet<String>();
    private XSSFWorkbook workbook;

    public SplitExcelFile(String fileName) throws InvalidFormatException {

        this.fileName = fileName;
        try {
            OPCPackage pkg = OPCPackage.open(new File(fileName));
            workbook = new XSSFWorkbook(pkg);
            getSetOfValues();
            writeWorkBooks();
            pkg.close();

        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private void getSetOfValues() {

        XSSFSheet sheet = this.workbook.getSheetAt(0);
        Iterator<Row> row = sheet.rowIterator();
        while (row.hasNext()) {
            Row r = (Row) row.next();
            Iterator<Cell> cell = r.cellIterator();
            while (cell.hasNext()) {
                Cell c = cell.next();
                switch (c.getCellType()) {
                    case STRING:
                        c.getRichStringCellValue();
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(c)) {
                            c.getDateCellValue();
                        } else {
                            c.getNumericCellValue();
                        }
                        break;
                    case BLANK:
                        c.getStringCellValue();

                    case BOOLEAN:
                        c.getBooleanCellValue();
                        break;

                    case FORMULA:
                        c.getCachedFormulaResultType();
                        break;
                    default:

                        System.out.println("Could not determine cell type");

                }
                if (c.getCellType().equals(CellType.STRING) & c.getColumnIndex() == 2 & r.getRowNum() > 8) {
                    uniqueValues.add(c.getRichStringCellValue().getString().trim());
                }
            }
        }

    }

    private void writeWorkBooks() throws FileNotFoundException {

        XSSFWorkbook book;

        for (String s : uniqueValues) {
            book = new XSSFWorkbook();
            XSSFRow nr;
            XSSFRow oldrow;

            int count = 0;
            for (int i = 0; i < this.workbook.getNumberOfSheets(); i++) {
                XSSFSheet newsheet = this.workbook.getSheetAt(i);
                XSSFSheet shnewbook = book.createSheet(newsheet.getSheetName());
                Iterator<Row> newrow = newsheet.rowIterator();
                while (newrow.hasNext()) {
                    Row newr = (Row) newrow.next();
                    oldrow = newsheet.getRow(newr.getRowNum());

                    if (newr.getCell(2) != null) {
                        if (newr.getRowNum() >= 8 && !newr.getCell(2).getStringCellValue().trim().equals(s)) {
                            continue;
                        }
                    }
                    nr = shnewbook.createRow(count);
                    count++;
                    Iterator<Cell> newcell = newr.cellIterator();
                    while (newcell.hasNext()) {
                        Cell newc = newcell.next();


                        if (newc != null) {

                            XSSFCell xssfCell = nr.createCell(newc.getColumnIndex());
                            XSSFCell hssfCell = oldrow.getCell(newc.getColumnIndex());

                            CellStyle nstyle = book.createCellStyle();
                            nstyle.cloneStyleFrom(newc.getCellStyle());
                            xssfCell.setCellStyle(nstyle);

                            if (hssfCell != null) {
                                switch (hssfCell.getCellType()) {
                                    case BOOLEAN:
                                        xssfCell.setCellValue(hssfCell.getBooleanCellValue());
                                        break;
                                    case NUMERIC:

                                        if (DateUtil.isCellDateFormatted(hssfCell)) {
                                            xssfCell.setCellValue(hssfCell.getDateCellValue());
                                        } else {

                                            xssfCell.setCellValue(hssfCell.getNumericCellValue());
                                        }
                                        break;
                                    case STRING:
                                        xssfCell.setCellValue(hssfCell.getStringCellValue());
                                        break;
                                    case FORMULA:
                                        xssfCell.setCellFormula(hssfCell.getCellFormula());
                                        break;
                                    case BLANK:
                                        xssfCell.setBlank();
                                        break;
                                    default:
                                        break;


                                }
                            }


                        }
                    }

                }

            }

            String fileNewName = s.toString() + "_new_" + ".xlsx";
            try {
                FileOutputStream out = new FileOutputStream(fileNewName);
                book.write(out);
                out.close();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }

        }

    }
}