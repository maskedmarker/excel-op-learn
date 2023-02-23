package org.example.learn.excel.poi.cell;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Date;

public class CellBaseTest {

    private static final String FILE_NAME_PREFIX = "cell-";

    @Test
    public void test1() throws Exception {
        String workbookName = FILE_NAME_PREFIX + "base.xlsx";

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet spreadsheet = workbook.createSheet("cell types");

        XSSFRow row = spreadsheet.createRow((short) 2);
        row.createCell(0).setCellValue("Type of Cell");
        row.createCell(1).setCellValue("cell value");

        row = spreadsheet.createRow((short) 3);
        row.createCell(0).setCellValue("set cell type BLANK");
        row.createCell(1);

        row = spreadsheet.createRow((short) 4);
        row.createCell(0).setCellValue("set cell type BOOLEAN");
        row.createCell(1).setCellValue(true);

        row = spreadsheet.createRow((short) 5);
        row.createCell(0).setCellValue("set cell type date");
        row.createCell(1).setCellValue(new Date());

        row = spreadsheet.createRow((short) 6);
        row.createCell(0).setCellValue("set cell type numeric");
        row.createCell(1).setCellValue(20);

        row = spreadsheet.createRow((short) 7);
        row.createCell(0).setCellValue("set cell type string");
        row.createCell(1).setCellValue("A String");

        try (OutputStream fileOut = new FileOutputStream(workbookName)) {
            workbook.write(fileOut);
        }

        System.out.printf("%s written successfully%n", workbookName);
    }



}
