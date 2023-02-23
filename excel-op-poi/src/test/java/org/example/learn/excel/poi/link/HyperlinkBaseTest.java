package org.example.learn.excel.poi.link;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;

public class HyperlinkBaseTest {

    private static final String FILE_NAME_PREFIX = "hyperlink-";

    @Test
    public void test() throws Exception {
        String workbookName = FILE_NAME_PREFIX + "base.xls";

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet spreadsheet = workbook.createSheet("Hyperlinks");
        CreationHelper createHelper = workbook.getCreationHelper();

        XSSFCell cell;
        XSSFCellStyle hlinkstyle = workbook.createCellStyle();
        XSSFFont hlinkfont = workbook.createFont();
        hlinkfont.setUnderline(XSSFFont.U_SINGLE);
        hlinkfont.setColor(IndexedColors.BLUE.index);
        hlinkstyle.setFont(hlinkfont);

        //URL Link
        cell = spreadsheet.createRow(1).createCell((short) 1);
        cell.setCellValue("URL Link");
        XSSFHyperlink link = (XSSFHyperlink)createHelper.createHyperlink(HyperlinkType.URL);
        link.setAddress("http://www.tutorialspoint.com/");
        cell.setHyperlink(link);
        cell.setCellStyle(hlinkstyle);

        //Hyperlink to a file in the current directory
/*        cell = spreadsheet.createRow(2).createCell((short) 1);
        cell.setCellValue("File Link");
        link = (XSSFHyperlink)createHelper.createHyperlink(HyperlinkType.FILE);
        link.setAddress("cellstyle.xlsx");
        cell.setHyperlink(link);
        cell.setCellStyle(hlinkstyle);

        //e-mail link
        cell = spreadsheet.createRow(3).createCell((short) 1);
        cell.setCellValue("Email Link");
        link = (XSSFHyperlink)createHelper.createHyperlink(HyperlinkType.EMAIL);
        link.setAddress("mailto:contact@tutorialspoint.com?" + "subject = Hyperlink");
        cell.setHyperlink(link);
        cell.setCellStyle(hlinkstyle);*/

        FileOutputStream out = new FileOutputStream(workbookName);
        workbook.write(out);
        out.close();
        System.out.println("hyperlink.xlsx written successfully");
    }
}
