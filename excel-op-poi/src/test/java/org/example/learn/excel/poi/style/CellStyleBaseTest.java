package org.example.learn.excel.poi.style;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;

public class CellStyleBaseTest {

    private static final String FILE_NAME_PREFIX = "cell-style-";

    @Test
    public void test() throws Exception {
        String workbookName = FILE_NAME_PREFIX + "base.xls";

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet spreadsheet = workbook.createSheet("cellstyle");

        XSSFRow row = spreadsheet.createRow((short) 1);
        row.setHeight((short) 800);
        XSSFCell cell = row.createCell((short) 1);
        cell.setCellValue("test of merging");

        //MEARGING CELLS
        //this statement for merging cells

        spreadsheet.addMergedRegion(
                new CellRangeAddress(
                        1, //first row (0-based)
                        1, //last row (0-based)
                        1, //first column (0-based)
                        4 //last column (0-based)
                )
        );

        //CELL Alignment
        row = spreadsheet.createRow(5);
        cell = row.createCell(0);
        row.setHeight((short) 800);

        // Top Left alignment
        XSSFCellStyle style1 = workbook.createCellStyle();
        spreadsheet.setColumnWidth(0, 8000);
        style1.setAlignment(HorizontalAlignment.LEFT);
        style1.setVerticalAlignment(VerticalAlignment.TOP);
        cell.setCellValue("Top Left");
        cell.setCellStyle(style1);
        row = spreadsheet.createRow(6);
        cell = row.createCell(1);
        row.setHeight((short) 800);

        // Center Align Cell Contents
        XSSFCellStyle style2 = workbook.createCellStyle();
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellValue("Center Aligned");
        cell.setCellStyle(style2);
        row = spreadsheet.createRow(7);
        cell = row.createCell(2);
        row.setHeight((short) 800);

        // Bottom Right alignment
        XSSFCellStyle style3 = workbook.createCellStyle();
        style3.setAlignment(HorizontalAlignment.RIGHT);
        style3.setVerticalAlignment(VerticalAlignment.BOTTOM);
        cell.setCellValue("Bottom Right");
        cell.setCellStyle(style3);
        row = spreadsheet.createRow(8);
        cell = row.createCell(3);

        // Justified Alignment
        XSSFCellStyle style4 = workbook.createCellStyle();
        style4.setAlignment(HorizontalAlignment.JUSTIFY);
        style4.setVerticalAlignment(VerticalAlignment.JUSTIFY);
        cell.setCellValue("Contents are Justified in Alignment");
        cell.setCellStyle(style4);

        //CELL BORDER
        row = spreadsheet.createRow((short) 10);
        row.setHeight((short) 800);
        cell = row.createCell((short) 1);
        cell.setCellValue("BORDER");

        XSSFCellStyle style5 = workbook.createCellStyle();
        style5.setBorderBottom(BorderStyle.THICK);
        style5.setBottomBorderColor(IndexedColors.BLUE.getIndex());
        style5.setBorderLeft(BorderStyle.DOUBLE);
        style5.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        style5.setBorderRight(BorderStyle.HAIR);
        style5.setRightBorderColor(IndexedColors.RED.getIndex());
        style5.setBorderTop(BorderStyle.DOTTED);
        style5.setTopBorderColor(IndexedColors.CORAL.getIndex());
        cell.setCellStyle(style5);

        //Fill Colors
        //background color
        row = spreadsheet.createRow((short) 10 );
        cell = row.createCell((short) 1);

        XSSFCellStyle style6 = workbook.createCellStyle();
        style6.setFillBackgroundColor(IndexedColors.LIME.index);
        style6.setFillPattern(FillPatternType.LESS_DOTS);
        style6.setAlignment(HorizontalAlignment.FILL);
        spreadsheet.setColumnWidth(1,8000);
        cell.setCellValue("FILL BACKGROUNG/FILL PATTERN");
        cell.setCellStyle(style6);

        //Foreground color
        row = spreadsheet.createRow((short) 12);
        cell = row.createCell((short) 1);

        XSSFCellStyle style7 = workbook.createCellStyle();
        style7.setFillForegroundColor(IndexedColors.BLUE.index);
        style7.setFillPattern( FillPatternType.LESS_DOTS);
        style7.setAlignment(HorizontalAlignment.FILL);
        cell.setCellValue("FILL FOREGROUND/FILL PATTERN");
        cell.setCellStyle(style7);

        try (FileOutputStream out = new FileOutputStream(workbookName)) {
            workbook.write(out);
        }

        System.out.println("cellstyle.xlsx written successfully");
    }
}
