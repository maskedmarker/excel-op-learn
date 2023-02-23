package org.example.learn.excel.poi.workbook;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

import java.io.File;

public class ReadWorkbookTest {

    @Test
    public void test6() throws Exception {
        String workbookName = WorkbookBaseTest.WORKBOOKNAME;

        // When opening a workbook, either a .xls HSSFWorkbook, or a .xlsx XSSFWorkbook,
        // the Workbook can be loaded from either a File or an InputStream.
        // Using a File object allows for lower memory consumption,
        // while an InputStream requires more memory as it has to buffer the whole file.
        // WorkbookFactory.create创建的Workbook只能是read-only
        Workbook wb = WorkbookFactory.create(new File(workbookName));
        Sheet sheetAt0 = wb.getSheetAt(0);
        String sheetName = sheetAt0.getSheetName();
        System.out.printf("%s.sheetAt0.getSheetName() is %s%n", workbookName, sheetName);
    }
}
