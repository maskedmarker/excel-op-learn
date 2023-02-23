package org.example.learn.excel.poi.workbook;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.OutputStream;

public class WorkbookBaseTest {

    public static final String FILE_NAME_PREFIX = "workbook-";
    public static final String WORKBOOKNAME = FILE_NAME_PREFIX + "base.xls";

    @Test
    public void test1() throws Exception {

        Workbook wb = new HSSFWorkbook();  // or new XSSFWorkbook();
        Sheet sheet1 = wb.createSheet("new sheet");
        Sheet sheet2 = wb.createSheet("second sheet");
        // Note that sheet name is Excel must not exceed 31 characters
        // and must not contain any of the any of the following characters:
        // 0x0000
        // 0x0003
        // colon (:)
        // backslash (\)
        // asterisk (*)
        // question mark (?)
        // forward slash (/)
        // opening square bracket ([)
        // closing square bracket (])
        String safeName = WorkbookUtil.createSafeSheetName("[O'Brien's sales*?]"); // returns " O'Brien's sales   "
        Sheet sheet3 = wb.createSheet(safeName);

        try (OutputStream fileOut = new FileOutputStream(WORKBOOKNAME)) {
            wb.write(fileOut);
        }

        System.out.printf("%s written successfully%n", WORKBOOKNAME);
    }
}
