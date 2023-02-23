package org.example.learn.excel.poi.cell;

import org.apache.poi.ss.usermodel.*;
import org.junit.Test;

import java.io.File;
import java.io.IOException;

public class CellGetTest {

    private static final String FILE_NAME_PREFIX = "cell-get-";

    private static final int MY_MINIMUM_COLUMN_COUNT = 1000;

    /**
     * Iterate over cells, with control of missing / blank cells
     */
    @Test
    public void test() throws IOException {
        String workbookName = FILE_NAME_PREFIX + "base.xls";

        Workbook wb = WorkbookFactory.create(new File(workbookName));

        Sheet sheet = wb.getSheetAt(0);

        // Decide which rows to process
        int rowStart = Math.min(15, sheet.getFirstRowNum());
        int rowEnd = Math.max(1400, sheet.getLastRowNum());
        for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
            Row r = sheet.getRow(rowNum);
            if (r == null) {
                // This whole row is empty
                // Handle it as needed
                continue;
            }
            int lastColumn = Math.max(r.getLastCellNum(), MY_MINIMUM_COLUMN_COUNT);
            for (int cn = 0; cn < lastColumn; cn++) {
                Cell c = r.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (c == null) {
                    // The spreadsheet is empty in this cell
                } else {
                    // Do something useful with the cell's contents
                }
            }
        }
    }
}
