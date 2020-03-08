import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class Test {
    public static void main(String[] args) throws IOException {
        // turn off auto-flushing and accumulate all rows in memory
        SXSSFWorkbook wb = new SXSSFWorkbook(-1);
        Sheet sh = wb.createSheet();
        for (int rownum = 0; rownum < 1000; rownum++) {
            Row row = sh.createRow(rownum);
            for (int cellnum = 0; cellnum < 10; cellnum++) {
                Cell cell = row.createCell(cellnum);
                String address = new CellReference(cell).formatAsString();
                cell.setCellValue(address);
            }
            // manually control how rows are flushed to disk
            if (rownum % 100 == 0) {
                ((SXSSFSheet) sh).flushRows(100); // retain 100 last rows and flush all others
            }
        }
        FileOutputStream out = new FileOutputStream("E:\\tmp\\test\\excelTest\\sxssf.xlsx");
        wb.write(out);
        out.close();
        // dispose of temporary files backing this workbook on disk
        wb.dispose();
    }
}
