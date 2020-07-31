package excel.read;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class TestExcelEventUserModel {
    public static void main(String[] args) throws IOException {
        // createExcelBySXSSFWorkbook();

        String path1 = "E:\\tmp\\test\\excelTest\\单Sheet_java代码创建_SXSSFWorkbook.xlsx";
        System.out.println("path1=" + path1);
        printExcelData(new excel.read.ExcelEventUserModel().processOneSheet(path1));

        String path2 = "E:\\tmp\\test\\excelTest\\多Sheet_Excel软件创建.xlsx";
        System.out.println("path2=" + path2);
        printExcelData(new excel.read.ExcelEventUserModel().processOneSheet(path2));

        String path3 = "E:\\tmp\\test\\excelTest\\多Sheet_java代码创建.xlsx";
        System.out.println("path3=" + path3);
        printExcelData(new excel.read.ExcelEventUserModel().processOneSheet(path3));
    }

    /**
     * https://poi.apache.org/components/spreadsheet/how-to.html#sxssf
     * 
     * 通过SXSSFWorkbook写Excel，支持大量数据（代码为POI官网示例，未做优化，仅修改了总行数和总列数）
     */
    private static void createExcelBySXSSFWorkbook() throws IOException {
        // turn off auto-flushing and accumulate all rows in memory
        SXSSFWorkbook wb = new SXSSFWorkbook(-1);
        Sheet sh = wb.createSheet();
        for (int rownum = 0; rownum < 3; rownum++) {
            Row row = sh.createRow(rownum);
            for (int cellnum = 0; cellnum < 3; cellnum++) {
                Cell cell = row.createCell(cellnum);
                String address = new CellReference(cell).formatAsString();
                cell.setCellValue(address);
            }
            // manually control how rows are flushed to disk
            if (rownum % 100 == 0) {
                ((SXSSFSheet) sh).flushRows(100); // retain 100 last rows and flush all others
            }
        }
        FileOutputStream out = new FileOutputStream(
                "E:\\tmp\\test\\excelTest\\单Sheet_java代码创建_SXSSFWorkbook.xlsx");
        wb.write(out);
        out.close();
        // dispose of temporary files backing this workbook on disk
        wb.dispose();
    }

    private static void printExcelData(Map<Integer, Map<Integer, String>> map) {
        for (Map.Entry<Integer, Map<Integer, String>> entry : map.entrySet()) {
            StringBuilder rowData = new StringBuilder();
            for (Map.Entry<Integer, String> cell : entry.getValue().entrySet()) {
                rowData.append(cell.getValue()).append(" ");
            }
            System.out.println(rowData.toString());
        }
    }

}
