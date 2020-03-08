import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) throws Exception {
        ExcelEventUserModel model = new ExcelEventUserModel();
        System.out.println("------读取使用Microsoft Excel创建的表格------");
        String MSExcelPath = "E:\\tmp\\test\\excelTest\\test-MSExcel.xlsx";
        Map<Integer, Map<Integer, String>> msExcelData = model.processOneSheet(MSExcelPath);
        printExcelData(msExcelData);

        System.out.println("------读取使用代码创建的表格------");
        String javaExcelPath = "E:\\tmp\\test\\excelTest\\test-JavaExcel.xlsx";
        createExcel(javaExcelPath);
        Map<Integer, Map<Integer, String>> javaExcelData = model.processOneSheet(javaExcelPath);
        printExcelData(javaExcelData);
    }

    private static void printExcelData(Map<Integer, Map<Integer, String>> map) {
        for (Map.Entry<Integer, Map<Integer, String>> entry : map.entrySet()) {
            StringBuilder rowData = new StringBuilder();
            for (Map.Entry<Integer, String> cell : entry.getValue().entrySet()) {
                rowData.append(cell.getKey()).append(": ").append(cell.getValue()).append(" ");
            }
            int row = entry.getKey();
            System.out.println(row + "-->" + rowData.toString());
        }
    }

    private static void createExcel(String filePath) throws Exception {

        createDirectory(filePath);

        OutputStream fileOut = null;
        XSSFWorkbook wb = null;
        try {
            fileOut = new FileOutputStream(filePath);
            wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.createSheet("Sheet1");

            XSSFRow row_0 = sheet.createRow(0);
            row_0.createCell(0).setCellValue("姓名");
            row_0.createCell(1).setCellValue("职位");

            XSSFRow row_1 = sheet.createRow(1);
            row_1.createCell(0).setCellValue("千手纲手");
            row_1.createCell(1).setCellValue("五代火影");

            XSSFRow row_2 = sheet.createRow(2);
            row_2.createCell(0).setCellValue("旗木卡卡西");
            row_2.createCell(1).setCellValue("六代火影");

            wb.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (wb != null) wb.close();
            } catch (Exception ex) {
                ex.printStackTrace();
            }
            try {
                if (fileOut != null) fileOut.close();
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }

    }

    private static void createDirectory(String filePath) throws Exception {
        if (filePath == null || "".equals(filePath.trim())) {
            throw new Exception("filePath is empty");
        }
        File file;
        if (filePath.contains(".")) {
            file = new File(filePath.substring(0, filePath.lastIndexOf(File.separator)));
        } else {
            file = new File(filePath);
        }
        if (!file.exists()) file.mkdirs();
    }

}
