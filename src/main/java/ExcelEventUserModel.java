import java.io.*;
import java.util.HashMap;
import java.util.Map;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;

/**
 * POI事件驱动模式示例
 *
 * http://poi.apache.org/components/spreadsheet/how-to.html#xssf_sax_api
 */
public class ExcelEventUserModel {
    public Map<Integer, Map<Integer, String>> data;
    private String rId;

    public ExcelEventUserModel() {
        this.data = new HashMap<>();
        this.rId = "rId1";
    }

    /**
     * 读取第一个sheet的数据
     *
     * @param filePath 文件路径
     * @return 表格数据
     */
    public Map<Integer, Map<Integer, String>> processOneSheet(String filePath) {

        this.data = new HashMap<>();

        InputStream in = null;
        OPCPackage pkg = null;
        SharedStringsTable sst = null;
        InputStream sheet = null;
        try {
            in = new BufferedInputStream(new FileInputStream(new File(filePath)));
            pkg = OPCPackage.open(in);
            XSSFReader r = new XSSFReader(pkg);
            // 通过读取workbook设置rid
            setRelationshipId(r);
            // 这里输出一下rId
            System.out.println("rId = " + rId);
            sst = r.getSharedStringsTable();
            XMLReader parser = fetchSheetParser(sst);
            sheet = r.getSheet(rId);
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (sheet != null) sheet.close();
            } catch (Exception closeException) {
                closeException.printStackTrace();
            }
            try {
                if (sst != null) sst.close();
            } catch (Exception closeException) {
                closeException.printStackTrace();
            }
            try {
                if (pkg != null) pkg.close();
            } catch (Exception closeException) {
                closeException.printStackTrace();
            }
            try {
                if (in != null) in.close();
            } catch (Exception closeException) {
                closeException.printStackTrace();
            }
        }
        return this.data;
    }

    private void setRelationshipId(XSSFReader r)
            throws IOException, InvalidFormatException, SAXException, ParserConfigurationException {
        InputStream workbookData = r.getWorkbookData();
        InputSource sheetSource = new InputSource(workbookData);
        XMLReader parser = WorkbookParser();
        parser.parse(sheetSource);
    }

    private XMLReader WorkbookParser() throws SAXException, ParserConfigurationException {
        XMLReader parser = SAXHelper.newXMLReader();
        ContentHandler handler = new WorkbookHandler();
        parser.setContentHandler(handler);
        return parser;
    }

    private class WorkbookHandler extends DefaultHandler {

        private WorkbookHandler() {}

        public void startElement(String uri, String localName, String name, Attributes attributes) {
            if ("sheet".equals(name)) {
                rId = attributes.getValue("r:id");
            }
        }

        public void endElement(String uri, String localName, String name) {}

    }

    private XMLReader fetchSheetParser(SharedStringsTable sst)
            throws SAXException, ParserConfigurationException {
        XMLReader parser = SAXHelper.newXMLReader();
        ContentHandler handler = new SheetHandler(sst);
        parser.setContentHandler(handler);
        return parser;
    }

    /**
     * See org.xml.sax.helpers.DefaultHandler javadocs
     */
    private class SheetHandler extends DefaultHandler {
        private SharedStringsTable sst;
        private String lastContents;
        private boolean nextIsString;
        private int rowIndex;
        private int colIndex;
        private Map<Integer, String> rowData;

        private SheetHandler(SharedStringsTable sst) {
            this.sst = sst;
            this.rowData = new HashMap<>();

            // 默认设置为第0行
            this.rowIndex = 0;
            // 默认设置为第0列
            this.colIndex = 0;
        }

        public void startElement(String uri, String localName, String name, Attributes attributes) {
            // 一行开始
            if ("row".equals(name)) {
                this.rowIndex = Integer.parseInt(attributes.getValue("r")) - 1;
            }

            // 单元格
            if ("c".equals(name)) {
                String cellType = attributes.getValue("t");
                this.nextIsString = "s".equals(cellType);
                this.colIndex = getColIndex(attributes.getValue("r"));
            }
            // Clear contents cache
            lastContents = "";
        }

        public void endElement(String uri, String localName, String name) {
            // Process the last contents as required.
            // Do now, as characters() may be called more than once
            if (nextIsString) {
                int idx = Integer.parseInt(lastContents);
                lastContents = sst.getItemAt(idx).getString();
                nextIsString = false;
            }
            // v => contents of a cell
            if ("v".equals(name) || "t".equals(name)) {
                // 放入行数据中，key=列数，value=单元格的值
                rowData.put(colIndex, lastContents);
            }

            // 一行的结束
            if ("row".equals(name)) {
                // 新的一行，存储上一行的数据
                data.put(rowIndex, rowData);
                this.rowData = new HashMap<>();
            }
        }

        public void characters(char[] ch, int start, int length) {
            lastContents += new String(ch, start, length);
        }

        /**
         * 转换表格引用为列编号，A-0，B-1
         *
         * @param cellReference 列引用，例：A1
         * @return 表格列位置，从0开始算
         */
        private int getColIndex(String cellReference) {
            String ref = cellReference.replaceAll("\\d+", "");
            int num;
            int result = 0;
            int length = ref.length();
            for (int i = 0; i < length; i++) {
                char ch = cellReference.charAt(length - i - 1);
                num = ch - 'A' + 1;
                num *= Math.pow(26, i);
                result += num;
            }
            return result - 1;
        }

        /**
         * 转换表格引用为行号
         * 
         * @param cellReference 列引用，例：A1
         * @return 行号，从0开始
         */
        private int getRowIndex(String cellReference) {
            String rowIndexStr = cellReference.replaceAll("[a-zA-Z]+", "");
            return Integer.parseInt(rowIndexStr) - 1;
        }


    }


}
