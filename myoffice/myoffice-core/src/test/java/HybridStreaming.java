import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * This demonstrates how a hybrid approach to workbook read can be taken, using
 * a mix of traditional XSSF and streaming one particular worksheet (perhaps one
 * which is too big for the ordinary DOM parse).
 */
public class HybridStreaming {

    private static final String SHEET_TO_STREAM = "large sheet";

    public static void main(String[] args)
            throws IOException, SAXException, ParserConfigurationException, OpenXML4JException {
        try (InputStream sourceBytes = new FileInputStream("/Users/qzz/业务追踪表1557385217666.xlsx")) {
            XSSFWorkbook workbook = new XSSFWorkbook(sourceBytes) {
                /**
                 * Avoid DOM parse of large sheet
                 */
                @Override
                public void parseSheet(java.util.Map<String, XSSFSheet> shIdMap, CTSheet ctSheet) {
                    if (!SHEET_TO_STREAM.equals(ctSheet.getName())) {
                        super.parseSheet(shIdMap, ctSheet);
                    }
                }
            };
            OPCPackage p = OPCPackage.open("/Users/qzz/业务追踪表1557385217666.xlsx", PackageAccess.READ);
            // Having avoided a DOM-based parse of the sheet, we can stream it instead.
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(p);

            XSSFReader xssfReader = new XSSFReader(p);
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            int index = 0;
            while (iter.hasNext()) {
                try (InputStream stream = iter.next()) {
                    String sheetName = iter.getSheetName();
                    //                    this.output.println();
                    System.out.println(sheetName + " [index=" + index + "]:");
                    DataFormatter formatter = new DataFormatter();
                    InputSource sheetSource = new InputSource(stream);
                    XMLReader sheetParser = SAXHelper.newXMLReader();
                    ContentHandler handler = new XSSFSheetXMLHandler(workbook.getStylesSource(), strings,
                            createSheetContentsHandler(), false);
                    sheetParser.setContentHandler(handler);
                    sheetParser.parse(sheetSource);
                }
                ++index;
            }

            workbook.close();
        }
    }

    private static SheetContentsHandler createSheetContentsHandler() {
        return new SheetContentsHandler() {

            @Override
            public void startRow(int rowNum) {
                System.out.println("row: " + rowNum);
            }

            @Override
            public void endRow(int rowNum) {
            }

            @Override
            public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            }
        };
    }
}