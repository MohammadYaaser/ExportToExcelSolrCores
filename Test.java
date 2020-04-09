import com.sun.org.apache.xerces.internal.parsers.SAXParser;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import java.io.BufferedInputStream;
import java.io.InputStream;
import java.util.Collections;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;


public class LowMemoryExcelFileReader {

    private String file;

    public LowMemoryExcelFileReader(String file) {
        this.file = file;
    }

    public List<String[]> read() {
        try {
            return processFirstSheet(file);
        } catch (Exception e) {
           throw new RuntimeException(e);
        }
    }

    private List<String []> readSheet(Sheet sheet) {
        List<String []> res = new LinkedList<>();
        Iterator<Row> rowIterator = sheet.rowIterator();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            int cellsNumber = row.getLastCellNum();
            String [] cellsValues = new String[cellsNumber];

            Iterator<Cell> cellIterator = row.cellIterator();
            int cellIndex = 0;

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                cellsValues[cellIndex++] = cell.getStringCellValue();
            }

            res.add(cellsValues);
        }
        return res;
    }

    public String getFile() {
        return file;
    }

    public void setFile(String file) {
        this.file = file;
    }

    private List<String []> processFirstSheet(String filename) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename, PackageAccess.READ);
        XSSFReader r = new XSSFReader(pkg);
        SharedStringsTable sst = r.getSharedStringsTable();

        SheetHandler handler = new SheetHandler(sst);
        XMLReader parser = fetchSheetParser(handler);
        Iterator<InputStream> sheetIterator = r.getSheetsData();

        if (!sheetIterator.hasNext()) {
            return Collections.emptyList();
        }

        InputStream sheetInputStream = sheetIterator.next();
        BufferedInputStream bisSheet = new BufferedInputStream(sheetInputStream);
        InputSource sheetSource = new InputSource(bisSheet);
        parser.parse(sheetSource);
        List<String []> res = handler.getRowCache();
        bisSheet.close();
        return res;
    }

    public XMLReader fetchSheetParser(ContentHandler handler) throws SAXException {
        XMLReader parser = new SAXParser();
        parser.setContentHandler(handler);
        return parser;
    }

    /**
     * See org.xml.sax.helpers.DefaultHandler javadocs
     */
    private static class SheetHandler extends DefaultHandler {

        private static final String ROW_EVENT = "row";
        private static final String CELL_EVENT = "c";

        private SharedStringsTable sst;
        private String lastContents;
        private boolean nextIsString;

        private List<String> cellCache = new LinkedList<>();
        private List<String[]> rowCache = new LinkedList<>();

        private SheetHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        public void startElement(String uri, String localName, String name,
                                 Attributes attributes) throws SAXException {
            // c => cell
            if (CELL_EVENT.equals(name)) {
                String cellType = attributes.getValue("t");
                if(cellType != null && cellType.equals("s")) {
                    nextIsString = true;
                } else {
                    nextIsString = false;
                }
            } else if (ROW_EVENT.equals(name)) {
                if (!cellCache.isEmpty()) {
                    rowCache.add(cellCache.toArray(new String[cellCache.size()]));
                }
                cellCache.clear();
            }

            // Clear contents cache
            lastContents = "";
        }

        public void endElement(String uri, String localName, String name)
                throws SAXException {
            // Process the last contents as required.
            // Do now, as characters() may be called more than once
            if(nextIsString) {
                int idx = Integer.parseInt(lastContents);
                lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                nextIsString = false;
            }

            // v => contents of a cell
            // Output after we've seen the string contents
            if(name.equals("v")) {
                cellCache.add(lastContents);
            }
        }

        public void characters(char[] ch, int start, int length)
                throws SAXException {
            lastContents += new String(ch, start, length);
        }

        public List<String[]> getRowCache() {
            return rowCache;
        }
    }

    public static void main(String[] args){
        
    }
}