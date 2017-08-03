/*
 * Created on 6-apr-2005
 *
 */
package nl.fountain.xelem.lex;

import java.util.List;
import java.util.Map;

import junit.framework.TestCase;
import nl.fountain.xelem.excel.AutoFilter;
import nl.fountain.xelem.excel.Cell;
import nl.fountain.xelem.excel.Column;
import nl.fountain.xelem.excel.DocumentProperties;
import nl.fountain.xelem.excel.ExcelWorkbook;
import nl.fountain.xelem.excel.NamedRange;
import nl.fountain.xelem.excel.Row;
import nl.fountain.xelem.excel.Table;
import nl.fountain.xelem.excel.Worksheet;
import nl.fountain.xelem.excel.WorksheetOptions;

/**
 *
 */
public class ExcelReaderListenerTest extends TestCase {
    
    protected boolean docStart;
    protected String sysId;
    protected int wbCounter;
    protected DocumentProperties props;
    protected int propsCounter;
    protected ExcelWorkbook excelWB;
    protected int excelWBCounter;
    protected NamedRange workbookNamedRange;
    protected int workbookNamedRangeCounter;
    protected String lastWorksheetName;
    protected int lastWorksheetIndex;
    protected int worksheetCounter;
    protected NamedRange worksheetNamedRange;
    protected int worksheetNamedRangeCounter;
    protected String lastWorksheetWithNamedRange;
    protected String lastSheetWithTable;
    protected int tableCounter;
    protected int lastExpandedRowCount;
    protected int lastExpandedColumnCount;
    protected int columnCounter;
    protected int lastColumnIndex;

    public static void main(String[] args) {
        junit.textui.TestRunner.run(ExcelReaderListenerTest.class);
    }
    
    public void testAddRemoveClearListener() throws Exception {
        ExcelReader xlReader = new ExcelReader();
        List<ExcelReaderListener> listeners = xlReader.getListeners();
        assertEquals(0, listeners.size());
        
        Listener listener = new Listener();
        xlReader.addExcelReaderListener(listener);
        assertEquals(1, listeners.size());
        xlReader.addExcelReaderListener(listener);
        assertEquals(1, listeners.size());
        
        Listener listener2 = new Listener();
        xlReader.addExcelReaderListener(listener2);
        assertEquals(2, listeners.size());
        assertSame(listener, listeners.get(0));
        assertSame(listener2, listeners.get(1));
        
        assertTrue(xlReader.removeExcelReaderListener(listener));
        assertFalse(xlReader.removeExcelReaderListener(new Listener()));
        assertEquals(1, listeners.size());
        
        xlReader.clearExcelReaderListeners();
        assertEquals(0, listeners.size());
    }
    
    public void testListener() throws Exception {
        ExcelReader xlReader = new ExcelReader();
        Listener listener = new Listener();
        xlReader.addExcelReaderListener(listener);

        xlReader.read("testsuitefiles/ReaderTest/reader.xml");
        
        assertEquals(1, wbCounter);
        assertTrue(sysId.endsWith("testsuitefiles/ReaderTest/reader.xml"));
        
        assertEquals(1, propsCounter);
        assertEquals("Asterix", props.getAuthor());
        
        assertEquals(1, excelWBCounter);
        assertEquals(360, excelWB.getWindowTopX());
        
        assertEquals(1, workbookNamedRangeCounter);
        assertEquals("='Tom Poes'!R9C4:R11C4", workbookNamedRange.getRefersTo());
        assertEquals("foo", workbookNamedRange.getName());
        
        assertEquals(5, worksheetCounter);
        assertEquals(4, lastWorksheetIndex);
        assertEquals("window", lastWorksheetName);
        
        assertEquals(1, worksheetNamedRangeCounter);
        assertEquals("_FilterDatabase", worksheetNamedRange.getName());
        assertTrue(worksheetNamedRange.isHidden());
        assertEquals("Tom Poes", lastWorksheetWithNamedRange);
        
        assertEquals(4, tableCounter);
        assertEquals("window", lastSheetWithTable);
        assertEquals(7, lastExpandedRowCount);
        assertEquals(4, lastExpandedColumnCount);
        
        assertEquals(6, columnCounter);
        assertEquals(16, lastColumnIndex);
    }
    
    
    private class Listener implements ExcelReaderListener {
        
        public void startDocument() {
            docStart = true;
        }
        
        public void processingInstruction(String target, String data) {
            assertTrue(docStart);
            //System.out.println("target=" + target + " data=" + data);
        }
        
        public void startWorkbook(String systemID) {
            //System.out.println("Workbook sytemID=" + systemID + " name=" + workbookName);
            sysId = systemID;
            assertTrue(sysId.endsWith("testsuitefiles/ReaderTest/reader.xml"));
            wbCounter++;
        }

        public void setDocumentProperties(DocumentProperties docprops) {
            //System.out.println(docprops);
            
            props = docprops;
            propsCounter++;
        }

        public void setExcelWorkbook(ExcelWorkbook xlwb) {
            //System.out.println(xlwb);
            assertEquals("Asterix", props.getAuthor());
            excelWB = xlwb;
            excelWBCounter++;
        }

        public void setNamedRange(NamedRange nr) {
            //System.out.println(nr);
            assertEquals(360, excelWB.getWindowTopX());
            workbookNamedRange = nr;
            workbookNamedRangeCounter++;
        }

        public void startWorksheet(int sheetIndex, Worksheet sheet) {
//            System.out.println("Worksheet index=" + sheetIndex 
//                    + " name=" + sheet.getName()
//                    + " " + sheet);
            lastWorksheetIndex = sheetIndex;
            lastWorksheetName = sheet.getName();
            worksheetCounter++;
        }

        public void setNamedRange(int sheetIndex, String sheetName, NamedRange nr) {
//            System.out.println("NamedRange sheetName=" + sheetName + " " + nr);
            lastWorksheetWithNamedRange = sheetName;
            worksheetNamedRange = nr;
            worksheetNamedRangeCounter++;
        }

        public void startTable(int sheetIndex, String sheetName, Table table) {
//            System.out.println("Table on " + sheetName 
//                    + ", rowCount=" + table.getExpandedRowCount() 
//                    + " columnCount=" + table.getExpandedColumnCount()
//                    + " " + table);
            lastSheetWithTable = sheetName;
            lastExpandedRowCount = table.getExpandedRowCount();
            lastExpandedColumnCount = table.getExpandedColumnCount();
            tableCounter++;
        }

        public void setColumn(int sheetIndex, String sheetName, Column column) {
//            System.out.println("Column on " + sheetName
//                    + ", columnIndex=" + column.getIndex() + " " + column);
            lastColumnIndex = column.getIndex();
            columnCounter++;
        }

        public void setRow(int sheetIndex, String sheetName, Row row) {
//            System.out.println("Row on " + sheetName
//                    + ", rowIndex=" + row.getIndex() + " " + row);
        }

        public void setCell(int sheetIndex, String sheetName, int rowIndex, Cell cell) {
//            System.out.println("Cell on " + sheetName
//                    + ", rowIndex=" + rowIndex
//                    + ", cellIndex=" + cell.getIndex()
//                    + ", data=" + cell.getData());
        }

        public void setWorksheetOptions(int sheetIndex, String sheetName, WorksheetOptions wso) {
//            System.out.println("WorksheetOptions on " + sheetName
//                    + " sheetIndex=" + sheetIndex
//                    + " " + wso);
        }

        public void setAutoFilter(int sheetIndex, String sheetName, AutoFilter autoFilter) {
//            System.out.println("AutoFilter on " + sheetName
//                    + " sheetIndex=" + sheetIndex
//                    + " " + autoFilter);
        }
        
        public void endWorksheet(int sheetIndex, String sheetName) {
//          System.out.println("End of Worksheet. index=" + sheetIndex 
//                  + " name=" + sheetName);  
          assertEquals(lastWorksheetIndex, sheetIndex);
          assertEquals(lastWorksheetName, sheetName);
        }
        
        public void endDocument(Map<String, String> prefixMap) {
            assertEquals(9, prefixMap.size());
        }
        
    }

}
