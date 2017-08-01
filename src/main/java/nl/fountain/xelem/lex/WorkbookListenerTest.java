/*
 * Created on 8-apr-2005
 *
 */
package nl.fountain.xelem.lex;

import junit.framework.TestCase;
import nl.fountain.xelem.excel.Workbook;


public class WorkbookListenerTest extends TestCase {

    public static void main(String[] args) {
        junit.textui.TestRunner.run(WorkbookListenerTest.class);
    }
    
    public void testWorkbookFileName() {
        WorkbookListener l = new WorkbookListener();
        //l.startWorkbook("file:/G:/eclipse_ws/xelem/testsuitefiles/ReaderTest/reader.xml", "reader");
        l.startWorkbook("source");
        Workbook wb = l.getWorkbook();
        assertEquals("source", wb.getFileName());
        assertEquals("source", wb.getName());
        
        l.startWorkbook("file:/G:/eclipse_ws/xelem/testsuitefiles/ReaderTest/reader.xml");
        wb = l.getWorkbook();
        assertEquals("file:/G:/eclipse_ws/xelem/testsuitefiles/ReaderTest/reader.xml", wb.getFileName());
        assertEquals("reader", wb.getName());
    }

}
