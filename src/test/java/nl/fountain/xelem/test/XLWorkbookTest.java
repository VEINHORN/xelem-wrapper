/*
 * Created on Sep 10, 2004
 *
 */
package nl.fountain.xelem.excel.ss;

import java.util.Iterator;
import java.util.NoSuchElementException;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.UnsupportedStyleException;
import nl.fountain.xelem.XFactory;
import nl.fountain.xelem.excel.DocumentProperties;
import nl.fountain.xelem.excel.DuplicateNameException;
import nl.fountain.xelem.excel.Workbook;
import nl.fountain.xelem.excel.Worksheet;
import nl.fountain.xelem.excel.XLElementTest;

/**
 *
 */
public class XLWorkbookTest extends XLElementTest {
    
    private Workbook wb;

    public static void main(String[] args) {
        junit.textui.TestRunner.run(XLWorkbookTest.class);
    }

    /*
     * @see TestCase#setUp()
     */
    protected void setUp() throws Exception {
        String configFileName =
            "testsuitefiles/XFactoryTest/XFactoryTest.xml";
        XFactory.setConfigurationFileName(configFileName);
        wb = new XLWorkbook("bestand");
    }
    
    public void testName() {
        Workbook wb = new XLWorkbook(null);
        assertEquals("null.xls", wb.getFileName());
        wb = new XLWorkbook();
        assertEquals(".xls", wb.getFileName());
    }
    
    public void testDocumentProperties() {
        String xml = xmlToString(wb, new GIO());
        assertEquals(-1, xml.indexOf("<o:DocumentProperties"));
        
        DocumentProperties dp = wb.getDocumentProperties();
        dp.setAuthor(this.getName());
        xml = xmlToString(wb, new GIO());
        assertTrue(xml.indexOf("<o:Author>testDocumentProperties</o:Author>") > 0);
        
        //System.out.println(xml);
    }
    
    public void testAddSheet() {
        Worksheet sheet1 = new SSWorksheet("Sheet foo");
        
        try {
            wb.addSheet(sheet1);
            wb.addSheet(sheet1);
            fail("Er kan geen tweede worksheet met eenzelfde naam worden toegevoegd.");
        } catch (DuplicateNameException e) {
            assertEquals("Duplicate name in worksheets collection: 'Sheet foo'."
                    , e.getMessage());
        }
    }
    
    public void testAddSheet_np() throws DuplicateNameException {
       Worksheet ws1 = wb.addSheet();
       Worksheet ws2 = wb.addSheet();
       Worksheet ws3 = wb.addSheet();
       assertSame(ws1, wb.getWorksheet("Sheet1"));
       assertSame(ws2, wb.getWorksheet("Sheet2"));
       assertSame(ws3, wb.getWorksheet("Sheet3"));
       wb.addSheet("Sheet4");
       wb.addSheet("foo bar");
       Worksheet ws6 = wb.addSheet();
       assertSame(ws6, wb.getWorksheet("Sheet6"));
    }
    
    public void testGetSheet() throws DuplicateNameException {
        assertNull(wb.getWorksheet("bla"));
        Worksheet sheet1 = new SSWorksheet("Sheet1");
        Worksheet sheet2 = new SSWorksheet("Sheet2");
        wb.addSheet(sheet1);
        wb.addSheet(sheet2);
        assertSame(sheet1, wb.getWorksheet("Sheet1"));
        assertSame(sheet2, wb.getWorksheet("Sheet2"));
    }
    
    public void testOrderOfSheets() throws DuplicateNameException {
        Worksheet z = wb.addSheet("Z");
        Worksheet a = wb.addSheet("A");
        Worksheet p = wb.addSheet("P");
        Iterator<Worksheet> iter = wb.getWorksheets().iterator();
        assertSame(z, iter.next());
        assertSame(a, iter.next());
        assertSame(p, iter.next());
        try {
            iter.next();
            fail("past limit.");
        } catch (NoSuchElementException e) {
        }
    }
    
    public void testRemoveSheet() throws DuplicateNameException {
        assertNull(wb.removeSheet("bla bla"));
        Worksheet ws1 = wb.addSheet("bla bla");
        assertSame(ws1, wb.removeSheet("bla bla"));
    }
    
    public void testGetSheetAtIndex() {
        assertNull(wb.getWorksheetAt(20));
    }
    
    public void testFilename() {
        assertEquals("bestand.xls", wb.getFileName());
        wb.setFileName("bestand.xml");
        assertEquals("bestand.xml", wb.getFileName());
    }
    
    
    public void testPrintComments() {
        assertTrue(wb.isPrintingElementComments());
        assertTrue(wb.isPrintingDocComments());
        wb.setPrintElementComments(false);
        assertTrue(!wb.isPrintingElementComments());
        wb.setPrintDocComments(false);
        assertTrue(!wb.isPrintingDocComments());
    }
    
    public void testMergeStyles() throws UnsupportedStyleException {
        wb.mergeStyles("wbNieuw", "b_yellow", "bold");
        wb.mergeStyles("wbNieuwer", "wbNieuw", "decimal2");
        wb.addSheet().addCell().setStyleID("wbNieuwer");
        
        String xml = xmlToString(wb, new GIO());
        //System.out.println(xml);
        assertTrue(xml.indexOf("<Style ss:ID=\"wbNieuwer\">") > 0);
        assertTrue(xml.indexOf("<Interior ss:Color=\"#FFFF00\" " +
        		"ss:Pattern=\"Solid\"/>") > 0);
        assertTrue(xml.indexOf("x:Family=\"Swiss\"") > 0);
        assertTrue(xml.indexOf("<NumberFormat ss:Format=\"0.00\"/>") > 0);       
    }
    

}
