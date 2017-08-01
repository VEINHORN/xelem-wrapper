/*
 * Created on 16-apr-2005
 *
 */
package nl.fountain.xelem.lex;

import java.io.File;

import junit.framework.TestCase;
import nl.fountain.xelem.Area;
import nl.fountain.xelem.XSerializer;
import nl.fountain.xelem.excel.Cell;
import nl.fountain.xelem.excel.Row;
import nl.fountain.xelem.excel.Workbook;
import nl.fountain.xelem.excel.Worksheet;


public class DefaultExcelReaderFilterTest extends TestCase {
    
    ExcelReader reader;

    public static void main(String[] args) {
        junit.textui.TestRunner.run(DefaultExcelReaderFilterTest.class);
    }
    
    public void testSwapAreas() throws Exception {
        reader = new ExcelReader();
        WorkbookListener wbListener = new WorkbookListener();
        ExcelReaderFilter saf = new SwapAreaFilter();
        saf.addExcelReaderListener(wbListener);
        reader.addExcelReaderListener(saf);
        reader.read("testsuitefiles/ReaderTest/reader.xml");
        Workbook wb = wbListener.getWorkbook();
        
        Worksheet sheet = wb.getWorksheetAt(0);
        assertTrue(sheet.hasColumnAt(1));
        //assertFalse(sheet.hasColumnAt(7));
        Cell cell = sheet.getCellAt(6, 2);
        assertEquals("=R[-2]C[1]+R[-1]C[1]", cell.getFormula());
        //assertFalse(sheet.hasCellAt(10, 2));
              
        sheet = wb.getWorksheetAt(1);
        //assertFalse(sheet.hasCellAt(4, 2));
        cell = sheet.getCellAt(11, 7);
        //assertEquals("vijgje", cell.getData());
        //assertEquals(1, sheet.getTable().rowCount());
        Row row = sheet.getRowAt(11);
        assertSame(cell, row.getCellAt(7));
        assertEquals(1, row.size());
        
        File out = new File("testoutput/ReaderTest/filterTest.xls");
        new XSerializer().serialize(wb, out);
    }
    
    private class SwapAreaFilter extends DefaultExcelReaderFilter {
        
        public void startWorksheet(int sheetIndex, Worksheet sheet) {
            switch (sheetIndex) {
//            	case 0:
//            	    reader.clearReadArea();
//            	    break;
            	case 1:
            	    reader.setReadArea(new Area("A1:C6"));
            	    break;
            	case 2:
            	    reader.setReadArea(new Area("G11:G11"));
            	    break;
            	default:
            	    reader.clearReadArea();
            }
            super.startWorksheet(sheetIndex, sheet);
        }
    }

}
