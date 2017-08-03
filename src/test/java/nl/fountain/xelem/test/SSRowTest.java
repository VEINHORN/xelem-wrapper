/*
 * Created on Sep 8, 2004
 *
 */
package nl.fountain.xelem.excel.ss;

import java.util.Iterator;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.excel.Cell;
import nl.fountain.xelem.excel.Row;
import nl.fountain.xelem.excel.XLElementTest;

/**
 *
 */
public class SSRowTest extends XLElementTest {
    
    private Row row;

    public static void main(String[] args) {
        junit.textui.TestRunner.run(SSRowTest.class);
    }

    /*
     * @see TestCase#setUp()
     */
    protected void setUp() throws Exception {
        row = new SSRow();
    }
    
    /*
     * Een nieuwe rij moet de eerste cel op index 1 plaatsen enz.
     */
    public void testAddCell() {
        assertEquals(0, row.size());
        Cell cell = row.addCell();
        assertEquals(1, row.size());
        assertSame(cell, row.getCellMap().get(new Integer(1)));
        Cell cell2 = row.addCell();
        assertEquals(2, row.size());
        assertSame(cell2, row.getCellMap().get(new Integer(2)));
        assertNotSame(cell, cell2);
        
        Iterator<Cell> iter = row.getCells().iterator();
        assertSame(cell, iter.next());
        assertSame(cell2, iter.next());
        try {
            iter.next();
            fail("Alle cellen zijn doorlopen");
        } catch (Exception e) {
        }
    }
    
    /*
     * De methode addCell(index) plaatst cellen op de gegeven index.
     */
    public void testAddCell_index() {
        assertEquals(0, row.size());
        Cell cell = row.addCellAt(5);
        assertEquals(1, row.size());
        assertSame(cell, row.getCellMap().get(new Integer(5)));
        
        Cell cell6 = row.addCell();
        assertEquals(2, row.size());
        assertSame(cell6, row.getCellMap().get(new Integer(6)));
        
        Cell cell3 = row.addCellAt(3);
        assertEquals(3, row.size());
        assertSame(cell3, row.getCellMap().get(new Integer(3)));
        
        Iterator<Cell> iter = row.getCells().iterator();
        assertSame(cell3, iter.next());
        assertSame(cell, iter.next());
        assertSame(cell6, iter.next());
        try {
            iter.next();
            fail("Alle cellen zijn doorlopen");
        } catch (Exception e) {
        }
        
        Cell cell5 = row.addCellAt(5);
        assertEquals(3, row.size());
        assertSame(cell5, row.getCellMap().get(new Integer(5)));
        
        Cell cell7 = row.addCell();
        assertEquals(4, row.size());
        assertSame(cell7, row.getCellMap().get(new Integer(7)));
        
        iter = row.getCells().iterator();
        assertSame(cell3, iter.next());
        assertSame(cell5, iter.next());
        assertSame(cell6, iter.next());
        assertSame(cell7, iter.next());
        try {
            iter.next();
            fail("Alle cellen zijn doorlopen");
        } catch (Exception e) {
        }
    }
    
    public void testAddCell_Cell() {
        assertEquals(0, row.size());
        Cell cell = new SSCell();
        Cell returnCell = row.addCell(cell);
        assertEquals(1, row.size());
        assertSame(cell, returnCell);
        assertSame(cell, row.getCellMap().get(new Integer(1)));
    }
    
    public void testAddCell_index_Cell() {
        assertEquals(0, row.size());
        Cell cell = new SSCell();
        Cell returnCell = null;
        try {
            returnCell = row.addCellAt(-1, cell);
            fail("geenexceptie gegooid.");
        } catch (IndexOutOfBoundsException e) {
            //
        }
        assertEquals(0, row.size());
        assertNull(returnCell);
        
        Cell returnCell2 = row.addCellAt(5, cell);
        assertEquals(1, row.size());
        assertSame(cell, returnCell2);
        assertSame(cell, row.getCellMap().get(new Integer(5)));
    }
    
    public void testRemoveCell() {
       row.addCell();
       Cell cell = row.addCell();
       row.addCell();
       assertEquals(3, row.size());
       assertSame(cell, row.removeCellAt(2));
       assertEquals(2, row.size());
       
       row.addCell();
       Iterator<Integer> iter = row.getCellMap().keySet().iterator();
       assertEquals(new Integer(1), iter.next());
       assertEquals(new Integer(3), iter.next());
       assertEquals(new Integer(4), iter.next());
       try {
           iter.next();
           fail("Alle cellen zijn doorlopen");
       } catch (Exception e) {
       }
    }
    
    public void testAssemble() {
        row.setStyleID("foo");
        row.addCell().setStyleID("bar");
        row.addCellAt(5);
        GIO gio = new GIO();
        String xml = xmlToString(row, gio);
        
        assertTrue(xml.indexOf("<ss:Row ss:StyleID=\"foo\">") > 0);
        assertTrue(xml.indexOf("<ss:Cell ss:StyleID=\"bar\"/>") > 0);
        assertTrue(xml.indexOf("<ss:Cell ss:Index=\"5\"/>") > 0);
        assertEquals(2, gio.getStyleIDSet().size());
                
        //System.out.println(xml);
    }
    
}
