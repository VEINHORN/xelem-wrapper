/*
 * Created on Sep 8, 2004
 *
 */
package nl.fountain.xelem.excel.ss;

import java.util.Iterator;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.excel.Row;
import nl.fountain.xelem.excel.Table;
import nl.fountain.xelem.excel.XLElementTest;

/**
 *
 */
public class SSTableTest extends XLElementTest {
    
    private Table table;

    public static void main(String[] args) {
        junit.textui.TestRunner.run(SSTableTest.class);
    }

    /*
     * @see TestCase#setUp()
     */
    protected void setUp() throws Exception {
        table = new SSTable();
    }
    
    /*
     * Een nieuwe tabel moet de eerste rij op index 1 plaatsen enz.
     */
    public void testAddRow() {
        assertEquals(0, table.rowCount());
        Row row = table.addRow();
        assertEquals(1, table.rowCount());
        assertSame(row, table.getRowMap().get(new Integer(1)));
        Row row2 = table.addRow();
        assertEquals(2, table.rowCount());
        assertSame(row2, table.getRowMap().get(new Integer(2)));
        assertNotSame(row, row2);
        
        Iterator<Row> iter = table.getRows().iterator();
        assertSame(row, iter.next());
        assertSame(row2, iter.next());
        try {
            iter.next();
            fail("Alle rijen zijn doorlopen");
        } catch (Exception e) {
        }
    }
    
    /*
     * De methode addRow(index) plaatst rijen op de gegeven index.
     */
    public void testAddRow_index() {
        assertEquals(0, table.rowCount());
        Row row = table.addRowAt(5);
        assertEquals(1, table.rowCount());
        assertSame(row, table.getRowMap().get(new Integer(5)));
        
        Row row6 = table.addRow();
        assertEquals(2, table.rowCount());
        assertSame(row6, table.getRowMap().get(new Integer(6)));
        
        Row row3 = table.addRowAt(3);
        assertEquals(3, table.rowCount());
        assertSame(row3, table.getRowMap().get(new Integer(3)));
        
        Iterator<Row> iter = table.getRows().iterator();
        assertSame(row3, iter.next());
        assertSame(row, iter.next());
        assertSame(row6, iter.next());
        try {
            iter.next();
            fail("Alle rijen zijn doorlopen");
        } catch (Exception e) {
        }
        
        Row row5 = table.addRowAt(5);
        assertEquals(3, table.rowCount());
        assertSame(row5, table.getRowMap().get(new Integer(5)));
        
        Row row7 = table.addRow();
        assertEquals(4, table.rowCount());
        assertSame(row7, table.getRowMap().get(new Integer(7)));
        
        iter = table.getRows().iterator();
        assertSame(row3, iter.next());
        assertSame(row5, iter.next());
        assertSame(row6, iter.next());
        assertSame(row7, iter.next());
        try {
            iter.next();
            fail("Alle rijen zijn doorlopen");
        } catch (Exception e) {
        }
    }
    
    public void testAddRow_Row() {
        assertEquals(0, table.rowCount());
        Row row = new SSRow();
        Row returnRow = table.addRow(row);
        assertEquals(1, table.rowCount());
        assertSame(row, returnRow);
        assertSame(row, table.getRowMap().get(new Integer(1)));
    }
    
    public void testAddRow_index_Row() {
        assertEquals(0, table.rowCount());
        Row row = new SSRow();
        table.addRow(row);
        try {
            table.addRowAt(-1, row);
            fail("geen exceptie gegooid");
        } catch (IndexOutOfBoundsException e) {
            assertEquals("rowIndex = -1", e.getMessage());
        }
        assertEquals(1, table.rowCount());
        assertSame(row, table.getRowMap().get(new Integer(1)));
        
        Row returnRow2 = table.addRowAt(5, row);
        assertEquals(2, table.rowCount());
        assertSame(row, returnRow2);
        assertSame(row, table.getRowMap().get(new Integer(5)));
    }
    
    public void testRemoveRow() {
       table.addRow();
       Row row = table.addRow();
       table.addRow();
       assertEquals(3, table.rowCount());
       assertSame(row, table.removeRowAt(2));
       assertEquals(2, table.rowCount());
       
       table.addRow();
       Iterator<Integer> iter = table.getRowMap().keySet().iterator();
       assertEquals(new Integer(1), iter.next());
       assertEquals(new Integer(3), iter.next());
       assertEquals(new Integer(4), iter.next());
       try {
           iter.next();
           fail("Alle rijen zijn doorlopen");
       } catch (Exception e) {
       }
    }
    
    
    
    public void testIterator() {
       Row row1 = table.addRow();
       Row row2 = table.addRow();
       
       Iterator<Row> iter = table.rowIterator();
       assertEquals(row1, iter.next());
       assertEquals(row2, iter.next());
       try {
        iter.next();
        fail("alle rijen doorlopen");
       } catch (Exception e) {
           //
       }
       iter = table.rowIterator();
       iter.next();
       try {
           iter.remove();
           fail("should throw UnsupportedPerationException.");
       } catch (UnsupportedOperationException e1) {
           //
       }
    }
    
    public void testIterator2() {
        Row row1 = table.addRow();
        Row row2 = table.addRowAt(5);
        Row row3 = table.addRowAt(7, row2);
        
        Iterator<Row> iter = table.rowIterator();
        Row r1 = (Row) iter.next();
        assertEquals(row1, r1);
        
        Row r5 = (Row) iter.next();
        assertEquals(row2, r5);
        
        Row r7 = (Row) iter.next();
        assertEquals(row2, r7);
        assertEquals(row3, r7);
    }
    
    public void testAssemble() {
        table.setStyleID("foo");
        table.addColumnAt(5).setStyleID("col");
        table.addRowAt(7).addCell(new Double(1.2345), "currency");
        GIO gio = new GIO();
        String xml = xmlToString(table, gio);
        
        assertTrue(xml.indexOf("<ss:Table ss:StyleID=\"foo\">") > 0);
        assertTrue(xml.indexOf("<ss:Column ss:Index=\"5\" ss:StyleID=\"col\"") > 0);
        assertTrue(xml.indexOf("<ss:Row ss:Index=\"7\">") > 0);
        assertTrue(xml.indexOf("<ss:Cell ss:StyleID=\"currency\">") > 0);
        assertTrue(xml.indexOf("<Data ss:Type=\"Number\">1.2345</Data>") > 0);
        assertEquals(3, gio.getStyleIDSet().size());
        
        //System.out.println(xml);
    }

}
