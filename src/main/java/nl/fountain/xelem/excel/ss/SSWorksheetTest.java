/*
 * Created on Sep 8, 2004
 *
 */
package nl.fountain.xelem.excel.ss;

import nl.fountain.xelem.CellPointer;
import nl.fountain.xelem.GIO;
import nl.fountain.xelem.excel.Cell;
import nl.fountain.xelem.excel.Column;
import nl.fountain.xelem.excel.Row;
import nl.fountain.xelem.excel.Table;
import nl.fountain.xelem.excel.Worksheet;
import nl.fountain.xelem.excel.WorksheetOptions;
import nl.fountain.xelem.excel.XLElementTest;

/**
 *  
 */
public class SSWorksheetTest extends XLElementTest {

    private Worksheet ws;

    public static void main(String[] args) {
        junit.textui.TestRunner.run(SSWorksheetTest.class);
    }

    /*
     * @see TestCase#setUp()
     */
    protected void setUp() throws Exception {
        ws = new SSWorksheet("Sheet1");
    }

    public void testConstructor() {
        assertEquals("Sheet1", ws.getName());
        assertEquals(1, ws.getCellPointer().getHorizontalStepDistance());
        assertEquals(1, ws.getCellPointer().getVerticalStepDistance());
        assertEquals(CellPointer.MOVE_RIGHT, ws.getCellPointer().getMovement());
        Worksheet ws2 = new SSWorksheet("Blad1");
        assertEquals("Blad1", ws2.getName());
        Worksheet ws3 = new SSWorksheet("this & that");
        assertEquals("this & that", ws3.getName());
        Worksheet ws4 = new SSWorksheet(null);
        assertEquals(null, ws4.getName());
    }

    public void testGetTable() {
        assertNotNull(ws.getTable());
        Table table = ws.getTable();
        assertSame(table, ws.getTable());
        Table table3 = ws.getTable();
        assertSame(table, ws.getTable());
        assertSame(table3, ws.getTable());
    }

    public void testAddCell() {
        assertEquals(1, ws.getCellPointer().getRowIndex());
        assertEquals(1, ws.getCellPointer().getColumnIndex());
        Cell cell = ws.addCell();
        assertEquals(1, ws.getCellPointer().getRowIndex());
        assertEquals(2, ws.getCellPointer().getColumnIndex());
        assertNotNull(cell);
        Cell rCell = ws.getCellAt(1, 1);
        assertSame(cell, rCell);
        Cell rCell2 = ws.getTable().getRowAt(1).getCellAt(1);
        assertSame(cell, rCell2);
    }

    public void testAddCellRC() {
        Cell cell = ws.addCellAt(5, 7);
        assertEquals(5, ws.getCellPointer().getRowIndex());
        assertEquals(8, ws.getCellPointer().getColumnIndex());
        assertNotNull(cell);
        Cell rCell = ws.getCellAt(5, 7);
        assertSame(cell, rCell);
        Cell rCell2 = ws.getTable().getRowAt(5).getCellAt(7);
        assertSame(cell, rCell2);
        Cell rCell3 = ws.getRowAt(5).getCellAt(7);
        assertSame(cell, rCell3);
    }

    public void testAddCellCell() {
        Cell cell = new SSCell();
        ws.addCell(cell);
        assertEquals(1, ws.getCellPointer().getRowIndex());
        assertEquals(2, ws.getCellPointer().getColumnIndex());
        assertEquals(1, ws.getTable().rowCount());
        assertEquals(1, ws.getRowAt(1).size());
        Cell rCell = ws.getCellAt(1, 1);
        assertSame(cell, rCell);
        Cell rCell2 = ws.getTable().getRowAt(1).getCellAt(1);
        assertSame(cell, rCell2);
    }
    
    public void testAddCell_CellPointer() {
        ws.getCellPointer().moveTo(5, 60);
        Cell cell = ws.addCell();
        assertSame(cell, ws.getCellAt(5, 60));
        ws.getCellPointer().move(1, -60);
        Cell cell2 = ws.addCell();
        assertSame(cell2, ws.getCellAt(6, 1));
        ws.getCellPointer().moveTo(10, 256);
        Cell cell3 = ws.addCell();
        assertSame(cell3, ws.getCellAt(10, 256));
        Cell cell4 = null;
        try {
            cell4 = ws.addCell();
            fail("geen exceptie gegooid.");
        } catch (IndexOutOfBoundsException e) {
            assertEquals("columnIndex = 257", e.getMessage());
        }
        assertNull(cell4);
        assertEquals(256, ws.getTable().maxCellIndex());
        assertEquals(258, ws.getCellPointer().getColumnIndex());
    }

    public void testRemoveCell() {
        assertNull(ws.removeCellAt(5, 9));
        Cell cell = ws.addCellAt(5, 9);
        Cell cell2 = ws.removeCellAt(5, 9);
        assertSame(cell, cell2);
    }
    
    public void testAddColumn() {
        assertEquals(0, ws.getTable().maxColumnIndex());
        assertEquals(0, ws.getTable().columnCount());
        Column column = ws.addColumn();
        assertEquals(1, ws.getTable().maxColumnIndex());
        assertEquals(1, ws.getTable().columnCount());
        Column column2 = ws.getColumnAt(1);
        assertSame(column, column2);
        
        Column column3 = ws.addColumn();
        assertEquals(2, ws.getTable().maxColumnIndex());
        assertEquals(2, ws.getTable().columnCount());
        Column column4 = ws.getColumnAt(2);
        assertSame(column3, column4);
        assertNotSame(column4, column2);
    }
    
    public void testAddColumnint() {
        Column column = ws.addColumnAt(5);
        assertEquals(5, ws.getTable().maxColumnIndex());
        assertEquals(1, ws.getTable().columnCount());
        Column column5 = ws.getColumnAt(5);
        assertSame(column, column5);
        
        Column column6 = ws.addColumn();
        assertEquals(6, ws.getTable().maxColumnIndex());
        assertEquals(2, ws.getTable().columnCount());
        Column column6r = ws.getColumnAt(6);
        assertSame(column6, column6r);
        
        Column column300 = null;
        try {
            column300 = ws.addColumnAt(300);
            fail("geen exceptie gegooid.");
        } catch (IndexOutOfBoundsException e) {
            assertEquals("columnIndex = 300", e.getMessage());
        }
        assertNull(column300);
    }
    
    public void testColumnExceptions() {
        assertFalse(ws.hasColumnAt(""));
        assertFalse(ws.hasColumnAt("ALL"));
        assertNotNull(ws.getColumnAt("BQ"));
        assertNull(ws.removeColumnAt("ALL"));
        try {
            ws.getColumnAt("");
            fail("should be error");
        } catch (IndexOutOfBoundsException e) {
            assertEquals("columnIndex = 0", e.getMessage());
        }
    }
    
    public void testRowExceptions() {
        assertFalse(ws.hasRowAt(0));
        assertFalse(ws.hasRowAt(123456789));
        assertNotNull(ws.getRowAt(10));
        assertNull(ws.removeRowAt(-3));
        try {
            ws.getRowAt(66666666);
            fail("should be error");
        } catch (IndexOutOfBoundsException e) {
            assertEquals("rowIndex = 66666666", e.getMessage());
        }
    }
    
    public void testAddGetHasColumn() {
        assertFalse(ws.hasColumnAt(7));
        assertFalse(ws.hasTable());
        Column col1 = ws.getColumnAt(7);
        assertNotNull(col1);
        Column col2 = ws.addColumnAt(7);
        assertNotSame(col1, col2);
    }
    
    public void testAddGetHasRow() {
        assertFalse(ws.hasRowAt(7));
        assertFalse(ws.hasTable());
        Row row1 = ws.getRowAt(7);
        assertNotNull(row1);
        Row row2 = ws.addRowAt(7);
        assertNotSame(row1, row2);
    }
    
    public void testAddGetHasCell() {
        assertFalse(ws.hasCellAt(7, 3));
        assertFalse(ws.hasTable());
        Cell cell1 = ws.getCellAt(7, 3);
        assertNotNull(cell1);
        Cell cell2 = ws.addCellAt(7, 3);
        assertNotSame(cell1, cell2);
    }

    public void testAssemble() {
        ws.addRowAt(5);
        ws.setProtected(true);
        ws.setRightToLeft(true);
        WorksheetOptions wso = ws.getWorksheetOptions();
        wso.setGridlineColor(0, 0, 255);
        wso.setSelected(true);
        
        GIO gio = new GIO();
        String xml = xmlToString(ws, gio);
        
        assertTrue(xml.indexOf("<ss:Worksheet ss:Name=\"Sheet1\" ss:Protected=\"1\" ss:RightToLeft=\"1\">") > 0);
        assertTrue(xml.indexOf("<ss:Row ss:Index=\"5\"/>") > 0);
        assertTrue(xml.indexOf("<x:WorksheetOptions>") > 0);
        assertTrue(xml.indexOf("<x:Selected/>") > 0);
        assertTrue(xml.indexOf("<x:GridlineColor>#0000ff</x:GridlineColor>") > 0);
        assertEquals(1, gio.getSelectedSheetsCount());
        assertEquals(1, gio.getSelectedSheetsCount());
        
        //System.out.println(xml);
    }
}