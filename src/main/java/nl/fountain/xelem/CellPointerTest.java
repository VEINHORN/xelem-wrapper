/*
 * Created on 28-nov-2004
 *
 */
package nl.fountain.xelem;

import junit.framework.TestCase;


public class CellPointerTest extends TestCase {

    public static void main(String[] args) {
        junit.textui.TestRunner.run(CellPointerTest.class);
    }
    
    public void testConstructor() {
        CellPointer cp = new CellPointer();
        assertEquals(1, cp.getColumnIndex());
        assertEquals(1, cp.getRowIndex());
    }
    
    public void testParameterlessMove() {
        CellPointer cp = new CellPointer();
        cp.move();
        assertEquals(2, cp.getColumnIndex());
        assertEquals(1, cp.getRowIndex());
    }
    
    public void testMoveToString() {
        CellPointer cp = new CellPointer();
        cp.moveTo("AA10");
        assertEquals(10, cp.getRowIndex());
        assertEquals(27, cp.getColumnIndex());
        cp.move(-2, -1);
        assertEquals("Z8", cp.getA1Reference());
        cp.move(7, 2);
        assertEquals("AB15", cp.getA1Reference());
        cp.moveTo("012point");
        assertEquals("POINT12", cp.getA1Reference());
    }
    
    public void testSetMovement() {
        CellPointer cp = new CellPointer();
        
        cp.setMovement(CellPointer.MOVE_DOWN);
        cp.move();
        assertEquals(1, cp.getColumnIndex());
        assertEquals(2, cp.getRowIndex());
        
        cp.setMovement(CellPointer.MOVE_LEFT);
        cp.move();
        assertEquals(0, cp.getColumnIndex());
        assertEquals(2, cp.getRowIndex());
        
        cp.setMovement(CellPointer.MOVE_RIGHT);
        cp.move();
        assertEquals(1, cp.getColumnIndex());
        assertEquals(2, cp.getRowIndex());
        
        cp.setMovement(CellPointer.MOVE_UP);
        cp.move();
        assertEquals(1, cp.getColumnIndex());
        assertEquals(1, cp.getRowIndex());
        
        try {
            cp.setMovement(-1);
            fail("geen exceptie gegooid.");
        } catch (IllegalArgumentException e) {
            assertEquals("-1. Legal values are 0, 1, 2 and 3.", e.getMessage());
        }
    } 
    
    public void testStepDistance() {
        CellPointer cp = new CellPointer();
        
        cp.setHorizontalStepDistance(10);
        cp.move();
        assertEquals(11, cp.getColumnIndex());
        assertEquals(1, cp.getRowIndex());
        
        cp.setVerticalStepDistance(5);
        cp.setMovement(CellPointer.MOVE_DOWN);
        cp.move();
        assertEquals(11, cp.getColumnIndex());
        assertEquals(6, cp.getRowIndex());
        
        cp.setMovement(CellPointer.MOVE_LEFT);
        cp.move();
        assertEquals(1, cp.getColumnIndex());
        assertEquals(6, cp.getRowIndex());
        
        cp.setMovement(CellPointer.MOVE_UP);
        cp.move();
        assertEquals(1, cp.getColumnIndex());
        assertEquals(1, cp.getRowIndex());
    }
    
    public void testIsWithinSheet() {
        CellPointer cp = new CellPointer();
        assertTrue(cp.isWithinSheet());
        cp.moveTo(0, 1);
        assertTrue(!cp.isWithinSheet());
        cp.moveTo(1, 0);
        assertTrue(!cp.isWithinSheet());
        cp.moveTo(65537, 1);
        assertTrue(!cp.isWithinSheet());
        cp.moveTo(1, 257);
        assertTrue(!cp.isWithinSheet());
        cp.moveTo(1000, 10);
        assertTrue(cp.isWithinSheet());
        
    }
    
    public void testGetAbsoluteAddress() {
        CellPointer cp = new CellPointer();
        assertEquals("R1C1", cp.getAbsoluteAddress());
        cp.moveTo(25, 107);
        assertEquals("R25C107", cp.getAbsoluteAddress());
    }
    
    public void testToString() {
        CellPointer cp = new CellPointer();
        assertEquals("nl.fountain.xelem.CellPointer[row=1,column=1]", cp.toString());
    }
    
    public void testEquals() {
        CellPointer cp = new CellPointer();
        Address adr = new Address(5, 7);
        assertTrue(!adr.equals(cp));
        assertTrue(!cp.equals(adr));
        
        cp.moveTo(5, 7);
        assertTrue(!adr.equals(cp));
        assertTrue(!cp.equals(adr));
        
        Object cpo = new CellPointer();
        assertTrue(!adr.equals(cpo));
        assertTrue(!cpo.equals(adr));
        
        Address adrn = new Address(1,1);
        assertTrue(!adrn.equals(cpo));
        assertTrue(!cpo.equals(adrn));
    }
    
    public void testCompareTo() {
        CellPointer cp = new CellPointer();
        try {
            cp.compareTo(null);
            fail("geen exceptie gegooid");
        } catch (NullPointerException e) {
            //
        }
        
        Address adr = new Address(5, 7);
        try {
            cp.compareTo(adr);
            fail("geen exceptie gegooid");
        } catch (ClassCastException e1) {
            //
        }
        
        CellPointer cp2 = new CellPointer();
        assertEquals(0, cp.compareTo(cp2));
        
        cp2.moveTo(1, 2);
        assertTrue(cp.compareTo(cp2) < 0);
        
        cp2.moveTo(2, 1);
        assertTrue(cp.compareTo(cp2) < 0);
        
        cp2.moveTo(1, 257);
        assertTrue(cp.compareTo(cp2) < 0);
 
        cp.moveTo(2, 1);
        assertTrue(cp.compareTo(cp2) > 0);
    }

}
