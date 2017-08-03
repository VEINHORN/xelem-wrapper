/*
 * Created on 5-apr-2005
 *
 */
package nl.fountain.xelem;

import junit.framework.TestCase;

/**
 *
 */
public class AreaTest extends TestCase {

    public static void main(String[] args) {
        junit.textui.TestRunner.run(AreaTest.class);
    }
    
    public void testIntConstructor() {
        Area area = new Area(1, 2, 3, 4);
        assertEquals(1, area.getFirstRow());
        assertEquals(2, area.getFirstColumn());
        assertEquals(3, area.getLastRow());
        assertEquals(4, area.getLastColumn());
        
        area = new Area(4, 3, 2, 1);
        assertEquals(2, area.getFirstRow());
        assertEquals(1, area.getFirstColumn());
        assertEquals(4, area.getLastRow());
        assertEquals(3, area.getLastColumn());
    }
    
    public void testAddressConstructor() {
        Address address1 = new Address(2, 3);
        Address address2 = new Address(5, 8);
        Area area = new Area(address1, address2);
        assertEquals(2, area.getFirstRow());
        assertEquals(3, area.getFirstColumn());
        assertEquals(5, area.getLastRow());
        assertEquals(8, area.getLastColumn());
    }
    
    public void testStringConstructor() {
        Area area = new Area("C2:H5");
        assertEquals(2, area.getFirstRow());
        assertEquals(3, area.getFirstColumn());
        assertEquals(5, area.getLastRow());
        assertEquals(8, area.getLastColumn());
        
        area = new Area("H5:C2");
        assertEquals(2, area.getFirstRow());
        assertEquals(3, area.getFirstColumn());
        assertEquals(5, area.getLastRow());
        assertEquals(8, area.getLastColumn());
        
        area = new Area("foo:bar");
        assertEquals(0, area.getFirstRow());
        assertEquals(1396, area.getFirstColumn());
        assertEquals(0, area.getLastRow());
        assertEquals(4461, area.getLastColumn());
        
        area = new Area("1:2");
        assertEquals(1, area.getFirstRow());
        assertEquals(0, area.getFirstColumn());
        assertEquals(2, area.getLastRow());
        assertEquals(0, area.getLastColumn());
        
        try {
            area = new Area(":");
            fail("geen exceptie gegooid.");
        } catch (IllegalArgumentException e1) {
            assertEquals("use format 'A1:A1'.", e1.getMessage());
        }
        
        try {
            area = new Area("foobar");
            fail("geen exceptie gegooid.");
        } catch (IllegalArgumentException e) {
            assertEquals("use format 'A1:A1'.", e.getMessage());
        }
    }
    
    public void testGetA1Reference() {
        Area area = new Area("bz50:ab60");
        assertEquals("AB50:BZ60", area.getA1Reference());
    }
    
    public void testGetAbsoluteRange() {
        Area area = new Area(2, 3, 5, 8);
        assertEquals("R2C3:R5C8", area.getAbsoluteRange());
    }
    
    public void testIsWithinArea() {
        Area area = new Area(2, 3, 5, 8);
        assertTrue(area.isWithinArea(2, 3));
        assertTrue(area.isWithinArea(5, 3));
        assertTrue(area.isWithinArea(2, 8));
        assertTrue(area.isWithinArea(5, 8));
        assertTrue(area.isWithinArea(3, 7));
        
        assertFalse(area.isWithinArea(1, 7));
    }

}
