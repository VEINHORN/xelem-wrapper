/*
 * Created on 3-dec-2004
 *
 */
package nl.fountain.xelem;

import java.util.ArrayList;
import java.util.Collection;

import junit.framework.TestCase;
import nl.fountain.xelem.excel.Worksheet;


public class AddressTest extends TestCase {

    public static void main(String[] args) {
        junit.textui.TestRunner.run(AddressTest.class);
    }
    
    public void testCalculateRow() {
        assertEquals(0, Address.calculateRow(""));
        assertEquals(1, Address.calculateRow("1"));
        assertEquals(10, Address.calculateRow("A10"));
        assertEquals(123, Address.calculateRow("a1@2^3n:"));
        assertEquals(Integer.MAX_VALUE, Address.calculateRow("2.147;483&647"));
        assertEquals(0, Address.calculateRow("JAVA"));
    }
    
    public void testCalculateColumn1() {
        assertEquals(0, Address.calculateColumn(""));
        assertEquals(1, Address.calculateColumn("A"));
        assertEquals(26, Address.calculateColumn("Z10"));
        assertEquals(26, Address.calculateColumn("123z"));
        assertEquals(27, Address.calculateColumn("AA520"));
        assertEquals(Integer.MAX_VALUE, Address.calculateColumn("FXSHRXW"));
        assertEquals(177009, Address.calculateColumn("JAVA"));
    }
    
    public void testCalculateColumn2() {
        for (int i = 0; i < 300; i++) {
            String column = Address.calculateColumn(i);
            assertEquals(i, Address.calculateColumn(column));
        }
        assertEquals("", Address.calculateColumn(-5));
        assertEquals("", Address.calculateColumn(0));
        assertEquals("JAVA", Address.calculateColumn(177009));
    }
    
    public void testStringConstructor() {
        assertTrue(new Address("BQ65").equals(new Address("65bq")));
        assertTrue(new Address("BQ65").equals(new Address("6Bq5")));
        
        Address adr = new Address("IV65536");
        assertEquals(65536, adr.getRowIndex());
        assertEquals(256, adr.getColumnIndex());
        assertEquals("IV65536", adr.getA1Reference());
        
        adr = new Address("xelem.2.0");
        assertEquals(20, adr.getRowIndex());
        assertEquals(11063559, adr.getColumnIndex());
        assertEquals("XELEM20", adr.getA1Reference());
    }

    public void testIsWithinSheet() {
        Address adr = new Address(Worksheet.firstRow, Worksheet.firstColumn);
        assertTrue(adr.isWithinSheet());
        adr = new Address(Worksheet.lastRow, Worksheet.lastColumn);
        assertTrue(adr.isWithinSheet());
        adr = new Address(Worksheet.firstRow -1, Worksheet.firstColumn);
        assertTrue(!adr.isWithinSheet());
        adr = new Address(Worksheet.firstRow, Worksheet.firstColumn - 1);
        assertTrue(!adr.isWithinSheet());
        adr = new Address(Worksheet.lastRow + 1, Worksheet.lastColumn);
        assertTrue(!adr.isWithinSheet());
        adr = new Address(Worksheet.lastRow, Worksheet.lastColumn + 1);
        assertTrue(!adr.isWithinSheet());
    }

    public void testGetAbsoluteAddress() {
        assertEquals("R5C7", new Address(5, 7).getAbsoluteAddress());
        assertEquals("R99999C-7", new Address(99999, -7).getAbsoluteAddress());
    }

    public void testGetAbsoluteRangeAddress() {
        Address adr = new Address(5, 3);
        assertEquals("R5C3", adr.getAbsoluteRange(new Address(5, 3)));
        assertEquals("R5C3:R8C10", adr.getAbsoluteRange(new Address(8, 10)));
        assertEquals("R2C1:R5C3", adr.getAbsoluteRange(new Address(2, 1)));
        assertEquals("R5C3:R5C10", adr.getAbsoluteRange(new Address(5, 10)));
        assertEquals("R3C3:R5C10", adr.getAbsoluteRange(new Address(3, 10)));
        
        CellPointer cp = new CellPointer();
        assertEquals("R1C1:R5C3", adr.getAbsoluteRange(cp));
    }

    public void testGetAbsoluteRangeintint() {
        Address adr = new Address(5, 3);
        assertEquals("R5C3", adr.getAbsoluteRange(5, 3));
        assertEquals("R5C3:R8C10", adr.getAbsoluteRange(8, 10));
        assertEquals("R2C1:R5C3", adr.getAbsoluteRange(2, 1));
        assertEquals("R5C3:R5C10", adr.getAbsoluteRange(5, 10));
        assertEquals("R3C3:R5C10", adr.getAbsoluteRange(3, 10));
    }

    public void testGetAbsoluteRangeCollection() {
        Collection<Address> list = new ArrayList<Address>();
        assertNull(Address.getAbsoluteRange(list));
        Address adr1 = new Address(5, 7);
        list.add(adr1);
        assertEquals("R5C7", Address.getAbsoluteRange(list));
        list.add(adr1);
        assertEquals("R5C7", Address.getAbsoluteRange(list));
        Address adr2 = new Address(5, 6);
        list.add(adr2);
        assertEquals("R5C6,R5C7", Address.getAbsoluteRange(list));
        Address adr3 = new Address(1, 256);
        list.add(adr3);
        assertEquals("R1C256,R5C6,R5C7", Address.getAbsoluteRange(list));
        
        CellPointer cp = new CellPointer();
        cp.moveTo(2, 10);
        list.add(cp);
        try {
            Address.getAbsoluteRange(list);
            fail("geen exceptie gegooid.");
        } catch (ClassCastException e) {
            //
        } 
        list.remove(cp);
        list.add(cp.getAddress());
        assertEquals("R1C256,R2C10,R5C6,R5C7", Address.getAbsoluteRange(list));
    }

    public void testGetRefToAddress() {
        Address adr = new Address(5, 7);
        Address adr2 = new Address(10, 10);
        assertEquals("R[-5]C[-3]", adr2.getRefTo(adr));
        assertEquals("R[5]C[3]", adr.getRefTo(adr2));
    }

    public void testGetRefTointint() {
        Address adr = new Address(5, 7);
        assertEquals("RC", adr.getRefTo(5, 7));
        assertEquals("R[-2]C", adr.getRefTo(3, 7));
        assertEquals("RC[-3]", adr.getRefTo(5, 4));
        assertEquals("R[-1]C[3]", adr.getRefTo(4, 10));
        assertEquals("R[65531]C[249]", adr.getRefTo(65536, 256));
    }

    public void testGetRefToAddressAddress() {
        Address adr = new Address(5, 7);
        Address adr1 = new Address(9, 4);
        Address adr2 = new Address(3, 8);
        assertEquals("R[4]C[-3]:R[-2]C[1]", adr.getRefTo(adr1, adr2));
    }


    public void testGetRefTointintintint() {
        Address adr = new Address(5, 7);
        assertEquals("RC", adr.getRefTo(5, 7, 5, 7));
        assertEquals("R[4]C[1]", adr.getRefTo(9, 8, 9, 8));
        assertEquals("RC:R[4]C[1]", adr.getRefTo(5, 7, 9, 8));
        assertEquals("R[-2]C[-3]:R[4]C[1]", adr.getRefTo(3, 4, 9, 8));
        assertEquals("R[4]C[-3]:R[-2]C[1]", adr.getRefTo(9, 4, 3, 8));
    }
    
    public void testGetRefToArea() {
        Address adr = new Address(5, 7);
        Area area = new Area(2, 3, 9, 8);
        assertEquals("R[-3]C[-4]:R[4]C[1]", adr.getRefTo(area));
        
        area = new Area(4, 6, 4, 6);
        assertEquals("R[-1]C[-1]", adr.getRefTo(area));
        area = new Area(5, 7, 5, 7);
        assertEquals("RC", adr.getRefTo(area));
    }

    /*
     * Class under test for String getRefTo(Collection)
     */
    public void testGetRefToCollection() {
        Address adr = new Address(5, 7);
        Collection<Address> list = new ArrayList<Address>();
        assertNull(adr.getRefTo(list));
        
        list.add(adr);
        assertEquals("RC", adr.getRefTo(list));
        list.add(new Address(7, 10));
        assertEquals("RC,R[2]C[3]", adr.getRefTo(list));
        list.add(new Address(7, 10));
        assertEquals("RC,R[2]C[3]", adr.getRefTo(list));
        CellPointer cp = new CellPointer();
        list.add(cp);
        try {
            adr.getRefTo(list);
            fail("geen exceptie gegooid.");
        } catch (ClassCastException e) {
            //
        }
        list.remove(cp);
        list.add(cp.getAddress());
        assertEquals("R[-4]C[-6],RC,R[2]C[3]", adr.getRefTo(list));
    }

    public void testToString() {
        Address adr = new Address(5, 3);
        assertEquals("nl.fountain.xelem.Address[row=5,column=3]", adr.toString());
    }

    public void testEquals() {
        Address adr = new Address(5, 3);
        Address adr2 = new Address(5, 3);
        assertTrue(adr.equals(adr));
        assertTrue(adr.equals(adr2));
        assertTrue(adr2.equals(adr));
        assertTrue(!adr.equals(null));
    }

    public void testCompareTo() {
        Address adr = new Address(5, 7);
        CellPointer cp = new CellPointer();
        try {
            adr.compareTo(cp);
            fail("geen exceptie gegooid.");
        } catch (ClassCastException e) {
            //
        }
        assertEquals(0, adr.compareTo(adr));
        assertEquals(0, adr.compareTo(new Address(5, 7)));
        assertTrue(adr.compareTo(new Address(5, 8)) < 0);
        assertTrue(adr.compareTo(new Address(5, 6)) > 0);
        assertTrue(adr.compareTo(new Address(6, 6)) < 0);
        assertTrue(adr.compareTo(new Address(4, 8)) > 0);
    }

}
