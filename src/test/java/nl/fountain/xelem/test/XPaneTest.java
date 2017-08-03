/*
 * Created on Nov 4, 2004
 *
 */
package nl.fountain.xelem.excel.x;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.excel.Pane;
import nl.fountain.xelem.excel.XLElementTest;

/**
 *
 */
public class XPaneTest extends XLElementTest {

    public static void main(String[] args) {
        junit.textui.TestRunner.run(XPaneTest.class);
    }
    
    public void testConstructor() {
        try {
            new XPane(4);
            fail("should throw exception.");
        } catch (IllegalArgumentException e) {
            assertEquals("4. Legal arguments are 0, 1, 2 and 3.", e.getMessage());
        }
        try {
            new XPane(-1);
            fail("should throw exception.");
        } catch (RuntimeException e1) {
            assertEquals("-1. Legal arguments are 0, 1, 2 and 3.", e1.getMessage());
        }
    }
    
    public void testAssemble() {
        Pane pane = new XPane(Pane.TOP_LEFT);
        pane.setActiveCell(2, 5);
        pane.setRangeSelection("R2C3:R2C10");
        
        String xml = xmlToString(pane, new GIO());
        assertTrue(xml.indexOf("<x:Pane>") > 0);
        assertTrue(xml.indexOf("<x:Number>3</x:Number>") > 0);
        assertTrue(xml.indexOf("<x:ActiveCol>4</x:ActiveCol>") > 0);
        assertTrue(xml.indexOf("<x:ActiveRow>1</x:ActiveRow>") > 0);
        assertTrue(xml.indexOf("<x:RangeSelection>R2C3:R2C10</x:RangeSelection>") > 0);
        
        //System.out.println(xml);
    }

}
