/*
 * Created on Nov 2, 2004
 *
 */
package nl.fountain.xelem.excel.ss;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.excel.XLElementTest;

/**
 *
 */
public class SSColumnTest extends XLElementTest {
    
    private SSColumn column;

    public static void main(String[] args) {
        junit.textui.TestRunner.run(SSColumnTest.class);
    }
    
    protected void setUp() {
        column = new SSColumn();
    }
    
    public void testAssemble() {
        column.setStyleID("abc");
        GIO gio = new GIO();
        String xml = xmlToString(column, gio);
        
        assertTrue(xml.indexOf("<ss:Column ss:StyleID=\"abc\"") > 0);
        assertEquals(1, gio.getStyleIDSet().size());
        assertEquals("abc", gio.getStyleIDSet().iterator().next());
        
        //System.out.println(xml);
    }
    
    public void testAssemble2() {
        column.setHidden(true);
        column.setSpan(5);
        column.setStyleID("foo");
        column.setWidth(2.3);
        GIO gio = new GIO();
        String xml = xmlToString(column, gio);
        //System.out.println(xml);
        assertTrue(xml.indexOf("<ss:Column") > 0);
        assertTrue(xml.indexOf("ss:StyleID=\"foo\"") > 0);
        assertTrue(xml.indexOf("ss:Span=\"5\"") > 0);
        assertTrue(xml.indexOf("ss:Width=\"2.3\"") > 0);
        assertTrue(xml.indexOf("ss:Hidden=\"1\"") > 0);       
    }

}
