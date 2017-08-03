/*
 * Created on Oct 29, 2004
 *
 */
package nl.fountain.xelem;

import java.util.Date;

import junit.framework.TestCase;

/**
 *
 */
public class XLUtilTest extends TestCase {

    public static void main(String[] args) {
        junit.textui.TestRunner.run(XLUtilTest.class);
    }
    
    public void testConvertToHex() {
        assertEquals("#000000", XLUtil.convertToHex(0, 0, 0));
        assertEquals("#000000", XLUtil.convertToHex(-1, -2, -3));
        assertEquals("#ffffff", XLUtil.convertToHex(255, 255, 255));
        assertEquals("#ffffff", XLUtil.convertToHex(256, 257, 258));
    }
    
    public void testFormat() {
        Date date = new Date(1234567890123L);
        assertEquals("2009-02-14T00:31:30.123", XLUtil.format(date));
    }

}
