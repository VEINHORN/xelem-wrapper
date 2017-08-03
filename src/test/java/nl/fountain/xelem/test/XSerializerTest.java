/*
 * Created on 9-nov-2004
 *
 */
package nl.fountain.xelem;

import junit.framework.TestCase;
import nl.fountain.xelem.excel.Workbook;
import nl.fountain.xelem.excel.ss.XLWorkbook;


public class XSerializerTest extends TestCase {
    
    private Workbook wb;

    public static void main(String[] args) {
        junit.textui.TestRunner.run(XSerializerTest.class);
    }
    
    protected void setUp() throws Exception {
        wb = new XLWorkbook("test");
    }
    
//    public void testToString() throws XelemException {
//        XSerializer xs = new XSerializer();
//        String xml = xs.serializeToString(wb);
//        assertEquals("<?xml version=\"1.0\" encoding=\"UTF-8\"?>", 
//                xml.substring(0, 38));
//        System.out.println(xml);
//    }
    
//    public void testStream() throws XelemException {
//        XSerializer xs = new XSerializer();
//        xs.serialize(wb, System.out);
//    }
    
    public void testFile() throws XelemException {
        XSerializer xs = new XSerializer();
        wb.setFileName("testoutput/xs.xml");
        wb.addSheet().addCell("BV Financiën");
        xs.serialize(wb);
    }
    
    public void testFile2() throws XelemException {
        XSerializer xs = new XSerializer(XSerializer.US_ASCII);
        wb.setFileName("testoutput/xs_US_ASCII.xml");
        wb.addSheet().addCell("BV Financiën");
        xs.serialize(wb);
    }

}
