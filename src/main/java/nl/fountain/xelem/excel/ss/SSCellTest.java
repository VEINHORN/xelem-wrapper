/*
 * Created on Sep 8, 2004
 *
 */
package nl.fountain.xelem.excel.ss;

import java.math.BigDecimal;
import java.util.Date;
import java.util.Locale;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.XLUtil;
import nl.fountain.xelem.excel.Cell;
import nl.fountain.xelem.excel.XLElementTest;

/**
 *
 */
public class SSCellTest extends XLElementTest {

    private SSCell cell;

    public static void main(String[] args) {
        junit.textui.TestRunner.run(SSCellTest.class);
    }
    
    protected void setUp() {
        cell = new SSCell();
    }
    
    public void testDatatype() {
       assertEquals("String", cell.getXLDataType());
       assertEquals("", cell.getData$());
    }
    
    public void testDataTypeNumber() {
       cell.setData(new Integer(9));
       assertEquals("Number", cell.getXLDataType());
       assertEquals("9", cell.getData$());
       cell.setData(new Double(1.23456D));
       assertEquals("Number", cell.getXLDataType());
       assertEquals("1.23456", cell.getData$());
       cell.setData(new Long(123456789));
       assertEquals("Number", cell.getXLDataType());
       assertEquals("123456789", cell.getData$());
       cell.setData(new Float(123.45679F));
       assertEquals("Number", cell.getXLDataType());
       assertEquals("123.45679", cell.getData$());
       
       Object integer = new Integer(9);
       cell.setData(integer);
       assertEquals("Number", cell.getXLDataType());
       assertEquals("9", cell.getData$());
       
       Object dubbel = new Double(1.23456D);
       cell.setData(dubbel);
       assertEquals("Number", cell.getXLDataType());
       assertEquals("1.23456", cell.getData$());
       
       Object lang = new Long(123456789);
       cell.setData(lang);
       assertEquals("Number", cell.getXLDataType());
       assertEquals("123456789", cell.getData$());
       
       Object drijf = new Float(123.45679F);
       cell.setData(drijf);
       assertEquals("Number", cell.getXLDataType());
       assertEquals("123.45679", cell.getData$());
       
       cell.setData(8);
       assertEquals("Number", cell.getXLDataType());
       assertEquals("8", cell.getData$());
       cell.setData(0.0005d);
       assertEquals("Number", cell.getXLDataType());
       assertEquals("5.0E-4", cell.getData$());
       cell.setData(Long.MAX_VALUE);
       assertEquals("Number", cell.getXLDataType());
       assertEquals("9223372036854775807", cell.getData$());
       
       Number n = null;
       cell.setData(n);
       assertEquals("Error", cell.getXLDataType());
       assertEquals("#N/A", cell.getData$());
       assertEquals("=#N/A", cell.getFormula());
       
       n = new BigDecimal(123456.789);
       cell.setData(n);
       assertEquals("Number", cell.getXLDataType());
       assertEquals("123456.789", cell.getData$().substring(0, 10));
    }
    
    public void testDataTypeDate() {
       Date date = new Date(123456789L);
       assertEquals("1970-01-02T11:17:36.789", XLUtil.format(date));
       cell.setData(date);
       assertEquals("DateTime", cell.getXLDataType());
       assertEquals("1970-01-02T11:17:36.789", cell.getData$());
       
       Object datum = new Date(123456789L);
       assertEquals("1970-01-02T11:17:36.789", XLUtil.format((Date) datum));
       cell.setData(datum);
       assertEquals("DateTime", cell.getXLDataType());
       assertEquals("1970-01-02T11:17:36.789", cell.getData$());
       
       Date d = null;
       cell.setData(d);
       assertEquals("Error", cell.getXLDataType());
       assertEquals("#N/A", cell.getData$());
       assertEquals("=#N/A", cell.getFormula());
    }
    
    public void testDataTypeBoolean() {
       cell.setData(new Boolean(true));
       assertEquals("Boolean", cell.getXLDataType());
       assertEquals("1", cell.getData$());
       cell.setData(new Boolean(false));
       assertEquals("Boolean", cell.getXLDataType());
       assertEquals("0", cell.getData$());
       
       cell.setData(true);
       assertEquals("Boolean", cell.getXLDataType());
       assertEquals("1", cell.getData$());
       cell.setData(false);
       assertEquals("Boolean", cell.getXLDataType());
       assertEquals("0", cell.getData$());
       
       Object boo = new Boolean(true);
       cell.setData(boo);
       assertEquals("Boolean", cell.getXLDataType());
       assertEquals("1", cell.getData$());
       
       Boolean b = null;
       cell.setData(b);
       assertEquals("Error", cell.getXLDataType());
       assertEquals("#N/A", cell.getData$());
       assertEquals("=#N/A", cell.getFormula());
    }
    
    public void testDataTypeString() {
       cell.setData("bla bla");
       assertEquals("String", cell.getXLDataType());
       assertEquals("bla bla", cell.getData$());
       cell.setData(new Locale("nl"));
       assertEquals("String", cell.getXLDataType());
       assertEquals("nl", cell.getData$());
       cell.setData("\"heden\"");
       assertEquals("String", cell.getXLDataType());
       assertEquals("\"heden\"", cell.getData$());
       cell.setData("'heden' & v < g & g > v");
       assertEquals("String", cell.getXLDataType());
       assertEquals("'heden' & v < g & g > v", cell.getData$());
       
       cell.setData('c');
       assertEquals("String", cell.getXLDataType());
       assertEquals("c", cell.getData$());
       
       cell.setData('&');
       assertEquals("String", cell.getXLDataType());
       assertEquals("&", cell.getData$());
       
       String s = null;
       cell.setData(s);
       assertEquals("Error", cell.getXLDataType());
       assertEquals("#N/A", cell.getData$());
       assertEquals("=#N/A", cell.getFormula());
    }
    
    public void testDataTypeError() {
       cell.setError(Cell.ERRORVALUE_NA);
       assertEquals("Error", cell.getXLDataType());
       assertEquals("#N/A", cell.getData$());
       assertEquals("=#N/A", cell.getFormula());
    }
    
    public void testSetHRef() {
        cell.setHRef("http://www.microsoft.com/downloads/details.aspx?familyid=fe118952-3547-420a-a412-00a2662442d9&displaylang=en");
        assertEquals("http://www.microsoft.com/downloads/details.aspx?familyid=fe118952-3547-420a-a412-00a2662442d9&displaylang=en", cell.getHRef());
    }
    
    public void testSetStyleID() {
        cell.setStyleID("foo");
        assertEquals("foo", cell.getStyleID());
    }
    
    public void testInfinity() {
        cell.setData((double)25/0);
        assertEquals("Infinity", cell.getData$());
        assertEquals("String", cell.getXLDataType());
    }
    
    public void testAssembleStyle() {
        cell.setStyleID("foo");
        GIO gio = new GIO();
        String xml = xmlToString(cell, gio);
        //System.out.println(xml);
        assertTrue(xml.indexOf("<ss:Cell ss:StyleID=\"foo\"/>") > 0);
        assertEquals(1, gio.getStyleIDSet().size());
        assertEquals("foo", gio.getStyleIDSet().iterator().next());
    }
    
    public void testAssembleData() {
        cell.setData(false);
        GIO gio = new GIO();
        String xml = xmlToString(cell, gio);
        assertTrue(xml.indexOf("<Data ss:Type=\"Boolean\">0</Data>") > 0);
        assertEquals(0, gio.getStyleIDSet().size());
        
        //System.out.println(xml);
    }
    
//    public void testSubclassing() {
//        Object o = "123456789";
//        SSCell sCell = new SSCell() {
//            public void setData(String data) {
//                setType(DATATYPE_NUMBER);
//                setData(data.length());
//             }
//        };
//        sCell.setData(o);
//        GIO gio = new GIO();
//        String xml = xmlToString(sCell, gio);
//        assertTrue(xml.indexOf("<Data ss:Type=\"Number\">9</Data>") > 0);
//        
//        //System.out.println(xml);
//    }
    
    public void testGetData() {
        Cell cell = new SSCell();
        
        cell.setData(true);
        assertEquals(Boolean.class, cell.getData().getClass());
        Boolean boo = (Boolean) cell.getData();
        assertTrue(boo.booleanValue());
        
        cell.setData(new Boolean(false));
        assertEquals(Boolean.class, cell.getData().getClass());
        boo = (Boolean) cell.getData();
        assertTrue(!boo.booleanValue());
        
        byte b = 5;
        cell.setData(b);
        assertEquals(Double.class, cell.getData().getClass());
        Double doo = (Double) cell.getData();
        assertEquals(5, doo.intValue() );
        
        char c = 10;
        cell.setData(c);
        assertEquals(String.class, cell.getData().getClass());
        assertEquals("\n", cell.getData());
        
        cell.setData(new Date(123456789L));
        assertEquals(Date.class, cell.getData().getClass());
        Date date = (Date) cell.getData();
        assertEquals(123456000L, date.getTime());
        //System.out.println(new Date(123456789L));
        //System.out.println(cell.getData());
        
        BigDecimal big = 
            new BigDecimal(Double.MAX_VALUE).add(new BigDecimal(Double.MAX_VALUE));
        cell.setData(big);
        assertEquals("Infinity", cell.getData());
        doo = new Double(cell.getData().toString());
        assertTrue(doo.isInfinite());
        assertTrue(doo.doubleValue() > Double.MAX_VALUE);
        
        cell.setData(Double.MAX_VALUE);
        assertEquals(Double.class, cell.getData().getClass());
        doo = (Double) cell.getData();
        assertEquals(1.7976931348623157E308, doo.doubleValue(), 0.0D);
        
        Object obj = null;
        cell.setData(obj);
        assertEquals("#N/A", cell.getData());
        assertTrue(cell.hasError());
        
        cell.setData(cell);
        assertEquals(cell.toString(), cell.getData());
    }
    
    public void testIntValue() {
        Cell cell = new SSCell();
        assertEquals(0, cell.intValue());
        cell.setData(5.499999D);
        assertEquals(5, cell.intValue());
        Double doo = null;
        cell.setData(doo);
        assertEquals(0, cell.intValue());
        assertTrue(cell.hasError());
        cell.setData("123.456");
        assertEquals(123, cell.intValue());
    }
    
    public void testDoubleValue() {
        Cell cell = new SSCell();
        assertEquals(0.0D, cell.doubleValue(), 0.0D);
        cell.setData(Double.MIN_VALUE);
        assertEquals(4.9E-324D, cell.doubleValue(), 0.0D);
        cell.setData(Double.NaN);
        assertEquals("NaN", cell.doubleValue()+"");
        cell.setData(Double.NEGATIVE_INFINITY);
        assertEquals("-Infinity", cell.doubleValue()+"");
    }
    
    public void testBooleanValue() {
        Cell cell = new SSCell();
        assertFalse(cell.booleanValue());
        cell.setData(true);
        assertTrue(cell.booleanValue());
        cell.setData("1");
        assertTrue(cell.booleanValue());
        cell.setData(1);
        assertTrue(cell.booleanValue());
        cell.setData(0);
        assertFalse(cell.booleanValue());
        cell.setData("True");
        assertFalse(cell.booleanValue());
    }

}
