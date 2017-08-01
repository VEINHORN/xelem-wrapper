/*
 * Created on 24-okt-2004
 *
 */
package nl.fountain.xelem;

import java.io.FileNotFoundException;
import java.util.Iterator;
import java.util.List;

import javax.xml.parsers.ParserConfigurationException;

import nl.fountain.xelem.excel.XLElement;
import nl.fountain.xelem.excel.XLElementTest;

import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;


public class XFactoryTest extends XLElementTest {

    public static void main(String[] args) {
        junit.textui.TestRunner.run(XFactoryTest.class);
    }
    
    protected void setUp() throws Exception {
        String configFileName =
            "testsuitefiles/XFactoryTest/XFactoryTest.xml";
        XFactory.setConfigurationFileName(configFileName);
    }
    
    protected void tearDown() throws Exception {
        XFactory.reset();
    }
    
    public void testNewInstance() throws XelemException {
        XFactory x = XFactory.newInstance();
        assertNotNull(x);
        XFactory x2 = XFactory.newInstance();
        assertNotSame(x, x2);
        XFactory.reset();
        assertNotNull(x);
        assertNotSame(x, x2);
        XFactory x3 = XFactory.newInstance();
        assertNotSame(x, x3);
    }
    
    public void testEmptyInstance() throws XelemException {
        XFactory x = XFactory.newInstance();
        assertEquals(3, x.getDocComments().size());
        assertEquals(4, x.getStylesCount());
        
        XFactory ef = XFactory.emptyFactory();
        assertEquals(0, ef.getDocComments().size());
        assertEquals(0, ef.getStylesCount());
        assertEquals(0, x.getDocComments().size());
        assertEquals(0, x.getStylesCount());
        assertEquals("config/xelem.xml", XFactory.getConfigurationFileName());
    }
    
    public void testReset() throws XelemException {
        XFactory x = XFactory.newInstance();
        assertEquals(3, x.getDocComments().size());
        assertEquals(4, x.getStylesCount());
        assertNotNull(XFactory.getConfigurationFileName());
        assertTrue(!"config/xelem.xml".equals(
                XFactory.getConfigurationFileName()));
        
        XFactory.reset();
        assertEquals(3, x.getDocComments().size());
        assertEquals(4, x.getStylesCount());
        assertEquals("config/xelem.xml", XFactory.getConfigurationFileName());
    }
    
    public void testGetDocComments() throws XelemException {
       List<String> docComments = XFactory.newInstance().getDocComments(); 
       assertNotNull(docComments);
       Iterator<String> iter = docComments.iterator();
       assertEquals("      f:comment 1      ", iter.next());
       assertEquals("      f:comment 2      ", iter.next());
       assertEquals("      f:comment 3      ", iter.next());      
    }
    
    public void testGetStyles() throws XelemException {
        XFactory x = XFactory.newInstance();
        assertNotNull(x.getStyle("decimal1"));
        assertNotNull(x.getStyle("decimal2"));
        assertNotNull(x.getStyle("bold"));
        assertNotNull(x.getStyle("b_yellow"));
    }
    
    public void testException() {
        XFactory.setConfigurationFileName("this/file/does/not/exist");
        try {
            XFactory.newInstance();
            fail("a file with a very inordinate name does exist.");
        } catch (XelemException e) {
            assertEquals(FileNotFoundException.class,
                    e.getCause().getClass());
            
//            System.out.println(e.toString());
//            StackTraceElement[] st = e.getStackTrace();
//            for (int i = 0; i < st.length; i++) {
//                System.out.println("\tat " + st[i].toString());
//            }
            
        }
    }
    
    public void testMergeStylesExc() throws XelemException {
        XFactory x = XFactory.newInstance();
        try {
            x.mergeStyles("noID", "foo", "bar");
            fail("geen Exceptie gegooid");
        } catch (UnsupportedStyleException e) {
            assertEquals("Style(s) 'foo' 'bar' not found.", e.getMessage());
        }
        assertNull(x.getStyle("noID"));
    }
    
    public void testMergeStyles() throws Exception {
        XFactory x = XFactory.newInstance();
        x.mergeStyles("nieuw", "bold", "b_yellow");
        assertNotNull(x.getStyle("nieuw"));
        Element nieuw = x.getStyle("nieuw");
        String xml = xmlToString(nieuw);       
        //System.out.println(xml);
        
        assertTrue(xml.indexOf("<Style ss:ID=\"nieuw\">") > 0);
        assertTrue(xml.indexOf("x:Family=\"Swiss\"") > 0);
        assertTrue(xml.indexOf("ss:Bold=\"1\"") > 0);
        
        assertTrue(xml.indexOf("<Interior ss:Color=\"#FFFF00\" " +
        		"ss:Pattern=\"Solid\"/>") > 0);
        
        x.mergeStyles("nieuwer", "nieuw", "decimal2");
        assertNotNull(x.getStyle("nieuwer"));
        assertEquals(6, x.getStylesCount());
        
        xml = xmlToString(x.getStyle("nieuwer"));
        //System.out.println(xml);
        
        assertTrue(xml.indexOf("<Style ss:ID=\"nieuwer\">") > 0);
        assertTrue(xml.indexOf("x:Family=\"Swiss\"") > 0);
        assertTrue(xml.indexOf("ss:Bold=\"1\"") > 0);
        assertTrue(xml.indexOf("<Interior ss:Color=\"#FFFF00\" " +
        		"ss:Pattern=\"Solid\"/>") > 0);
        assertTrue(xml.indexOf("<NumberFormat ss:Format=\"0.00\"/>") > 0);
        
        x.mergeStyles("nieuwer", null, null);
        assertEquals(6, x.getStylesCount());
        
        x.mergeStyles("stay0.00", "nieuwer", "decimal1");
        assertNotNull(x.getStyle("stay0.00"));
        assertEquals(7, x.getStylesCount());
        
        xml = xmlToString(x.getStyle("stay0.00"));
        //System.out.println(xml);
        
        assertTrue(xml.indexOf("<NumberFormat ss:Format=\"0.00\"/>") > 0);
        assertEquals(-1, xml.indexOf("<NumberFormat ss:Format=\"0.0\"/>"));
    }
    
    public void testAddStyle() throws XelemException {
        XFactory x = XFactory.newInstance();
        Element style = x.getStyle("b_yellow");
        assertTrue(!x.addStyle(style));
        Element nStyle = style.getOwnerDocument().createElement("Style");
        Attr attr = style.getOwnerDocument().createAttributeNS(XLElement.XMLNS_SS, "ID");
        attr.setPrefix(XLElement.PREFIX_SS);
        attr.setNodeValue("nieuw");
        nStyle.setAttributeNodeNS(attr);
        assertTrue(x.addStyle(nStyle));
        
        Element noStyle = style.getOwnerDocument().createElement("Style");
        attr = style.getOwnerDocument().createAttributeNS(XLElement.XMLNS_SS, "FOO");
        attr.setPrefix(XLElement.PREFIX_SS);
        attr.setNodeValue("no_id");
        noStyle.setAttributeNodeNS(attr);
        try {
            x.addStyle(noStyle);
            fail("should have thrown Exception.");
        } catch (NullPointerException e) {
            //
        }
    }
    
    public void testGetInfoSheet() throws XelemException {
        XFactory x = XFactory.newInstance();
        Node infoSheet = x.loadInfoSheet();
        
        String xml = xmlToString(infoSheet);
        assertTrue(xml.indexOf("<ss:Worksheet ss:Name=") > 0);
        //System.out.println(xml);
    }
    
    public void testAppendInfoSheet() throws ParserConfigurationException, XelemException {
        Document doc = getDoc();
        XFactory x = XFactory.newInstance();
        GIO gio = new GIO();
        
        x.appendInfoSheet(doc.getDocumentElement(), gio);
        // depends on whether the infoSheet.xml will declare styles
        assertTrue(gio.getStyleIDSet().size() > 0);
        assertTrue(!gio.getStyleIDSet().contains("Default"));
        
        String xml = xmlToString(doc.getElementsByTagNameNS(
                XLElement.XMLNS_SS, "Worksheet").item(0));
        assertTrue(xml.indexOf("<ss:Worksheet ss:Name=") > 0);
        //System.out.println(xml);
    }

}
