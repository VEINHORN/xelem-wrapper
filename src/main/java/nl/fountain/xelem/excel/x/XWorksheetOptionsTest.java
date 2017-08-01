/*
 * Created on Nov 2, 2004
 *
 */
package nl.fountain.xelem.excel.x;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.excel.Pane;
import nl.fountain.xelem.excel.WorksheetOptions;
import nl.fountain.xelem.excel.XLElementTest;

/**
 *
 */
public class XWorksheetOptionsTest extends XLElementTest {
    
    private XWorksheetOptions wso;
    
    public static void main(String[] args) {
        junit.textui.TestRunner.run(XWorksheetOptionsTest.class);
    }
    
    protected void setUp() {
        wso = new XWorksheetOptions();
    }
    
    public void testSetVisible() {
        try {
            wso.setVisible("foo");
            fail("should throw exception");
        } catch (IllegalArgumentException e) {
            assertEquals("foo. Should be one of WorksheetOptions.SHEET_xxx values.",
                    e.getMessage());
        }
    }
    
    public void testSetActiveCell() {
        try {
            wso.setActiveCell(4, 2, 3);
            fail("should throw exception.");
        } catch (IllegalArgumentException e) {
            assertEquals("4. Legal arguments are 0, 1, 2 and 3.", e.getMessage());
        }
        try {
            wso.setActiveCell(0, 3);
            fail("should throw exception.");
        } catch (IllegalArgumentException e) {
            assertEquals("0. Can't be less than 1.", e.getMessage());
        }
        try {
            wso.setActiveCell(5, -2);
            fail("should throw exception.");
        } catch (IllegalArgumentException e) {
            assertEquals("-2. Can't be less than 1.", e.getMessage());
        }
    }
    
    public void testSplitHorizontal() {
        try {
            wso.splitHorizontal(5000, 0);
            fail("should throw exception.");
        } catch (IllegalArgumentException e) {
            assertEquals("0. Can't be less than 1.", e.getMessage());
        }
        wso.splitHorizontal(5000, 8);
        wso.setActiveCell(Pane.BOTTOM_LEFT, 2, 3);
        String xml = xmlToString(wso, new GIO());
        assertTrue(xml.indexOf("<x:SplitHorizontal>5000</x:SplitHorizontal>") > 0);
        assertTrue(xml.indexOf("<x:TopRowBottomPane>7</x:TopRowBottomPane>") > 0);
        assertTrue(xml.indexOf("<x:ActivePane>2</x:ActivePane>") > 0);
        assertTrue(xml.indexOf("<x:Number>2</x:Number>") > 0);
        assertTrue(xml.indexOf("<x:ActiveCol>2</x:ActiveCol>") > 0);
        assertTrue(xml.indexOf("<x:ActiveRow>1</x:ActiveRow>") > 0);
        
        //System.out.println(xml);
    }
    
    public void testSplitVertical() {
        try {
            wso.splitVertical(5000, 0);
            fail("should throw exception.");
        } catch (IllegalArgumentException e) {
            assertEquals("0. Can't be less than 1.", e.getMessage());
        }
        wso.splitVertical(5000, 8);
        wso.setActiveCell(Pane.TOP_RIGHT, 2, 3);
        String xml = xmlToString(wso, new GIO());
        assertTrue(xml.indexOf("<x:SplitVertical>5000</x:SplitVertical>") > 0);
        assertTrue(xml.indexOf("<x:LeftColumnRightPane>7</x:LeftColumnRightPane>") > 0);
        assertTrue(xml.indexOf("<x:ActivePane>1</x:ActivePane>") > 0);
        assertTrue(xml.indexOf("<x:Number>1</x:Number>") > 0);
        assertTrue(xml.indexOf("<x:ActiveCol>2</x:ActiveCol>") > 0);
        assertTrue(xml.indexOf("<x:ActiveRow>1</x:ActiveRow>") > 0);
        
        //System.out.println(xml);
    }
    
    public void testAssemble() {
        wso.doDisplayFormulas(true);
        wso.doNotDisplayGridlines(true);
        wso.doNotDisplayHeadings(true);
        wso.setGridlineColor(255, 0, 100);
        wso.setLeftColumnVisible(6);
        wso.setSelected(true);
        wso.setTabColorIndex(123456);
        wso.setTopRowVisible(7);
        wso.setVisible(WorksheetOptions.SHEET_HIDDEN);
        wso.setZoom(150);
        GIO gio = new GIO();
        String xml = xmlToString(wso, gio);
        
        assertTrue(xml.indexOf("<x:WorksheetOptions>") > 0);
        assertTrue(xml.indexOf("<x:Selected/>") > 0);
        assertTrue(xml.indexOf("<x:DoNotDisplayHeadings/>") > 0);
        assertTrue(xml.indexOf("<x:DoNotDisplayGridlines/>") > 0);
        assertTrue(xml.indexOf("<x:DisplayFormulas/>") > 0);
        assertTrue(xml.indexOf("<x:TopRowVisible>7</x:TopRowVisible>") > 0);
        assertTrue(xml.indexOf("<x:LeftColumnVisible>6</x:LeftColumnVisible>") > 0);
        assertTrue(xml.indexOf("<x:Zoom>150</x:Zoom>") > 0);
        assertTrue(xml.indexOf("<x:TabColorIndex>123456</x:TabColorIndex>") > 0);
        assertTrue(xml.indexOf("<x:GridlineColor>#ff0064</x:GridlineColor>") > 0);
        assertTrue(xml.indexOf("<x:Visible>SheetHidden</x:Visible>") > 0);
        
        //System.out.println(xml);
    }
    
    public void testPanesAssembly() {
        wso.setActiveCell(Pane.BOTTOM_LEFT, 2, 3);
        GIO gio = new GIO();
        String xml = xmlToString(wso, gio);
        assertEquals(-1, xml.indexOf("<x:ActivePane>"));
        assertTrue(xml.indexOf("<x:Panes/>") > 0);
        
        wso.setActiveCell(Pane.TOP_RIGHT, 2, 3);
        xml = xmlToString(wso, gio);
        assertEquals(-1, xml.indexOf("<x:ActivePane>"));
        assertTrue(xml.indexOf("<x:Panes/>") > 0);
        
        wso.setActiveCell(Pane.BOTTOM_RIGHT, 2, 3);
        xml = xmlToString(wso, gio);
        assertEquals(-1, xml.indexOf("<x:ActivePane>"));
        assertTrue(xml.indexOf("<x:Panes/>") > 0);
        
        wso.setActiveCell(Pane.TOP_LEFT, 2, 3);
        xml = xmlToString(wso, gio);
        assertTrue(xml.indexOf("<x:ActivePane>3</x:ActivePane>") > 0);
        assertTrue(xml.indexOf("<x:Number>3</x:Number>") > 0);
        assertTrue(xml.indexOf("<x:ActiveCol>2</x:ActiveCol>") > 0);
        assertTrue(xml.indexOf("<x:ActiveRow>1</x:ActiveRow>") > 0);
        
        //System.out.println(xml);
    }
    
    public void testSplitVertical2() {
        wso.splitVertical(5000, 8);
        
        wso.setActiveCell(Pane.BOTTOM_LEFT, 2, 3);
        String xml = xmlToString(wso, new GIO());
        assertEquals(-1, xml.indexOf("<x:ActivePane>"));
        assertTrue(xml.indexOf("<x:Panes/>") > 0);
        
        wso.setActiveCell(Pane.BOTTOM_RIGHT, 2, 3);
        xml = xmlToString(wso, new GIO());
        assertEquals(-1, xml.indexOf("<x:ActivePane>"));
        assertTrue(xml.indexOf("<x:Panes/>") > 0);
        
        //System.out.println(xml);
    }
    
    public void testSplitHorizontal2() {
        wso.splitHorizontal(5000, 8);
        
        wso.setActiveCell(Pane.TOP_RIGHT, 2, 3);
        String xml = xmlToString(wso, new GIO());
        assertEquals(-1, xml.indexOf("<x:ActivePane>"));
        assertTrue(xml.indexOf("<x:Panes/>") > 0);
        
        wso.setActiveCell(Pane.BOTTOM_RIGHT, 2, 3);
        xml = xmlToString(wso, new GIO());
        assertEquals(-1, xml.indexOf("<x:ActivePane>"));
        assertTrue(xml.indexOf("<x:Panes/>") > 0);
        
        //System.out.println(xml);
    }
    
    public void testFreezePanesAt() {
        wso.freezePanesAt(2, 3);
                
        String xml = xmlToString(wso, new GIO());
        assertTrue(xml.indexOf("<x:FreezePanes/>") > 0);
        assertTrue(xml.indexOf("<x:FrozenNoSplit/>") > 0);
        assertTrue(xml.indexOf("<x:SplitHorizontal>2</x:SplitHorizontal>") > 0);
        assertTrue(xml.indexOf("<x:TopRowBottomPane>2</x:TopRowBottomPane>") > 0);
        assertTrue(xml.indexOf("<x:SplitVertical>3</x:SplitVertical>") > 0);
        assertTrue(xml.indexOf("<x:LeftColumnRightPane>3</x:LeftColumnRightPane>") > 0);
        assertTrue(xml.indexOf("<x:ActivePane>0</x:ActivePane>") > 0);
        
        //System.out.println(xml);
    }
    
    

}
