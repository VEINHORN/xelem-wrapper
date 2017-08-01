/*
 * Created on Nov 2, 2004
 *
 */
package nl.fountain.xelem.excel.ss;

import java.io.File;
import java.io.StringWriter;
import java.math.BigDecimal;
import java.util.Date;
import java.util.Iterator;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import junit.framework.TestCase;
import nl.fountain.xelem.XFactory;
import nl.fountain.xelem.XelemException;
import nl.fountain.xelem.excel.Cell;
import nl.fountain.xelem.excel.Column;
import nl.fountain.xelem.excel.Comment;
import nl.fountain.xelem.excel.DocumentProperties;
import nl.fountain.xelem.excel.ExcelWorkbook;
import nl.fountain.xelem.excel.Pane;
import nl.fountain.xelem.excel.Row;
import nl.fountain.xelem.excel.Table;
import nl.fountain.xelem.excel.Workbook;
import nl.fountain.xelem.excel.Worksheet;
import nl.fountain.xelem.excel.WorksheetOptions;

import org.w3c.dom.Document;


// Assertions *are* tested in these tests -...- but
// the proof of the pudding is in the eating....
// 
public class CreateDocumentTest extends TestCase {
    
    // the path to the directory for test files.
    private String testOutputDir = "testoutput/CreateDocumentTest/";
    
    // when set to true, test files will be created.
    // the path mentioned after 'testFileDir' should exist.
    private boolean toFile = true;
    
    private int warnings;
    private boolean printWarnings;

    public static void main(String[] args) {
        junit.textui.TestRunner.run(CreateDocumentTest.class);
    }
    
    // @see junit.framework.TestCase#setUp()
    protected void setUp() throws Exception {
        String configFileName =
            "testsuitefiles/CreateDocumentTest/CreateDocumentTest.xml";
        XFactory.setConfigurationFileName(configFileName);
        warnings = 0;
        printWarnings = true;
    }
    
    public void testAdvise() {
       if (toFile) {
           System.out.println();
           System.out.println(this.getClass() + " is writing files to: "
                   + testOutputDir);
       }
    }
    
    public void testConfigurationFile() throws XelemException {
        assertTrue(!"config/xelem.xml".equals(XFactory.getConfigurationFileName()));
        XFactory xFactory = XFactory.newInstance();
        assertNotNull(xFactory.getStyle("Default"));
        assertNotNull(xFactory.getStyle("bold"));
        assertNotNull(xFactory.getStyle("b_yellow"));
        assertNotNull(xFactory.getStyle("b_lblue"));
        assertNotNull(xFactory.getStyle("gray_date"));
        assertNotNull(xFactory.getStyle("currency"));
    }
    
    public void testConfigurationFile2() throws Exception {
        XFactory.reset();
        XFactory.setConfigurationFileName("foo");
        Workbook wb = new XLWorkbook("test_config");
        wb.addSheet().addCell().setStyleID("bar");
        
        warnings = 2;
        printWarnings = false;
        String xml = xmlToString(wb);
        Iterator<String> iter = wb.getWarnings().iterator();
        String w1 = iter.next();
        String w2 = iter.next();
        assertTrue(w1.indexOf("WARNING 1): java.io.FileNotFoundException:") > 0);
        assertTrue(w2.indexOf("WARNING 2): nl.fountain.xelem." +
        		"UnsupportedStyleException: Style 'bar' not found.") > 0);
        assertTrue(xml.indexOf("<Style ss:ID=\"bar\"/>") > 0);
        
        //System.out.println(xml);
        XFactory.reset();
    }
    
    public void testWorkbook() throws Exception {
        Workbook wb = new XLWorkbook("test00");
        
        String xml = xmlToString(wb);
        //System.out.println(xml);
        assertTrue(xml.indexOf(
            "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" ") > 0);
        assertTrue(xml.indexOf(
        	"xmlns:o=\"urn:schemas-microsoft-com:office:office\"") > 0); 
        assertTrue(xml.indexOf(
    		"xmlns:x=\"urn:schemas-microsoft-com:office:excel\"") > 0); 
        assertTrue(xml.indexOf(
			"xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"") > 0);         
        assertTrue(xml.indexOf(
			"xmlns:html=\"http://www.w3.org/TR/REC-html40\"") > 0);

        assertTrue(xml.indexOf("<?mso-application progid=\"Excel.Sheet\"?>") > 0);
        assertTrue(xml.indexOf("<ss:Worksheet ss:Name=\"Sheet1\"/>") > 0);
        assertTrue(xml.indexOf("<Style ss:ID=\"Default\" ss:Name=\"Normal\">") > 0);
        
        if (toFile) xmlToFile(wb);
    }
    
    public void testExcelWorkbook() throws Exception {
        Workbook wb = new XLWorkbook("test01");
        ExcelWorkbook ewb = wb.getExcelWorkbook();
        ewb.setWindowHeight(7000);
        ewb.setWindowWidth(10000);
        ewb.setWindowTopX(1000);
        ewb.setWindowTopY(500);        
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<x:WindowHeight>7000</x:WindowHeight>") > 0);
        assertTrue(xml.indexOf("<x:WindowWidth>10000</x:WindowWidth>") > 0);
        assertTrue(xml.indexOf("<x:WindowTopX>1000</x:WindowTopX>") > 0);
        assertTrue(xml.indexOf("<x:WindowTopY>500</x:WindowTopY>") > 0);
        assertTrue(xml.indexOf("<x:ProtectStructure>False</x:ProtectStructure>") > 0);
        assertTrue(xml.indexOf("<x:ProtectWindows>False</x:ProtectWindows>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testNoNameWorksheet() throws Exception {
        Workbook wb = new XLWorkbook("test02");
        wb.addSheet("");
        String xml = xmlToString(wb);
        
        assertTrue(xml.indexOf("<ss:Worksheet ss:Name=\"Sheet1\"") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testProtectedWorksheet() throws Exception {
        Workbook wb = new XLWorkbook("test03");
        Worksheet sheet = wb.addSheet();
        sheet.setProtected(true);
        String xml = xmlToString(wb);
        
        assertTrue(xml.indexOf("<ss:Worksheet ss:Name=\"Sheet1\" ss:Protected=\"1\"/>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testRightToLeftWorksheet() throws Exception {
        Workbook wb = new XLWorkbook("test04");
        Worksheet sheet = wb.addSheet();
        sheet.setProtected(true);
        sheet.setRightToLeft(true);
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<ss:Worksheet ss:Name=\"Sheet1\" " +
        		"ss:Protected=\"1\" ss:RightToLeft=\"1\"/>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testMultipleWorksheets() throws Exception {
        Workbook wb = new XLWorkbook("test05");
        wb.addSheet();
        Worksheet sheet = wb.addSheet("foo");
        sheet.setRightToLeft(true);
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<ss:Worksheet ss:Name=\"Sheet1\"/>") > 0);
        assertTrue(xml.indexOf("<ss:Worksheet ss:Name=\"foo\" " +
        		"ss:RightToLeft=\"1\"/>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testTableAttributes() throws Exception {
        Workbook wb = new XLWorkbook("test06");
        Worksheet sheet = wb.addSheet();
        Table table = sheet.getTable();
        table.setStyleID("no_definition");
        table.setDefaultRowHeight(50.35);
        table.setDefaultColumnWidth(50.35);
        
        warnings = 1;
        printWarnings = false;
        
        String xml = xmlToString(wb);
        //System.out.println(xml);
        assertTrue(xml.indexOf("ss:StyleID=\"no_definition\"") > 0);
        assertTrue(xml.indexOf("ss:DefaultRowHeight=\"50.35\"") > 0);
        assertTrue(xml.indexOf("ss:DefaultColumnWidth=\"50.35\"") > 0);

        String warningString = (String) wb.getWarnings().get(0);
        assertTrue(warningString.indexOf(
           "WARNING 1): nl.fountain.xelem.UnsupportedStyleException: " +
           "Style 'no_definition' not found.") > 0);
        
        //printWarnings = true;
        if (toFile) xmlToFile(wb);
    }
    
    public void testColumnAttributes() throws Exception {
        Workbook wb = new XLWorkbook("test07");
        Worksheet sheet = wb.addSheet();
        
        Column column = sheet.addColumnAt(5);
        column.setStyleID("b_yellow");
        column.setWidth(25.2);
        column.setSpan(5);
        
        //illegal: first free index = 5 + 5 + 1 = 11
        //Column column2 = sheet.addColumn(8);
        
        sheet.addColumn().setStyleID("b_lblue");
        
        Column column2 = sheet.addColumnAt(12);
        column2.setStyleID("bold");
        
        Column b = sheet.addColumnAt(2);
        b.setHidden(true);
        
        String xml = xmlToString(wb);
        //System.out.println(xml);
        assertTrue(xml.indexOf("ss:Index=\"2\"") > 0);
        assertTrue(xml.indexOf("ss:Index=\"5\"") > 0);
        assertTrue(xml.indexOf("ss:StyleID=\"b_lblue\"") > 0);
        assertTrue(xml.indexOf("ss:Index=\"12\"") > 0);
        
        if (toFile) xmlToFile(wb);
    }
    
    public void testRowAttributes() throws Exception {
        Workbook wb = new XLWorkbook("test08");
        Worksheet sheet = wb.addSheet();
        
        Row row = sheet.addRowAt(5);
        row.setStyleID("b_yellow");
        row.setHeight(25.2);
        row.setSpan(5);
        
        //illegal: first free index = 5 + 5 + 1 = 11
        //Row row2 = sheet.addRow(8);
        
        sheet.addRow().setStyleID("b_lblue");
        
        Row row2 = sheet.addRowAt(12);
        row2.setStyleID("bold");
        
        sheet.addRowAt(3).setHidden(true);
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("ss:Index=\"3\"") > 0);
        assertTrue(xml.indexOf("ss:Index=\"5\"") > 0);
        assertTrue(xml.indexOf("ss:StyleID=\"b_lblue\"") > 0);
        assertTrue(xml.indexOf("ss:StyleID=\"bold\"") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testColumnsAndRows() throws Exception {
        Workbook wb = new XLWorkbook("test09");
        Worksheet sheet = wb.addSheet();
        
        Column column = sheet.addColumnAt(5);
        column.setStyleID("b_yellow");
        column.setSpan(5);
        column.setWidth(25.2);
        sheet.addColumn().setStyleID("b_lblue");
        
        Row row = sheet.addRowAt(5);
        row.setStyleID("b_yellow");
        row.setSpan(5);
        row.setHeight(25.2);
        sheet.addRow().setStyleID("b_lblue");
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("ss:Index=\"5\"") > 0);
        assertTrue(xml.indexOf("ss:StyleID=\"b_lblue\"") > 0);
        assertTrue(xml.indexOf("ss:Index=\"5\"") > 0);
        assertTrue(xml.indexOf("<ss:Row ss:StyleID=\"b_lblue\"") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testCellAttributes() throws Exception {
        Workbook wb = new XLWorkbook("test10");
        Worksheet sheet = wb.addSheet();
        
        Cell c1 = sheet.addCellAt(5, 2);
        c1.setStyleID("b_yellow");
        Cell c2 = sheet.addCell();
        c2.setHRef("http://www.microsoft.com/downloads/details.aspx" +
        		"?familyid=fe118952-3547-420a-a412-00a2662442d9&displaylang=en");
        Cell c3 = sheet.addCell();
        c3.setFormula("=R1C1+R1C2");
        
        Cell c4 = sheet.addCellAt(8, 2);
        c4.setMergeAcross(2);
        Cell c5 = sheet.addCell();
        c5.setStyleID("b_lblue");
        c5.setMergeDown(2);
        
        Cell c6 = sheet.addCellAt(12, 2);
        c6.setStyleID("bold");
        c6.setMergeAcross(3);
        c6.setMergeDown(1);
        c6.setFormula("=R[-7]C[2]*1.23456789");
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<ss:Cell " +
        		"ss:HRef=\"http://www.microsoft.com/downloads/details.aspx" +
        		"?familyid=fe118952-3547-420a-a412-00a2662442d9&amp;displaylang=en\"/>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testData() throws Exception {
        Workbook wb = new XLWorkbook("test11");
        Worksheet sheet = wb.addSheet();
        sheet.addCell("test11", "bold");
        
        // formula-result goes before data
        Cell c2 = sheet.addCellAt(3,1);
        c2.setFormula("=R[-2]C&\": testing data\"");
        c2.setData("foo bar");
        
        sheet.addCellAt(5,1);
        Cell c3 = sheet.addCell(new Date(), "gray_date");
        c3.setMergeAcross(6);
        
        Cell c4 = sheet.addCellAt(7, 2);
        c4.setFormula("=NOW()");
        c4.setStyleID("gray_date");
        
        sheet.addColumnAt(2).setWidth(200);
        
        sheet.addCellAt(9, 1);
        sheet.addCell("&1<2>3\" ' € @         ");
        
        sheet.addCellAt(11, 1);
        sheet.addCell("'=1+2");
        sheet.addCell("=1+2");
        // raises error while loading:
        //sheet.addCell().setFormula("'=1+2");
        sheet.addCell().setFormula("=1+2");
        
        Cell c5 = sheet.addCellAt(13, 2);
        c5.setData(Double.MAX_VALUE);
        sheet.addCell("Double.MAX_VALUE").setStyleID("b_yellow");
        
        Cell c6 = sheet.addCellAt(14, 2);
        c6.setData(Integer.MAX_VALUE);
        sheet.addCell("Integer.MAX_VALUE", "b_yellow").setMergeAcross(2);
        
        sheet.addCell(2, "gray_date");
        sheet.addCell(2);
        
        Number number = null;
        sheet.addCellAt(18, 3).setData(number);
        sheet.addCellAt(19, 3).setData(
                new BigDecimal("1234567890123456789012345678901234567890"
                        + "12345678901234567890123456789012345678901234567890"
                        + "12345678901234567890123456789012345678901234567890"
                        + "12345678901234567890123456789012345678901234567890"
                        + "12345678901234567890123456789012345678901234567890"
                        + "12345678901234567890123456789012345678901234567890"
                        + "1234567890123456789"));
        sheet.addCellAt(20, 3).setData(Double.NaN);
        sheet.addCellAt(21, 3).setData(Double.NEGATIVE_INFINITY);
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<ss:Cell ss:StyleID=\"bold\">") > 0);
        assertTrue(xml.indexOf("<Data ss:Type=\"String\">test11</Data>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
        
    }
    
    public void testWorksheetOptions() throws Exception {
        Workbook wb = new XLWorkbook("test12");
        
        Worksheet sheet = wb.addSheet();
        sheet.addCell("test12", "bold");
        Cell c2 = sheet.addCellAt(3,1);
        c2.setFormula("=R[-2]C&\": testing WorksheetOptions\"");
        sheet.addCellAt(6, 3, c2);
        
        WorksheetOptions wso = sheet.getWorksheetOptions();
        wso.doDisplayFormulas(true);
        wso.doNotDisplayGridlines(true);
        wso.doNotDisplayHeadings(true);
        wso.setLeftColumnVisible(2);
        
        wso.setSelected(true);
        wso.setTabColorIndex(123456);
        wso.setTopRowVisible(2);
        wso.setZoom(150);
        
        
        Worksheet sheet2 = wb.addSheet("selected as well");
        WorksheetOptions wso2 = sheet2.getWorksheetOptions();
        wso2.setSelected(true);
        wso2.setGridlineColor(255, 255, 0);
        
        wb.addSheet("not visible").getWorksheetOptions()
        	.setVisible(WorksheetOptions.SHEET_HIDDEN);
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<x:GridlineColor>#ffff00</x:GridlineColor>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testSelectedSheets() throws Exception {
        Workbook wb = new XLWorkbook("test13");       
        wb.addSheet().getWorksheetOptions().setSelected(true);
        wb.addSheet().getWorksheetOptions().setSelected(true);
        wb.addSheet().getWorksheetOptions().setSelected(false);
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<x:SelectedSheets>2</x:SelectedSheets>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testActiveCell() throws Exception {
        Workbook wb = new XLWorkbook("test14");
        Worksheet sheet = wb.addSheet();
        WorksheetOptions wso = sheet.getWorksheetOptions();
        wso.setActiveCell(5, 3);
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<x:Number>3</x:Number>") > 0);
        assertTrue(xml.indexOf("<x:ActiveCol>2</x:ActiveCol>") > 0);
        assertTrue(xml.indexOf("<x:ActiveRow>4</x:ActiveRow>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testSplitHorizontal() throws Exception {
        Workbook wb = new XLWorkbook("test15");
        Worksheet sheet = wb.addSheet();
        WorksheetOptions wso = sheet.getWorksheetOptions();
        wso.setActiveCell(Pane.BOTTOM_LEFT, 11, 3);
        wso.splitHorizontal(5000, 5);
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<x:SplitHorizontal>5000</x:SplitHorizontal>") > 0);
        assertTrue(xml.indexOf("<x:TopRowBottomPane>4</x:TopRowBottomPane>") > 0);
        assertTrue(xml.indexOf("<x:ActivePane>2</x:ActivePane>") > 0);
        assertTrue(xml.indexOf("<x:Number>2</x:Number>") > 0);
        assertTrue(xml.indexOf("<x:ActiveCol>2</x:ActiveCol>") > 0);
        assertTrue(xml.indexOf("<x:ActiveRow>10</x:ActiveRow>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testSplitVertical() throws Exception {
        Workbook wb = new XLWorkbook("test16");
        Worksheet sheet = wb.addSheet();
        WorksheetOptions wso = sheet.getWorksheetOptions();
        wso.setActiveCell(Pane.TOP_RIGHT, 11, 3);
        wso.splitVertical(8000, 1);
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<x:SplitVertical>8000</x:SplitVertical>") > 0);
        assertTrue(xml.indexOf("<x:LeftColumnRightPane>0</x:LeftColumnRightPane>") > 0);
        assertTrue(xml.indexOf("<x:ActivePane>1</x:ActivePane>") > 0);
        assertTrue(xml.indexOf("<x:Number>1</x:Number>") > 0);
        assertTrue(xml.indexOf("<x:ActiveCol>2</x:ActiveCol>") > 0);
        assertTrue(xml.indexOf("<x:ActiveRow>10</x:ActiveRow>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testSplitHorizontalAndVertical() throws Exception {
        Workbook wb = new XLWorkbook("test17");
        Worksheet sheet = wb.addSheet();
        WorksheetOptions wso = sheet.getWorksheetOptions();
        wso.setActiveCell(Pane.TOP_RIGHT, 11, 3);
        wso.splitVertical(8000, 2);
        wso.splitHorizontal(5000, 5);
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<x:SplitHorizontal>5000</x:SplitHorizontal>") > 0);
        assertTrue(xml.indexOf("<x:TopRowBottomPane>4</x:TopRowBottomPane>") > 0);
        assertTrue(xml.indexOf("<x:SplitVertical>8000</x:SplitVertical>") > 0);
        assertTrue(xml.indexOf("<x:LeftColumnRightPane>1</x:LeftColumnRightPane>") > 0);
        assertTrue(xml.indexOf("<x:ActivePane>1</x:ActivePane>") > 0);
        assertTrue(xml.indexOf("<x:Number>1</x:Number>") > 0);
        assertTrue(xml.indexOf("<x:ActiveCol>2</x:ActiveCol>") > 0);
        assertTrue(xml.indexOf("<x:ActiveRow>10</x:ActiveRow>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testFreezePanes() throws Exception {
        Workbook wb = new XLWorkbook("test18");
        Worksheet sheet = wb.addSheet();
        WorksheetOptions wso = sheet.getWorksheetOptions();
        wso.freezePanesAt(5, 2);
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<x:ActivePane>0</x:ActivePane>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testAutoFilter() throws Exception {
        Workbook wb = new XLWorkbook("test19");
        Worksheet sheet = wb.addSheet();
        sheet.getCellPointer().moveTo(10, 1);
        sheet.addCell("foo", "b_yellow");
        sheet.addCell("bar", "b_yellow");
        sheet.addCell("tender", "b_yellow");
        for (int r = 0; r < 11; r++) {
            sheet.getCellPointer().moveCRLF();
            sheet.addCell(r * 5);
            sheet.addCell(r - 5);
            sheet.addCell((double)(r*5)/(r-5), "currency");
        }
        sheet.setAutoFilter("R10C1:R10C3");
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<x:AutoFilter x:Range=\"R10C1:R10C3\"/>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testAutoFitWidth() throws Exception {
        Workbook wb = new XLWorkbook("test20");
        Worksheet sheet = wb.addSheet();
        sheet.addCell(new Date(), "gray_date");
        sheet.addCell(new Date(), "gray_date");
        sheet.addCell(new Date(), "gray_date");
        
        sheet.addColumnAt(2).setAutoFitWidth(false);
        sheet.addColumnAt(3).setAutoFitWidth(true);
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("ss:AutoFitWidth=\"0\"") > 0);
        assertEquals(-1, xml.indexOf("<ss:Column ss:Index=\"3\""));
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
 
    // Doesn't seem to work
//    public void testAutoFitHeight() throws Exception {
//        Workbook wb = new XLWorkbook("test21");
//        Worksheet sheet = wb.addSheet();
//        Cell cell1 = sheet.addCellAt(1, 1);
//        cell1.setData("bla bla bla bla bla bla bla bla bla bla bla bla bla bla bla");
//        sheet.addCellAt(2, 1, cell1);
//        sheet.addCellAt(3, 1, cell1);
//        sheet.getRow(2).setAutoFitHeight(false);
//        sheet.getRow(3).setAutoFitHeight(true);
//        
//        // This works
//        sheet.getTable().setStyleID("wrap");
//        // This works
//        
//        String xml = xmlToString(wb);
//        System.out.println(xml);
//        xmlToFile(wb);
//    }
    
    public void testNamedRanges() throws Exception {
        Workbook wb = new XLWorkbook("test22");
        Worksheet blad = wb.addSheet("blad");
        
        for (int i = 1; i < 11; i++) {
            Cell cell = blad.addCellAt(i, 3);
            cell.setData(i * 5);
        }
        wb.addNamedRange("foo", "blad!R1C3:R10C3");
        Worksheet blad2 = wb.addSheet("blad2");
        blad2.addCellAt(5, 3).setFormula("=SUM(foo)");
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<ss:NamedRange ss:Name=\"foo\" " +
        		"ss:RefersTo=\"blad!R1C3:R10C3\"/>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testElementComments() throws Exception {
        Workbook wb = new XLWorkbook("test23");
        Worksheet sheet = wb.addSheet();
        sheet.addElementComment("   commentaar 1   ");
        sheet.addElementComment("   commentaar 2   ");
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<!--   commentaar 1   -->") > 0);
        assertTrue(xml.indexOf("<!--   commentaar 2   -->") > 0);
        
        wb.setPrintElementComments(false);
        xml = xmlToString(wb);
        assertEquals(-1, xml.indexOf("<!--   commentaar 1   -->"));
        assertEquals(-1, xml.indexOf("<!--   commentaar 2   -->"));
        
        wb.addElementComment("  workbook comment  ");
        wb.setPrintElementComments(true);
        xml = xmlToString(wb);
        assertTrue(xml.indexOf("<!--  workbook comment  -->") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testDocumentProperties() throws Exception {
        Workbook wb = new XLWorkbook("test24");
        wb.addSheet().addCell().setHRef("href");
        DocumentProperties dp = wb.getDocumentProperties();
        dp.setAppName("appname");
        dp.setAuthor("author");
        dp.setCategory("category");
        dp.setCompany("company");
        dp.setCreated(new Date(0));
        dp.setLastSaved(new Date());
        dp.setDescription("description");
        dp.setHyperlinkBase("file://D:/bla/bla/");
        dp.setKeywords("key words foo bar");
        dp.setLastAuthor("last author");
        dp.setManager("manager");
        dp.setSubject("subject");
        dp.setTitle("title");
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<o:Created>1970-01-01T01:00Z</o:Created>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testManipulatingTheOrderOfSheets() throws Exception {
        Workbook wb = new XLWorkbook("test25");
        wb.addSheet("blad1").addCell("cell 1");
        wb.addSheet("blad2").addCell("cell 2");
        wb.addSheet("blad3").addCell("cell 3");
        
        Worksheet s1 = wb.removeSheet("blad1");
        Worksheet s2 = wb.removeSheet("blad2");
        
        wb.addSheet(s2);
        wb.addSheet(s1);
        
        String xml = xmlToString(wb);
        int i1 = xml.indexOf("<ss:Worksheet ss:Name=\"blad1\">");
        int i2 = xml.indexOf("<ss:Worksheet ss:Name=\"blad2\">");
        int i3 = xml.indexOf("<ss:Worksheet ss:Name=\"blad3\">");
        assertTrue(i1 > i2);
        assertTrue(i2 > i3);
        
//        Collections.swap(wb.getSheetNames(), 0, 2);
//        xml = xmlToString(wb);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testRangeSelection() throws Exception {
        Workbook wb = new XLWorkbook("test26");
        WorksheetOptions wso = wb.addSheet().getWorksheetOptions();
        wso.setRangeSelection("R3C4:R12C8");
        wso.setActiveCell(3, 4);
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<x:RangeSelection>R3C4:R12C8</x:RangeSelection>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void testRangeSelection2() throws Exception {
        Workbook wb = new XLWorkbook("test27");
        WorksheetOptions wso = wb.addSheet().getWorksheetOptions();
        wso.setRangeSelection(Pane.BOTTOM_RIGHT, "R10C5:R12C8");
        wso.setActiveCell(Pane.BOTTOM_RIGHT, 11, 6);
        wso.freezePanesAt(5, 2);
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<x:RangeSelection>R10C5:R12C8</x:RangeSelection>") > 0);
        
        //System.out.println(xml);
        
        if (toFile) xmlToFile(wb);
    }
    
    public void testSpecialCharacters() throws Exception {
        Workbook wb = new XLWorkbook("test28");
        wb.addSheet().addCell("2020 BV Financiën");
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<Data ss:Type=\"String\">2020 BV Financiën</Data>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    public void test29() throws Exception {
        Workbook wb = new XLWorkbook("test29");
        wb.addSheet();
        wb.addSheet();
        Worksheet sheet = wb.addSheet();
        sheet.addCell("this sheet is selected; "
                + "the workbooks structure is protected. (cannot move sheets)");
        sheet.addCellAt(2, 1).setData("choose >Tools >Protection >Unprotect Workbook");
        wb.addSheet();
        wb.addSheet();
        wb.getExcelWorkbook().setActiveSheet(2);
        wb.getExcelWorkbook().setWindowHeight(3000);
        wb.getExcelWorkbook().setWindowTopY(400);
        wb.getExcelWorkbook().setProtectStructure(true);
        wb.getExcelWorkbook().setProtectWindows(false);
        
        String xml = xmlToString(wb);
        assertTrue(xml.indexOf("<x:ActiveSheet>2</x:ActiveSheet>") > 0);
        assertTrue(xml.indexOf("<x:ProtectStructure>True</x:ProtectStructure>") > 0);
        assertTrue(xml.indexOf("<x:ProtectWindows>False</x:ProtectWindows>") > 0);
        
        //System.out.println(xml);
        if (toFile) xmlToFile(wb);
    }
    
    
    public void testCellPointer() throws Exception {
        Workbook wb = new XLWorkbook("test30");
        Worksheet sheet = wb.addSheet();
        for (int r = 1; r < 3; r++) {
	        for (int i = 0; i < 256; i++) {
	            sheet.addCell(sheet.getCellPointer().getAbsoluteAddress());
	        }
	        sheet.getCellPointer().moveCRLF();
        }
        
        if (toFile) xmlToFile(wb);
    }
    
    public void testComments31() throws Exception {
        Workbook wb = new XLWorkbook("test31");
        Worksheet sheet = wb.addSheet();
        sheet.getCellPointer().moveTo(10, 5);
        sheet.addCell("no comment").addComment();
        Comment comment = sheet.addCell("1e commentaar").addComment();
        comment.setData("this is comment");
        
        comment = sheet.addCell("2e commentaar").addComment();
        comment.setShowAlways(true);
        comment.setData("this comment shows always.");
        sheet.getCellPointer().move(0, 3);
        sheet.addCell("3e commentaar").addComment(comment);
        
        sheet.getCellPointer().moveCRLF();
        comment = sheet.addCell("empty comment").addComment();
        comment.setData("");
        
        // deprecated
//        comment = sheet.addCell("lf comment").addComment();
//        comment.setAuthor("the great author");
//        comment.setData("<B><I><Font html:Face=\"Tahoma\" html:Size=\"9\" "
//                + "html:Color=\"#000000\">the great author:</Font></I></B>"
//                + "<Font html:Face=\"Tahoma\" html:Size=\"9\" "
//                + "html:Color=\"#000000\">&#10;dit is&#10;commentaar.</Font>");
        
        if (toFile) xmlToFile(wb);
    }
    
//    public void testNoDocumentCreation() {
//    		 //goes out of memory at 1700 rows
//        Workbook wb = new XLWorkbook("noname");
//        Worksheet sheet = wb.addSheet();
//        for (int r = 1; r < 1600; r++) {
//	        for (int i = 0; i < 256; i++) {
//	            sheet.addCell(sheet.getCellPointer().getAbsoluteAddress());
//	        }
//	        sheet.getCellPointer().moveCRLF();
//        }
//    }
    
    private String xmlToString(Workbook wb) throws Exception {
        StringWriter sw = new StringWriter();
        Result result = new StreamResult(sw);
        transform(wb, result);
        return sw.toString();
    }
    
    private void xmlToFile(Workbook wb) throws Exception {
        Result result = new StreamResult(
                new File(testOutputDir + wb.getFileName()));
        transform(wb, result);
    }
    
    private void transform(Workbook wb, Result result) throws Exception {
        Document doc = wb.createDocument();
        if (printWarnings) {
		    for (String s : wb.getWarnings()) {
		        System.out.println(s);           
		    }
        }
        assertEquals(warnings, wb.getWarnings().size());
        TransformerFactory tFactory = TransformerFactory.newInstance();
        Transformer xformer = tFactory.newTransformer();
        xformer.setOutputProperty(OutputKeys.METHOD, "xml");
        xformer.setOutputProperty(OutputKeys.INDENT, "yes");
        
        Source source = new DOMSource(doc);
        xformer.transform(source, result);
    }

}
