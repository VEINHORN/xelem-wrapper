/*
 * Created on Dec 24, 2004
 * Copyright (C) 2005  Henk van den Berg
 * 
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
 * 
 * see license.txt
 *
 */
package nl.fountain.xelem.expat;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.NoSuchElementException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.FactoryConfigurationError;
import javax.xml.parsers.ParserConfigurationException;

import nl.fountain.xelem.Address;
import nl.fountain.xelem.XelemException;
import nl.fountain.xelem.excel.Cell;
import nl.fountain.xelem.excel.Row;
import nl.fountain.xelem.excel.XLElement;

import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

/**
 * A class that will take an existing xml-spreadsheet as a template.
 * <P>
 * <i>Sometimes you just want to dump a lot of data produced within your java-
 * application into a preformatted template. Building an 
 * {@link nl.fountain.xelem.excel.ss.XLWorkbook XLWorkbook} from scratch may be
 * time consuming and seems a waste of effort.</i>
 * <P>
 * Whereas the strategy of xelem can be described as arranging
 * a conglomerate of Java-classinstances that transform themselves into
 * an xml-document, this class uses a different approach. It parses an
 * existing template-file and tinkers with xml-elements in the derived document,
 * acting as a facade around these surgical opperations.
 * <P>
 * Create your template in Excel. Select columns to set formatting (i.e. style-id's)
 * on entire columns. If you wish, you may put table headings, formula's, titles
 * and what have you, in the first rows of the sheets. Save your template as an
 * XML Spreadsheet and
 * instantiate an instance of XLDocument with the parameter 
 * <code>fileName</code> set to the path where your template resides. The methods
 * {@link #appendRow(String, Row) appendRow(String sheetName, Row row)} and 
 * {@link #appendRows(String, Collection) appendRows(String sheetName, Collection rows)}
 * will append the {@link nl.fountain.xelem.excel.Row rows} as row-elements
 * directly under the last row-element of the sheet. As long as the appended rows
 * and cells do not have a StyleID set, your data
 * will be formatted according to the style-id's which were set on your columns
 * during the creation of the template. And offcourse, 
 * you could set StyleID's on rows and cells, but take care to only use id's that have
 * a definition in the section <code>&lt;Styles&gt;</code> in your template.
 * <P>
 * The method {@link #setCellData(Cell, String, int, int)} will set or replace
 * (only) the data-element of the mentioned cell and is typically used
 * to add header information to a bunch of data appended with
 * appendRow/appendRows.
 * <P>&nbsp;<P>
 * <H2>The PivotTable Cookbook</H2>
 * If you want to include PivotTables in your template, it is recommended to
 * read this cook-book before endeavouring on such a quest.<br>
 * In this recipe we'll create a template with a PivotTable. After that we
 * populate it with data from within our Java application and output a new
 * file with both our data and the PivotTable ready for analyses. 
 * Open your Excel-application with a new
 * workbook and we'll start.
 * 
 * <H3>Step 1. Set up the data source table</H3>
 * Rename a sheet "data" and set up a table heading, starting in cell A1,
 * like this:
 * <P>
 * <img src="doc-files/sourcetable.gif">
 * <P>
 * We could set a currency format on column D.
 * 
 * <H3>Step 2. Set up the pivot table</H3>
 * Rename another sheet "report" and place the cellpointer in cell A5. Point to
 * <b>Data</b> in the menu, and choose <b>PivotTable and PivotChart Report...</b>.
 * <UL>
 * <LI>Answer the question 'Where is the data that you want to analyze?' with 
 * 'Microsoft Office 
 * Excel list or database', and answer 'What kind of report do you want to create?' 
 * with 'PivotTable'. 
 * <LI>Choose 'Next...'. 
 * <LI>Click on the find-range-button and point to
 * 'data!$A$1:$E$2'. 
 * <LI>Choose 'Next...'. 
 * <LI>Do the initial layout of your PivotTable...
 * <LI>Choose 'Options...' and check the 'Refresh on open' box. 
 * <LI>Choose 'Finish...'. 
 * </UL>
 * We
 * could apply an autoformat by pointing to <b>Format</b> &gt; <b>Autoformat...</b>
 * in the menu and choose a format. After that our PivotTable may look like this:
 * <P>
 * <img src="doc-files/ptreport.gif">
 * 
 * <H3>Step 3. Save the template</H3>
 * Save the workbook as an XML Spreadsheet (*.xml). We'll name it 'template.xml'.
 * Mark that a section of the PTSource-element in the saved xml-file looks like this:
 * <PRE>
 *    &lt;ConsolidationReference&gt;
 *      &lt;FileName&gt;[template.xml]data&lt;/FileName&gt;
 *      &lt;Reference&gt;R1C1:R2C5&lt;/Reference&gt;
 *    &lt;/ConsolidationReference&gt;</PRE>
 * We'll need to change the text in both FileName- and Reference-elements
 * in the file that is going to be output. We'll do that in the Java-part.
 * 
 * <H3>Step 4. Do the Java-part</H3>
 * <PRE>
 *     public void testPivot() {
 *       Object[][] data = {
 *           	{"blue", "A", "Star", new Double(2.95), new Double(55.6)},
 *           	{"red", "A", "Planet", new Double(3.10), new Double(123.5)},
 *           	{"green", "C", "Star", new Double(3.21), new Double(20.356)},
 *           	{"green", "B", "Star", new Double(4.23), new Double(456)},
 *           	{"red", "B", "Planet", new Double(4.21), new Double(789)},
 *           	{"blue", "D", "Planet", new Double(4.51), new Double(9.6)},
 *           	{"yellow", "A", "Commet", new Double(4.15), new Double(19.8)}
 *           };
 *       
 *       // set up a collection of rows and populate them with data.
 *       // your data will probably be collected in a more sophisticated way.
 *       Collection rows = new ArrayList();
 *       for (int r = 0; r < data.length; r++) {
 *           Row row = new SSRow();
 *           for (int c = 0; c < data[r].length; c++) {
 *               row.addCell(data[r][c]);
 *           }
 *           rows.add(row);
 *       }
 *       
 *       OutputStream out;
 *       try {
 * 
 *           // create the XLDocument
 *           XLDocument xlDoc = new XLDocument("template.xml");
 * 
 *           // append the rows to the sheet 'data'
 *           xlDoc.appendRows("data", rows);
 * 
 *           // this will be the new filename
 *           String fileName = "prices.xls";
 * 
 *           // we'll change the FileName-element to reflect our new filename
 *           xlDoc.setPTSourceFileName(fileName, "data");
 * 
 *           // we'll change the Reference-element to reflect
 *           // the size of the new source table
 *           xlDoc.setPTSourceReference(1, 1, rows.size() + 1, 5);
 * 
 *           // output the new file
 *           out = new BufferedOutputStream(new FileOutputStream(fileName));
 *           new XSerializer().serialize(xlDoc.getDocument(), out);
 *           out.close();
 * 
 *       } catch (XelemException e) {
 *           e.printStackTrace();
 *       } catch (IOException e) {
 *           e.printStackTrace();
 *       }        
 *   }
 * </PRE>
 * 
 * @since xelem.1.0.1
 * 
 */
public class XLDocument {
    
    private Document doc;
    private Map<String, Node> tableMap;
    
    /**
     * Creates a new XLDocument by parsing the specified file into a
     * {@link org.w3c.dom.Document}.
     * 
     * @param fileName	the filename of the xml-spreadsheet template
     * @throws XelemException	if loading or parsing of the document fails.
     * <br>Could be any of:
     * <UL>
     * <LI>{@link java.io.FileNotFoundException}
     * <LI>{@link javax.xml.parsers.FactoryConfigurationError}
     * <LI>{@link javax.xml.parsers.ParserConfigurationException}
     * <LI>{@link org.xml.sax.SAXException}
     * <LI>{@link java.io.IOException}
     * </UL>
     */
    public XLDocument(String fileName) throws XelemException {
        doc = loadDocument(fileName);
        tableMap = new HashMap<String, Node>();
    }
    
    /**
     * Gets the underlying {@link org.w3c.dom.Document Document}-implementation.
     * Use one of the serialize-methods of {@link nl.fountain.xelem.XSerializer}
     * to serialize this document to a file, outputstream or writer.
     * 
     * @return the document for which this XLDocument acts as a facade
     */
    public Document getDocument() {
        return doc;
    }
    
    /**
     * Appends the {@link nl.fountain.xelem.excel.Row} to the mentioned sheet.
     * The row will be appended right under the last row-elemnt.
     * 
     * @param sheetName	the name of the sheet to which the row must be appended
     * @param row	the row to append
     * 
     * @throws java.util.NoSuchElementException	if the mentioned sheet 
     * 	was not found in the template.
     */
    public void appendRow(String sheetName, Row row) {
        Element rowElement = row.createElement(doc);
        getTableElement(sheetName).appendChild(rowElement);
    }
    
    /**
     * Appends all the {@link nl.fountain.xelem.excel.Row rows} in the Collection
     * to the mentioned sheet. The rows will be appended in the order of the 
     * collection right under the last row-element.
     * 
     * @param sheetName	the name of the sheet to which the rows must be appended
     * @param rows	a Collection of rows
     * 
     * @throws java.util.NoSuchElementException	if the mentioned sheet 
     * 	was not found in the template.
     */
    public void appendRows(String sheetName, Collection<Row> rows) {
        Node table = getTableElement(sheetName);
        for (Iterator<Row> iter = rows.iterator(); iter.hasNext();) {
            table.appendChild(iter.next().createElement(doc));
        }      
    }
    
    /**
     * Will set or replace (only) the data-element of the cell-element at 
     * the intersection of the mentioned row- and columnIndex. 
     * This is a time-consuming opperation. A test with a thousand 
     * iterations of adding a cell with the Worksheet-method
     * {@link nl.fountain.xelem.excel.Worksheet#addCell(int)} took 31 milliseconds,
     * while this method took as long as 1219 ms to do the same.
     * As long as you do not use it for bulk-opperations no significant
     * time-lack will be manifest. Typically this method is used
     * to add header information to a bunch of data which you added with
     * appendRow/appendRows.
     * 
     * @param cell	the Cell holding the data
     * @param sheetName the name of the template-sheet
     * @param rowIndex	the vertical cell-index
     * @param columnIndex	the horizontal cell-index
     * 
     * @throws java.util.NoSuchElementException	if the mentioned sheet 
     * 	was not found in the template.
     */
    public void setCellData(Cell cell, String sheetName, int rowIndex, int columnIndex) {
        Element data = cell.getDataElement(doc);
        Element cellElement = getCellElement(sheetName, rowIndex, columnIndex);
        NodeList nodelist = cellElement.getChildNodes();
        Node oldData = null;
        for (int i = 0; i < nodelist.getLength(); i++) {
            if ("Data".equals(nodelist.item(i).getLocalName())) {
                oldData = nodelist.item(i);
            }
        }
        if (oldData == null) {
            cellElement.appendChild(data);
        } else {
            cellElement.replaceChild(data, oldData);
        }
    }
    
    /**
     * Replaces the old text in FileName elements with a
     * new one. This method may be usefull if your template contains
     * pivot tables and the ConsolidationReference in the 
     * PTSource-element looks like this:
     * <PRE>
     *    &lt;ConsolidationReference&gt;
     *      &lt;FileName&gt;<b>template.xls</b>&lt;/FileName&gt;
     *      &lt;Name&gt;products&lt;/Name&gt;
     *    &lt;/ConsolidationReference&gt;
     * </PRE>
     * The ConsolidationReference element has two children: FileName and Name.
     * You defined a name to point to the data source range of your pivot table.
     * The FileName element contains the name of the template-file and needs
     * to be replaced because the new workbook will have a different name.
     * 
     * @param fileName 	the new text for the element FileName
     * @return 	the number of replacements
     */
    public int setPTSourceFileName(String fileName) {
        int n = 0;
        NodeList nodelist = doc.getElementsByTagName("FileName");
        if (nodelist.getLength() == 0) {
            nodelist = doc.getElementsByTagNameNS(XLElement.XMLNS_X, "FileName");
        }
        for (int i = 0; i < nodelist.getLength(); i++) {
            Node fileNameElement = nodelist.item(i);           
            Node oldTekst = fileNameElement.getFirstChild();
            Node newTekst = doc.createTextNode(fileName);
            fileNameElement.replaceChild(newTekst, oldTekst);
            n++;
        }
        return n;
    }
    
    /**
     * Replaces the old text in FileName elements with a
     * new one. This method may be usefull if your template contains
     * pivot tables and the ConsolidationReference in the 
     * PTSource-element looks like this:
     * <PRE>
     *    &lt;ConsolidationReference&gt;
     *      &lt;FileName&gt;<b>[template.xls]Sheet2</b>&lt;/FileName&gt;
     *      &lt;Reference&gt;R1C1:R2C5&lt;/Reference&gt;
     *    &lt;/ConsolidationReference&gt;
     * </PRE>
     * The ConsolidationReference element has two children: FileName and Reference.
     * The FileName element contains the name of the template-file 
     * in square brackets, followed by the name of the worksheet.
     * The text in the FileName element needs
     * to be replaced because the new workbook will have a different name.
     * 
     * @param fileName the new workbook name
     * @param sheetName the name of the sheet containing the data source range
     * @return 	the number of replacements
     */
    public int setPTSourceFileName(String fileName, String sheetName) {
        return setPTSourceFileName("[" + fileName + "]" + sheetName);
    }
    
    /**
     * Replaces the old text in Reference elements with a
     * new one. This method may be usefull if your template contains
     * pivot tables and the ConsolidationReference in the 
     * PTSource-element looks like this:
     * <PRE>
     *   &lt;ConsolidationReference&gt;
     *     &lt;FileName&gt;[prices.xls]Sheet1&lt;/FileName&gt;
     *     &lt;Reference&gt;<b>R1C1:R2C5</b>&lt;/Reference&gt;
     *   &lt;/ConsolidationReference&gt;
     * </PRE>
     * 
     * @param reference		a String of R1C1 reference style
     * @return 	the number of replacements
     */
    public int setPTSourceReference(String reference) {
        int n = 0;
        NodeList nodelist = doc.getElementsByTagName("Reference");
        if (nodelist.getLength() == 0) {
            nodelist = doc.getElementsByTagNameNS(XLElement.XMLNS_X, "Reference");
        }
        for (int i = 0; i < nodelist.getLength(); i++) {
            Node element = nodelist.item(0);           
            Node oldTekst = element.getFirstChild();
            Node newTekst = doc.createTextNode(reference);
            element.replaceChild(newTekst, oldTekst);
            n++;
        }
        return n;
    }
    
    /**
     * Replaces the old text in Reference elements with a
     * new string of R1C1 reference style. The string is composed as the absolute range
     * of address1 and address2. This method may be usefull if your template contains
     * pivot tables and the ConsolidationReference in the 
     * PTSource-element looks like this:
     * <PRE>
     *   &lt;ConsolidationReference&gt;
     *     &lt;FileName&gt;[prices.xls]Sheet1&lt;/FileName&gt;
     *     &lt;Reference&gt;<b>R1C1:R2C5</b>&lt;/Reference&gt;
     *   &lt;/ConsolidationReference&gt;
     * </PRE>
     * 
     * @param address1	the top-left address of the data source range
     * @param address2	the bottom-right address of the data source range
     * @return the number of replacements
     */
    public int setPTSourceReference(Address address1, Address address2) {
       String ref = address1.getAbsoluteRange(address2);
       return setPTSourceReference(ref);
    }
    
    /**
     * Replaces the old text in Reference elements with a
     * new string of R1C1 reference style. The string is composed as the absolute range
     * of (row) r1, (column) c1 and (row) r2, (column) c2.
     * This method may be usefull if your template contains
     * pivot tables and the ConsolidationReference in the 
     * PTSource-element looks like this:
     * <PRE>
     *   &lt;ConsolidationReference&gt;
     *     &lt;FileName&gt;[prices.xls]Sheet1&lt;/FileName&gt;
     *     &lt;Reference&gt;<b>R1C1:R2C5</b>&lt;/Reference&gt;
     *   &lt;/ConsolidationReference&gt;
     * </PRE>
     * 
     * @param r1 	the top row of the data source range
     * @param c1	the left-most column of the data source range
     * @param r2	the bottom row of the data source range
     * @param c2	the right-most column of the data source range
     * @return the number of replacements
     */
    public int setPTSourceReference(int r1, int c1, int r2, int c2) {
        return setPTSourceReference(new Address(r1, c1).getAbsoluteRange(r2, c2));
    }
    
    /**
     * Gets the Worksheet element with the given name.
     * 
     * @param sheetName the name of the worksheet
     * @return the Worksheet element
     * @throws java.util.NoSuchElementException if a Worksheet element with such
     * a name does not exist
     */
    protected Element getSheetElement(String sheetName) {
        NodeList nodelist = doc.getElementsByTagName("Worksheet");
        if (nodelist.getLength() == 0) {
            nodelist = doc.getElementsByTagNameNS(XLElement.XMLNS_SS, "Worksheet");
        }
        for (int i = 0; i < nodelist.getLength(); i++) {
            Node sheetNode = nodelist.item(i);
            NamedNodeMap atrbs = sheetNode.getAttributes();
            Node nameAtrb = atrbs.getNamedItemNS(XLElement.XMLNS_SS, "Name");
            if (nameAtrb != null && sheetName.equals(nameAtrb.getNodeValue())) {
                return (Element) sheetNode;
            }
        }
        throw new NoSuchElementException("The worksheet '" + sheetName +
                "' does not exist.");
    }
    
    /**
     * Gets the Table element that is the child of the Worksheet element with the 
     * given name. If this element did not exist, it will be created.
     * 
     * @param sheetName the name of the worksheet
     * @return the Table element
     * @throws java.util.NoSuchElementException if a Worksheet element with such
     * a name does not exist
     */
    protected Element getTableElement(String sheetName) {
        Element tableElement = (Element) tableMap.get(sheetName);
        if (tableElement == null) {
	        Node sheet = getSheetElement(sheetName);
            NodeList sheetKits = sheet.getChildNodes();
            for (int i = 0; i < sheetKits.getLength(); i++) {
                Node node = sheetKits.item(i);
                if ("Table".equals(node.getLocalName())) {
                    NamedNodeMap atrbs = node.getAttributes();
                    if (atrbs.getNamedItemNS(
                            XLElement.XMLNS_SS, "ExpandedColumnCount") != null) {
                        atrbs.removeNamedItemNS(
                                XLElement.XMLNS_SS, "ExpandedColumnCount");
                    }
                    if (atrbs.getNamedItemNS(
                            XLElement.XMLNS_SS, "ExpandedRowCount") != null) {
                        atrbs.removeNamedItemNS(
                                XLElement.XMLNS_SS, "ExpandedRowCount");
                    }
                    tableMap.put(sheetName, node);
                    return (Element) node;
                }
            }
            if (tableElement == null) {
                tableElement = doc.createElementNS(XLElement.XMLNS_SS, "Table");
                sheet.appendChild(tableElement);
                tableMap.put(sheetName, tableElement);
            }
        }
        
        return tableElement;
    }
    
    /**
     * Gets the Cell element at the given worksheet at the given index. 
     * If this element did not exist, it will be created.
     * 
     * @param sheetName	the name of the worksheet
     * @param rowIndex	the vertical cell-index 
     * @param columnIndex	the horizontal cell-index
     * @return a Cell element
     * @throws java.util.NoSuchElementException if a Worksheet element with such
     * a name does not exist
     */
    protected Element getCellElement(String sheetName, int rowIndex, int columnIndex) {
        Element row = getRowElement(sheetName, rowIndex);
        NodeList nodelist = row.getChildNodes();
        int teller = 0;
        for (int i = 0; i < nodelist.getLength(); i++) {
            Node node = nodelist.item(i);
            if ("Cell".equals(node.getLocalName())) {
                teller++;
                Node idx = node.getAttributes().getNamedItemNS(
                        XLElement.XMLNS_SS, "Index");
                if (idx != null) {
                    teller = Integer.parseInt(idx.getNodeValue());
                }
                if (teller == columnIndex) {
                    return (Element) node;
                } else if (teller > columnIndex) {
                    Element cellElement = createIndexedElement("Cell", columnIndex);
                    row.insertBefore(cellElement, node);
                    return cellElement;
                }
            }
        }
        Element cellElement = createIndexedElement("Cell", columnIndex);
        row.appendChild(cellElement);
        return cellElement;
    }
    
    /**
     * Gets the Row element at the given worksheet with the given index.
     * If this element did not exist, it will be created.
     * 
     * @param sheetName	the name of the worksheet
     * @param index the index of the row
     * @return a Row element
     * @throws java.util.NoSuchElementException if a Worksheet element with such
     * a name does not exist
     */
    protected Element getRowElement(String sheetName, int index) {
        Element table = getTableElement(sheetName);
        NodeList nodelist = table.getChildNodes();
        int teller = 0;
        for (int i = 0; i < nodelist.getLength(); i++) {
            Node node = nodelist.item(i);
            if ("Row".equals(node.getLocalName())) {
                teller++;
                Node idx = node.getAttributes().getNamedItemNS(
                        XLElement.XMLNS_SS, "Index");
                if (idx != null) {
                    teller = Integer.parseInt(idx.getNodeValue());
                }
                if (teller == index) {
                    return (Element) node;
                } else if (teller > index) {
                    Element rowElement = createIndexedElement("Row", index);
                    table.insertBefore(rowElement, node);
                    return rowElement;
                }
            }
        }
        Element rowElement = createIndexedElement("Row", index);
        table.appendChild(rowElement);
        return rowElement;
    }
    
    
    private Element createIndexedElement(String qName, int index) {
        Element element = 
            doc.createElementNS(XLElement.XMLNS_SS, qName);
        element.setPrefix(XLElement.PREFIX_SS);
        element.setAttributeNodeNS(
            createAttributeNS("Index", XLElement.XMLNS_SS, 
                    XLElement.PREFIX_SS, index));
        return element;
    }
    
    private Attr createAttributeNS(String qName, String nameSpace, String prefix, int i) {
        Attr attr = doc.createAttributeNS(nameSpace, qName);
        attr.setPrefix(prefix);
        attr.setValue("" + i);
        return attr;
    }
    
    private Document loadDocument(String fileName) throws XelemException {
        Document document = null;
        InputStream is = null;
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            is = new FileInputStream(fileName);
            document = builder.parse(is);
            is.close();
        } catch (FileNotFoundException e) {
            throw new XelemException(e.fillInStackTrace());
        } catch (FactoryConfigurationError e) {
            throw new XelemException(e.fillInStackTrace());
        } catch (ParserConfigurationException e) {
            throw new XelemException(e.fillInStackTrace());
        } catch (SAXException e) {
            throw new XelemException(e.fillInStackTrace());
        } catch (IOException e) {
            throw new XelemException(e.fillInStackTrace());
        }
        return document;
    }
    

}
