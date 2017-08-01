/*
 * Created on Sep 7, 2004
 * Copyright (C) 2004  Henk van den Berg
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
package nl.fountain.xelem.excel.ss;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.UnsupportedStyleException;
import nl.fountain.xelem.XFactory;
import nl.fountain.xelem.XelemException;
import nl.fountain.xelem.excel.AbstractXLElement;
import nl.fountain.xelem.excel.DocumentProperties;
import nl.fountain.xelem.excel.DuplicateNameException;
import nl.fountain.xelem.excel.ExcelWorkbook;
import nl.fountain.xelem.excel.NamedRange;
import nl.fountain.xelem.excel.Workbook;
import nl.fountain.xelem.excel.Worksheet;
import nl.fountain.xelem.excel.o.ODocumentProperties;
import nl.fountain.xelem.excel.x.XExcelWorkbook;

import org.w3c.dom.DOMImplementation;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

/**
 * An implementation of the XLElement Workbook, the root of a SpreadsheetML document.
 * <P>
 * Typically, the XLWorkbook is at the start of creating an Excel workbook in 
 * SpreadsheetML. Mostly all of the other objects in xelem can be obtained
 * from it or through it by one of the addXxx- methods.
 * <P>
 * After setting up the workbook, you can obtain the 
 * {@link org.w3c.dom.Document org.w3c.dom.Document} from it (see
 * {@link #createDocument()}) or serialize the workbook by means of one of the
 * serialize-methods of the 
 * {@link nl.fountain.xelem.XSerializer XSerializer.class},
 * which was included in xelem for convenience.
 * <P>
 * Worksheets are displayed in Excel in the order that they were added to the 
 * workbook. The methods {@link #removeSheet(String)} and 
 * {@link #addSheet(Worksheet)} may savely be applied to change this order.
 * A more convenient way is to obtain the 
 * list of sheet names with {@link #getSheetNames()} and manupulating this list.
 * 
 * @see <a href="../../../../../overview-summary.html#overview_description">overview</a>
 * 
 */
public class XLWorkbook extends AbstractXLElement implements Workbook {
    
    private DocumentProperties documentProperties;
    private ExcelWorkbook excelWorkbook;
    private String name;
    private String filename;
    private Collection<String> docComments;
    private Map<String, Worksheet>sheets;
    private List<String> sheetList;
    private Map<String, NamedRange> namedRanges;
    private boolean printComments = true;
    private boolean printDocComments = true;
    private boolean appendInfoSheet;
    private XFactory xFactory;
    private List<String> warnings;
    private SimpleDateFormat sdf;
    
    /**
     * Creates a new XLWorkbook.
     *
     */
    public XLWorkbook() {
        this("");
    }
    
    /**
     * Creates a new XLWorkbook with the given name.
     */
    public XLWorkbook(String name) {
        sheets = new HashMap<String, Worksheet>();
        sheetList = new ArrayList<String>();
        this.name = name;
    }
    
    public void setName(String name) {
        this.name = name;
    }
    
    public String getName() {
        return name;
    }

    public void setFileName(String filename) {
        this.filename = filename;
    }

    public String getFileName() {
        if (filename == null) {
            return name + ".xls";
        } else {
            return filename;
        }
    }
    
    public void mergeStyles(String newID, String id1, String id2) 
    										throws UnsupportedStyleException {
        getFactory().mergeStyles(newID, id1, id2);
    }
    
    public void appendInfoSheet() {
        appendInfoSheet = true;
    }
    
    public void setDocumentProperties(DocumentProperties docProps) {
        documentProperties = docProps;
    }

    public DocumentProperties getDocumentProperties() {
        if (documentProperties == null) {
            documentProperties = new ODocumentProperties();
        }
        return documentProperties;
    }

    public boolean hasDocumentProperties() {
        return documentProperties != null;
    }
    
    public void setExcelWorkbook(ExcelWorkbook excelWb) {
        excelWorkbook = excelWb;
    }
    
    public ExcelWorkbook getExcelWorkbook() {
        if (excelWorkbook == null) {
            excelWorkbook = new XExcelWorkbook();
        }
        return excelWorkbook;
    }
    
    public boolean hasExcelWorkbook() {
        return excelWorkbook != null;
    }
    
    public NamedRange addNamedRange(NamedRange nr) {
        if (namedRanges == null) {
            namedRanges = new HashMap<String, NamedRange>();
        }
        namedRanges.put(nr.getName(), nr);
        return nr;
    }
    
    public NamedRange addNamedRange(String name, String refersTo) {
        return addNamedRange(new SSNamedRange(name, refersTo));
    }  
    
    public Map<String, NamedRange> getNamedRanges() {
        if (namedRanges == null) {
            return Collections.emptyMap();
        } else {
            return namedRanges;
        }
    }
    
    public Worksheet addSheet() {
        int nr = sheets.size();
        String name;
        do
            name = "Sheet" + ++nr;
        while(sheetList.contains(name));
        return addSheet(name);
    }
    
    public Worksheet addSheet(String name) {
        if (name == null || "".equals(name)) {
            return addSheet();
        }
        return addSheet(new SSWorksheet(name));
    }

    public Worksheet addSheet(Worksheet sheet) {
        if (sheetList.contains(sheet.getName())) {
           throw new DuplicateNameException(
                   "Duplicate name in worksheets collection: '"
                   + sheet.getName() + "'."); 
        }
        sheetList.add(sheet.getName());
        sheets.put(sheet.getName(), sheet);
        return sheet;
    }

    public List<Worksheet> getWorksheets() {
        List<Worksheet> worksheets = new ArrayList<Worksheet>();
        for (String s : sheetList) {
            worksheets.add(sheets.get(s));           
        }
        return worksheets;
    }
    
    public List<String> getSheetNames() {
        return sheetList;
    }

    public Worksheet getWorksheet(String name) {
        return sheets.get(name);
    }
    
    public Worksheet getWorksheetAt(int index) {
        Worksheet ws = null;
        try {
            ws = sheets.get(sheetList.get(index));
        } catch (IndexOutOfBoundsException e) {
            //
        }
        return ws;
    }
    
    public Worksheet removeSheet(String name) {
        int index = sheetList.indexOf(name);
        if (index < 0) return null;
        sheetList.remove(index);
        return sheets.remove(name);
    }

    public void setPrintElementComments(boolean print) {
        printComments = print;
    }

    public void setPrintDocComments(boolean print) {
        printDocComments = print;
    }

    public boolean isPrintingElementComments() {
        return printComments;
    }

    public boolean isPrintingDocComments() {
        return printDocComments;
    }

    public String getTagName() {
        return "Workbook";
    }
    
    public String getNameSpace() {
        return XMLNS;
    }
    
    public String getPrefix() {
        return PREFIX_SS;
    }
    
    public Document createDocument() throws ParserConfigurationException {
        GIO gio = new GIO();
        Document doc = getDoc();
        Element root = doc.getDocumentElement();
        assemble(root, gio);
        return doc;
    }
    
    public Element assemble(Element root, GIO gio) {
        Document doc = root.getOwnerDocument();
        warnings = null;
        gio.setPrintComments(isPrintingElementComments());
        doc.insertBefore(doc.createProcessingInstruction("mso-application",
                "progid=\"Excel.Sheet\""), root);
        
        if (isPrintingDocComments()) {
            for (String s : getFactory().getDocComments()) {
                doc.insertBefore(doc.createComment(s), root);
            }
        }
        
        root.setAttribute("xmlns", XMLNS);
        root.setAttribute("xmlns:o", XMLNS_O);
        root.setAttribute("xmlns:x", XMLNS_X);
        root.setAttribute("xmlns:ss", XMLNS_SS);
        root.setAttribute("xmlns:html", XMLNS_HTML);
        
        if (isPrintingElementComments() && getElementComments() != null) {
            for (String s : getElementComments()) {
                root.appendChild(doc.createComment(s));
            }
        }

        //o:DocumentProperties
        if (hasDocumentProperties()) {
            documentProperties.assemble(root, gio);
        }

        //x:ExcelWorkbook
        Element xlwbe = getExcelWorkbook().assemble(root, gio);

        //Styles
        Element styles = doc.createElement("Styles");
        root.appendChild(styles);
        appendDefaultStyle(doc, styles);
        
        // Names
        if (namedRanges != null) {
            Element names = doc.createElement("Names");
            root.appendChild(names);
            for (NamedRange nr : namedRanges.values()) {
                nr.assemble(names, gio);
            }
        }

        // Worksheets
        if (sheets.size() < 1) {
            addSheet();
        }
        for (String s : sheetList) {
            Worksheet ws = sheets.get(s);
            ws.assemble(root, gio);
        }
        
        // append xelem-info sheet
        if (appendInfoSheet) {
            try {
                getFactory().appendInfoSheet(root, gio);
            } catch (XelemException e) {
                addWarning(e.getCause());
            }
        }
        
        // append Global Information
        int selectedSheets = gio.getSelectedSheetsCount();
        if ( selectedSheets > 1) {
            Element n = doc.createElementNS(XMLNS_X, "SelectedSheets");
            n.setPrefix(PREFIX_X);
            n.appendChild(doc.createTextNode("" + selectedSheets));
            xlwbe.appendChild(n);
        }        
        appendStyles(doc, styles, gio);
        
        return root;
    }

    public List<String> getWarnings() {
        if (warnings == null) {
            return Collections.emptyList();
        } else {
            return warnings;
        }
    }
    
    private Document getDoc() throws ParserConfigurationException {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        DocumentBuilder builder = factory.newDocumentBuilder();
        DOMImplementation domImpl = builder.getDOMImplementation();
        return domImpl.createDocument(XMLNS, getTagName(), null);
    }
    
    private void appendDefaultStyle(Document doc, Element styles) {
        Element dse = getFactory().getStyle("Default");
        if (dse == null) {
            dse = doc.createElement("Style");
            
            dse.setAttributeNodeNS(createAttributeNS(doc, "ID", "Default"));
            dse.setAttributeNodeNS(createAttributeNS(doc, "Name", "Normal"));
            
            Element alignment = doc.createElement("Alignment");
            dse.appendChild(alignment);
            alignment.setAttributeNodeNS(createAttributeNS(doc, "Vertical", "Bottom"));
            
            dse.appendChild(doc.createElement("Borders"));
            dse.appendChild(doc.createElement("Font"));
            dse.appendChild(doc.createElement("Interior"));
            dse.appendChild(doc.createElement("NumberFormat"));
            dse.appendChild(doc.createElement("Protection"));
        } else {
            dse = (Element) doc.importNode(dse, true);
        }
        styles.appendChild(dse);
    }
    
    private void appendStyles(Document doc, Element styles, GIO gio) {
        for (String id : gio.getStyleIDSet()) {;
            Element style = getFactory().getStyle(id);
            if (style == null) {
                // last resort: create one on the spot
                style = doc.createElement("Style");
                style.setAttributeNodeNS(createAttributeNS(doc, "ID", id));
                addWarning(new UnsupportedStyleException(
                        "Style '" + id + "' not found."));
            } else {
                // we have a style from the XFactory
                style = (Element) doc.importNode(style, true);
            }
            styles.appendChild(style);
        }
    }
    
    private XFactory getFactory() {
        if (xFactory == null) {
            try {
                xFactory = XFactory.newInstance();
            } catch (XelemException e) {
                addWarning(e.getCause());
                // return an empty factory
                xFactory = XFactory.emptyFactory();
            } 
        }
        return xFactory;
    }
    
    private void addWarning(Throwable e) {
        if (warnings == null) {
            warnings = new ArrayList<String>();
            sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss:SSS");
        }
        StringBuffer msg = new StringBuffer("\nWARNING ");
        msg.append(warnings.size() + 1);
        msg.append("): ");
        msg.append(e.toString());
        StackTraceElement[] st = e.getStackTrace();
        for (int i = 0; i < st.length; i++) {
            msg.append("\n\tat ");
            msg.append(st[i].toString());
        }
        warnings.add(sdf.format(new Date()) + msg.toString());
    }



}
