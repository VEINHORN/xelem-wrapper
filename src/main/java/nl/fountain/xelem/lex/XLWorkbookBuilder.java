/*
 * Created on 15-mrt-2005
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
 */
package nl.fountain.xelem.lex;

import nl.fountain.xelem.excel.DocumentProperties;
import nl.fountain.xelem.excel.ExcelWorkbook;
import nl.fountain.xelem.excel.NamedRange;
import nl.fountain.xelem.excel.Worksheet;
import nl.fountain.xelem.excel.XLElement;
import nl.fountain.xelem.excel.o.ODocumentProperties;
import nl.fountain.xelem.excel.ss.SSNamedRange;
import nl.fountain.xelem.excel.ss.SSWorksheet;
import nl.fountain.xelem.excel.x.XExcelWorkbook;

import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 *
 */
class XLWorkbookBuilder extends AnonymousBuilder {
    
    private int sheetCounter;
    
    XLWorkbookBuilder(Director director) {
        super(director);
    }
    
    public void build(XMLReader reader, ContentHandler parent) {
        super.build(reader, parent);
        sheetCounter = 0;
    }
            
    public void startElement(String uri, String localName, String qName,
            Attributes atts) throws SAXException {
        
        // DocumentProperties
        if (XLElement.XMLNS_O.equals(uri) && "DocumentProperties".equals(localName)) {            
            DocumentProperties docprops = new ODocumentProperties();
            Builder builder = director.getAnonymousBuilder();
            builder.build(reader, this, docprops);
        
        // ExcelWorkbook
        } else if (XLElement.XMLNS_X.equals(uri) && "ExcelWorkbook".equals(localName)) {
            ExcelWorkbook xlwb = new XExcelWorkbook();
            Builder builder = director.getAnonymousBuilder();
            builder.build(reader, this, xlwb);
                       
        // Styles
            
        // NamedRange
        } else if (XLElement.XMLNS_SS.equals(uri) && "NamedRange". equals(localName)) {
            NamedRange nr = new SSNamedRange(
                    atts.getValue(XLElement.XMLNS_SS, "Name"), null);
            nr.setAttributes(atts);
            for (ExcelReaderListener listener : director.getListeners()) {
                listener.setNamedRange(nr);
            }
            
        // Worksheet
        } else if (XLElement.XMLNS_SS.equals(uri) && "Worksheet". equals(localName)) {
            String sheetName = atts.getValue(XLElement.XMLNS_SS, "Name");
            Worksheet sheet = new SSWorksheet(sheetName);
            director.setCurrentSheetIndex(sheetCounter);
            director.setCurrentSheetName(sheetName);
            sheet.setAttributes(atts);
            for (ExcelReaderListener listener : director.getListeners()) {
                listener.startWorksheet(sheetCounter, sheet);
            }
            
            sheetCounter++;
            Builder builder = director.getSSWorksheetBuilder();
            builder.build(reader, this);
        }
    }
    
    
    public void endElement(String uri, String localName, String qName) throws SAXException {
        if (XLElement.XMLNS_SS.equals(uri)) {
            if ("Workbook".equals(localName)) {
                reader.setContentHandler(parent);
                return;
            }
        }
        // no child elements
    }
    

}
