/*
 * Created on 5-apr-2005
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

import nl.fountain.xelem.excel.Cell;
import nl.fountain.xelem.excel.Row;
import nl.fountain.xelem.excel.XLElement;
import nl.fountain.xelem.excel.ss.SSRow;

import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 *
 */
class SSRowBuilder extends AnonymousBuilder {
    
    private Row currentRow;
    private int currentCellIndex;
    
    SSRowBuilder(Director director) {
        super(director);
    }
    
    public void build(XMLReader reader, ContentHandler parent, XLElement xle) {
        setUpBuilder(reader, parent);
        currentRow = (SSRow) xle;
        currentCellIndex = 0;
    }
    
    public void startElement(String uri, String localName, String qName,
            Attributes atts) throws SAXException {
        if (XLElement.XMLNS_SS.equals(uri)) {
            if ("Cell".equals(localName)) {
                String index = atts.getValue(XLElement.XMLNS_SS, "Index");
                if (index != null) {
                    currentCellIndex = Integer.parseInt(index);
                } else {
                    currentCellIndex++;
                }
                if (director.getBuildArea().isColumnPartOfArea(currentCellIndex)) {
	                Cell cell = currentRow.addCellAt(currentCellIndex);
	                cell.setIndex(currentCellIndex);
	                cell.setAttributes(atts);
	                Builder builder = director.getSSCellBuilder();
	                builder.build(reader, this, cell);
                }
            }
        }
    }
    
    public void endElement(String uri, String localName, String qName) throws SAXException {
        if (currentRow.getTagName().equals(localName)) {
            if (currentRow.getNameSpace().equals(uri)) {
            	for (ExcelReaderListener listener : director.getListeners()) {
                    listener.setRow(director.getCurrentSheetIndex(),
                            director.getCurrentSheetName(), currentRow);
                }
                reader.setContentHandler(parent);
                return;
            }
        }
        // no child elements?
    }

}
