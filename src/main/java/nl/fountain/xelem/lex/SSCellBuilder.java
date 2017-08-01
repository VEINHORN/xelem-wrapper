/*
 * Created on 22-mrt-2005
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

import nl.fountain.xelem.excel.Comment;
import nl.fountain.xelem.excel.XLElement;
import nl.fountain.xelem.excel.ss.SSCell;

import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 *
 */
class SSCellBuilder extends AnonymousBuilder {
    
    private SSCell current;
    
    SSCellBuilder(Director director) {
        super(director);
    }
    
    public void build(XMLReader reader, ContentHandler parent, XLElement xle) {
        setUpBuilder(reader, parent);
        current = (SSCell) xle;
    }
    
    public void startElement(String uri, String localName, String qName,
            Attributes atts) throws SAXException {
        contents.reset();
        if (XLElement.XMLNS_SS.equals(uri)) {
            if ("Data".equals(localName)) {
	            // set the atts of the data element
	            current.setAttributes(atts);
            } else if ("Comment".equals(localName)) {
                Comment comment = current.addComment();
                comment.setAttributes(atts);
                Builder builder = director.getAnonymousBuilder();
                builder.build(reader, this, comment);
            }
        }
    }
    
    public void endElement(String uri, String localName, String qName) throws SAXException {
        if (current.getNameSpace().equals(uri)) {
            if (current.getTagName().equals(localName)) {
            	for (ExcelReaderListener listener : director.getListeners()) {
                    listener.setCell(director.getCurrentSheetIndex(),
                            director.getCurrentSheetName(), 
                            director.getCurrentRowIndex(), current);
                }
                reader.setContentHandler(parent);
                return;
            }
        }
        current.setChildElement(localName, contents.toString());
    }

}
