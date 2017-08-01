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

import java.io.CharArrayWriter;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import nl.fountain.xelem.excel.DocumentProperties;
import nl.fountain.xelem.excel.ExcelWorkbook;
import nl.fountain.xelem.excel.WorksheetOptions;
import nl.fountain.xelem.excel.XLElement;

import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

/**
 *
 */
class AnonymousBuilder extends DefaultHandler implements Builder {
    
    private static Map<String, Method> methodMap;
    private boolean occupied;
    protected XMLReader reader;
    protected ContentHandler parent;
    protected Director director;
    protected CharArrayWriter contents;
    protected XLElement current;
    
    protected AnonymousBuilder(Director director) {
        this.director = director;
        contents = new CharArrayWriter();
    }
    
    public void build(XMLReader reader, ContentHandler parent, XLElement xle) {
        setUpBuilder(reader, parent);
        current = xle;
    }
    
    public void build(XMLReader reader, ContentHandler parent) {
        setUpBuilder(reader, parent);
    }
    
    protected void setUpBuilder(XMLReader reader, ContentHandler parent) {
        this.reader = reader;
        this.parent = parent;
        reader.setContentHandler(this);
    }
    
    protected void setOccupied(boolean b) {
        occupied = b;
    }
    
    protected boolean isOccupied() {
        return occupied;
    }
    
    private Method getMethod(String tagName) {
        return getMethodMap().get(tagName);
    }
    
    private Map<String, Method> getMethodMap() {
        if (methodMap == null) {
            methodMap = new HashMap<String, Method>();
            Object[][] methods = null;
            try {
                methods = new Object[][]{
                   {"DocumentProperties",  ExcelReaderListener.class.getMethod("setDocumentProperties", new Class[]{DocumentProperties.class})},
                   {"ExcelWorkbook", ExcelReaderListener.class.getMethod("setExcelWorkbook", new Class[]{ExcelWorkbook.class})},
                   {"WorksheetOptions", ExcelReaderListener.class.getMethod("setWorksheetOptions", new Class[]{int.class, String.class, WorksheetOptions.class})}
                };
                for (int i = 0; i < methods.length; i++) {
                    methodMap.put((String)methods[i][0], (Method)methods[i][1]);
                }
            } catch (SecurityException e) {
                e.printStackTrace();
            } catch (NoSuchMethodException e) {
                e.printStackTrace();
            }
        }
        return methodMap;
    }
    
    private void informListeners() {
        if (director.getListeners().size() > 0) {
            Method m = getMethod(current.getTagName());
            if (m != null) {
	            try {
	                for (Iterator<ExcelReaderListener> iter = director.getListeners().iterator(); iter.hasNext();) {
	                    ExcelReaderListener listener = iter.next();
	                    if (m.getParameterTypes()[0].equals(int.class)) {
	                        m.invoke(listener, new Object[] {new Integer(director.getCurrentSheetIndex()), director.getCurrentSheetName(), current});
	                    } else {
	                        m.invoke(listener, new Object[] { current });
	                    }
	                }
	            } catch (Exception e) {
	                e.printStackTrace();
	            } 
            }
        }
    }

    public void startElement(String uri, String localName, String qName,
            Attributes atts) throws SAXException {
        //System.out.println(localName);
        if (!XLElement.XMLNS_HTML.equals(uri)) {
            contents.reset();
        }
    }

    public void endElement(String uri, String localName, String qName)
            throws SAXException {
        //System.out.println(localName);
        if (current.getNameSpace().equals(uri)) {	        
            if (current.getTagName().equals(localName)) {
                informListeners();
                reader.setContentHandler(parent);
                occupied = false;
                return;
            }
        }
        current.setChildElement(localName, contents.toString());
    }

    public void characters(char[] ch, int start, int length)
		throws SAXException {
        contents.write(ch, start, length);
    }
    
}
