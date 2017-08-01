/*
 * Created on 31-okt-2004
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
package nl.fountain.xelem.excel;

import java.util.List;

import nl.fountain.xelem.GIO;

import org.w3c.dom.Element;

import org.xml.sax.Attributes;


/**
 * Represents an element in SpreadsheetML. 
 * <P>
 * XLElements are unaware of their parent. This makes it possible to move and
 * duplicate them to other, allthough appropriate, XLElements of the Workbook.
 * Even to other Workbooks.
 * <P>
 * An XLElement is capable of assembling it's state into an
 * {@link org.w3c.dom.Element} and attach this Element to a parent-element in a
 * {@link org.w3c.dom.Document}.
 * 
 */
public interface XLElement {
    
    /**
     * The default SpreadsheetML namespace: {@value}
     */
    public static final String XMLNS="urn:schemas-microsoft-com:office:spreadsheet";
    
    /**
     * The Office namespace.
     */
    public static final String XMLNS_O="urn:schemas-microsoft-com:office:office";
    
    /**
     * The Excel namespace.
     */
    public static final String XMLNS_X="urn:schemas-microsoft-com:office:excel";
    
    /**
     * The SpreadsheetML namespace.
     */
    public static final String XMLNS_SS="urn:schemas-microsoft-com:office:spreadsheet";
    
    /**
     * HTML-namespace.
     */
    public static final String XMLNS_HTML="http://www.w3.org/TR/REC-html40";
    
    /**
     * The Office namespace prefix: {@value}.
     */
    public static final String PREFIX_O = "o";
    
    /**
     * The Excel namespace prefix: {@value}.
     */
    public static final String PREFIX_X = "x";
    
    /**
     * The SpreadsheetML namespace prefix: {@value}.
     */
    public static final String PREFIX_SS = "ss";
    
    /**
     * HTML-namespace prefix: {@value}.
     */
    public static final String PREFIX_HTML = "html";
    
    /**
     * Gets the tagname (localName) of this XLElement.
     * @return qualified name
     */
    String getTagName();
    
    /**
     * Gets the namespace of this XLElement.
     * @return namespace
     */
    String getNameSpace();
    
    /**
     * Gets the prefix of this XLElement.
     * @return prefix
     */
    String getPrefix();
    
    /**
     * Adds a comment to this XLElement.
     * @param comment	The string to be printed in xml as a comment to the element.
     * @see nl.fountain.xelem.excel.Workbook#setPrintElementComments(boolean print)
     */
    void addElementComment(String comment);
    
    /**
     * Gets a List of added comments.
     * @return List of Strings
     */
    List getElementComments();   
    
    /**
     * Assembles the state of this XLElement and all of it's children into an
     * {@link org.w3c.dom.Element}. Attaches the new element to it's
     * parent-element if needed.
     * 
     * @param 	parent the parent-element to which the new formed element 
     * 			will be appended
     * @param 	gio a global information object
     * 
     * @return	the newly assembled element. may be null.
     */
    Element assemble(Element parent, GIO gio);
    
    /**
     * Called by a builder during a read by {@link nl.fountain.xelem.lex.ExcelReader}.
     * Sets the (xml-)attributes found on the (xml-)element.
     * 
     * @param attrs the attributes found on the element
     * @since xelem.2.0
     */
    void setAttributes(Attributes attrs);
    
    /**
     * Called by a builder during a read by {@link nl.fountain.xelem.lex.ExcelReader}.
     * Sets a child element found on the (xml-)element.
     * 
     * @param localName	the local name of the child
     * @param content the content of the child
     * @since xelem.2.0
     */
    void setChildElement(String localName, String content);
}
