/*
 * Created on Oct 28, 2004
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

import java.lang.reflect.Method;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.excel.AbstractXLElement;
import nl.fountain.xelem.excel.Column;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.xml.sax.Attributes;

/**
 * An implementation of the XLElement Column.
 */
public class SSColumn extends AbstractXLElement implements Column {
    
    private int idx;
    private String styleID;
    private int span;
    private double width;
    private boolean hidden;
    private boolean autoFitWidth = true;
    
    /**
     * Creates a new SSColumn.
     * 
     * @see nl.fountain.xelem.excel.Worksheet#addColumn()
     */
    public SSColumn() {}

    public void setStyleID(String id) {
        styleID = id;
    }

    public String getStyleID() {
        return styleID;
    }
    
    public void setAutoFitWidth(boolean autoFit) {
        autoFitWidth = autoFit;
    }
    
    //method called by ExcelReader
    private void setAutoFitWidth(String s) {
        autoFitWidth = s.equals("1");
    }
    
    public boolean getAutoFitWith() {
        return autoFitWidth;
    }
    
    public void setSpan(int s) {
        span = s;
    }
    
    //method called by ExcelReader
    private void setSpan(String s) {
        span = Integer.parseInt(s);
    }
    
    public int getSpan() {
        return span;
    }
    
    public void setWidth(double w) {
        width = w;
    }
    
    // method called by ExcelReader
    private void setWidth(String s) {
        width = Double.parseDouble(s);
    }
    
    public double getWidth() {
        return width;
    }
    
    public void setHidden(boolean hide) {
        hidden = hide;
    }
    
    public boolean isHidden() {
        return hidden;
    }
    
    private void setHidden(String s) {
        hidden = "1".equals(s);
    }
    
    public String getTagName() {
        return "Column";
    }
    
    public String getNameSpace() {
        return XMLNS_SS;
    }
    
    public String getPrefix() {
        return PREFIX_SS;
    }
    
    public Element assemble(Element parent, GIO gio) {
        Document doc = parent.getOwnerDocument();
        Element ce = assemble(doc, gio);
        
        if (idx != 0) ce.setAttributeNodeNS(
                createAttributeNS(doc, "Index", idx));
        if (getStyleID() != null) {
            ce.setAttributeNodeNS(createAttributeNS(doc, "StyleID", getStyleID()));
            gio.addStyleID(getStyleID());
        }
        
        if (span > 0) ce.setAttributeNodeNS(
                createAttributeNS(doc, "Span", span));
        if (width > 0.0) ce.setAttributeNodeNS(
                createAttributeNS(doc, "Width", "" + width));
        if (hidden) ce.setAttributeNodeNS(
                createAttributeNS(doc, "Hidden", "1"));
        if (!autoFitWidth) ce.setAttributeNodeNS(
                createAttributeNS(doc, "AutoFitWidth", "0"));
        
        parent.appendChild(ce);
        return ce;
    }
    
    public void setAttributes(Attributes attrs) {
        for (int i = 0; i < attrs.getLength(); i++) {
            invokeMethod(attrs.getLocalName(i), attrs.getValue(i));
        }
    }
    
	private void invokeMethod(String name, Object value) {
        Class[] types = new Class[] { value.getClass() };
        Method method = null;
        try {
            method = this.getClass().getDeclaredMethod("set" + name, types);
            method.invoke(this, new Object[] { value });
        } catch (NoSuchMethodException e) {
            // no big deal
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Sets the value of the ss:Index-attribute of this Column-element. This method is 
     * called by {@link nl.fountain.xelem.excel.Table#columnIterator()} to set the
     * index of this column during assembly.
     */
    public void setIndex(int index) {
        idx = index;
    }
    
    public int getIndex() {
        return idx;
    }

}
