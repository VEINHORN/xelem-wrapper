/*
 * Created on 31-mrt-2005
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
package nl.fountain.xelem.excel.ss;

import java.lang.reflect.Method;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.excel.AbstractXLElement;
import nl.fountain.xelem.excel.Comment;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.xml.sax.Attributes;

/**
 * An implementation of the XLElement Comment.
 */
public class SSComment extends AbstractXLElement implements Comment {
    
    private String author;
    private boolean showAlways;
    private String data;

    public void setAuthor(String author) {
        this.author = author;
    }

    public String getAuthor() {
        return author;
    }

    public void setShowAlways(boolean show) {
        showAlways = show;
    }
    
    private void setShowAlways(String s) {
        showAlways = "1".equals(s);
    }

    public boolean showsAlways() {
        return showAlways;
    }

    public void setData(String data) {
        this.data = data;
    }

    public String getData() {
        return data;
    }
    
    public String getDataClean() {
        if (data != null) {            
            String s = data;
            if (author != null) {
                s = s.replaceFirst(author + ":", "");
            }
            return s.trim();
        } else {
            return data;
        }
    }

    public String getTagName() {
        return "Comment";
    }

    public String getNameSpace() {
        return XMLNS_SS;
    }

    public String getPrefix() {
        return PREFIX_SS;
    }

    public Element assemble(Element parent, GIO gio) {
        if (data == null) return null;
        
        Document doc = parent.getOwnerDocument();
        Element coe = assemble(doc, gio);
        
        if (author != null) coe.setAttributeNodeNS(
                createAttributeNS(doc, "Author", author));
        if (showAlways) coe.setAttributeNodeNS(
                createAttributeNS(doc, "ShowAlways", "1"));
        
        Element dae = doc.createElementNS(XMLNS_SS, "Data");
        dae.setPrefix(PREFIX_SS);
        dae.appendChild(doc.createTextNode(data));
        coe.appendChild(dae);
        
        parent.appendChild(coe);
        
        return coe;
    }
    
    public void setAttributes(Attributes attrs) {
        for (int i = 0; i < attrs.getLength(); i++) {
            invokeMethod(attrs.getLocalName(i), attrs.getValue(i));
        }
    }
    
    public void setChildElement(String localName, String content) {
        //System.out.println(localName+"="+content);
        setData(content);
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

}
