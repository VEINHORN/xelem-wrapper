/*
 * Created on Nov 4, 2004
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
package nl.fountain.xelem.excel.x;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.excel.AbstractXLElement;
import nl.fountain.xelem.excel.AutoFilter;

/**
 * An implementation of the XLElement AutoFilter.
 * 
 * @see nl.fountain.xelem.excel.Worksheet#setAutoFilter(String)
 */
public class XAutoFilter extends AbstractXLElement implements AutoFilter {
    
    private String range;
    
    /**
     * Constructs a new XAutoFilter.
     * 
     * @see nl.fountain.xelem.excel.Worksheet#setAutoFilter(String)
     */
    public XAutoFilter() {}

    public void setRange(String rcString) {
        range = rcString;
    }
    
    public String getRange() {
        return range;
    }

    public String getTagName() {
        return "AutoFilter";
    }

    public String getNameSpace() {
        return XMLNS_X;
    }

    public String getPrefix() {
        return PREFIX_X;
    }

    public Element assemble(Element parent, GIO gio) {
        if (getRange() != null) {
            Document doc = parent.getOwnerDocument();
	        Element afe = assemble(doc, gio);
	        afe.setAttributeNodeNS(createAttributeNS(doc, "Range", getRange()));
	        parent.appendChild(afe);
	        return afe;
        } else {
            return null;
        }
    }

}
