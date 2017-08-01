/*
 * Created on Nov 3, 2004
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

import nl.fountain.xelem.Area;
import nl.fountain.xelem.GIO;
import nl.fountain.xelem.excel.AbstractXLElement;
import nl.fountain.xelem.excel.Pane;

/**
 * An implementation of the XLElement Pane. Methods of this class are mainly used
 * by {@link nl.fountain.xelem.excel.WorksheetOptions}-methods.
 * 
 * @see nl.fountain.xelem.excel.WorksheetOptions#setActiveCell(int, int, int)
 * @see nl.fountain.xelem.excel.WorksheetOptions#setActiveCell(int, int)
 * @see nl.fountain.xelem.excel.WorksheetOptions#freezePanesAt(int, int)
 * @see nl.fountain.xelem.excel.WorksheetOptions#splitHorizontal(int, int)
 * @see nl.fountain.xelem.excel.WorksheetOptions#splitVertical(int, int)
 * @see nl.fountain.xelem.excel.WorksheetOptions#setRangeSelection(int, String)
 * @see nl.fountain.xelem.excel.WorksheetOptions#setRangeSelection(String)
 */
public class XPane extends AbstractXLElement implements Pane {
    
    private int number;
    private int activeCol = -1;
    private int activeRow = -1;
    private String rangeSelection;
    
    /**
     * Constructs a new XPane with the given pane number.
     * 
     * @throws java.lang.IllegalArgumentException if paneNumber < 0 or
     * 			paneNumber > 3.
     */
    public XPane(int paneNumber) {
        if (paneNumber < 0 || paneNumber > 3) {
            throw new IllegalArgumentException(paneNumber 
                    + ". Legal arguments are 0, 1, 2 and 3.");
        }
        number = paneNumber;
    }

    public int getNumber() {
        return number;
    }

    public void setActiveCol(int col) {
        activeCol = col - 1;
    }

    public int getActiveCol() {
        return activeCol;
    }

    public void setActiveRow(int row) {
        activeRow = row - 1;
    }

    public int getActiveRow() {
        return activeRow;
    }
    
    public void setActiveCell(int row, int col) {
        setActiveRow(row);
        setActiveCol(col);
    }

    public void setRangeSelection(String rc) {
        rangeSelection = rc;
    }
    
    public void setRangeSelection(Area area) {
        rangeSelection = area.getAbsoluteRange();
    }

    public String getRangeSelection() {
        return rangeSelection;
    }

    public String getTagName() {
        return "Pane";
    }

    public String getNameSpace() {
        return XMLNS_X;
    }

    public String getPrefix() {
        return PREFIX_X;
    }

    public Element assemble(Element parent, GIO gio) {
        Document doc = parent.getOwnerDocument();
        Element pe = assemble(doc, gio);
        
        pe.appendChild(createElementNS(doc, "Number", number));
        if (activeCol > -1) pe.appendChild(
                createElementNS(doc, "ActiveCol", activeCol));
        if (activeRow > -1) pe.appendChild(
                createElementNS(doc, "ActiveRow", activeRow));
        if (rangeSelection != null) pe.appendChild(
                createElementNS(doc, "RangeSelection", rangeSelection));
        
        parent.appendChild(pe);
        return pe;
    }

}
