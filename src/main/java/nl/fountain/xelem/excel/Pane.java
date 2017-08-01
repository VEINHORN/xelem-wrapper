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
package nl.fountain.xelem.excel;

import nl.fountain.xelem.Area;

/**
 * Represents the Pane element. 
 * <P>
 * Excel can split a window into a maximum of 4 panes. Panes come into vision
 * when you apply the 
 * {@link nl.fountain.xelem.excel.WorksheetOptions#splitHorizontal(int, int) splitHorizontal}-,
 * {@link nl.fountain.xelem.excel.WorksheetOptions#splitVertical(int, int) splitVertical}- or
 * {@link nl.fountain.xelem.excel.WorksheetOptions#freezePanesAt(int, int) freezePanesAt}-methods
 * of the {@link nl.fountain.xelem.excel.WorksheetOptions}.
 * 
 * 
 * @see nl.fountain.xelem.excel.WorksheetOptions#setActiveCell(int, int, int)
 * @see nl.fountain.xelem.excel.WorksheetOptions#setActiveCell(int, int)
 * @see nl.fountain.xelem.excel.WorksheetOptions#freezePanesAt(int, int)
 * @see nl.fountain.xelem.excel.WorksheetOptions#splitHorizontal(int, int)
 * @see nl.fountain.xelem.excel.WorksheetOptions#splitVertical(int, int)
 * @see nl.fountain.xelem.excel.WorksheetOptions#setRangeSelection(int, String)
 * @see nl.fountain.xelem.excel.WorksheetOptions#setRangeSelection(String)
 */
public interface Pane extends XLElement {
    
    /**
     * Variable indicating the top left pane. The top left pane is the standard
     * pane which is allways visible.
     */
    public static final int TOP_LEFT = 3;
    
    /**
     * Variable indicating the bottom left pane. The bottom left pane can
     * only be addressed if a window is split horizontally, either by splitting
     * or freezing panes.
     */
    public static final int BOTTOM_LEFT = 2;
    
    /**
     * Variable indicating the top right pane. The top right pane can 
     * only be addressed if a window is split vertically, either by splitting 
     * or freezing panes.
     */
    public static final int TOP_RIGHT = 1;
    
    /**
     * Variable indicating the bottom right pane. The bottom right pane can
     * only be addressed if a window is split vertically <em>and</em> horizontally,
     * either by spltting or freezing panes.
     */
    public static final int BOTTOM_RIGHT = 0;
    
    int getNumber();
    void setActiveCell(int row, int col);
    void setActiveCol(int col);
    int getActiveCol();
    void setActiveRow(int row);
    int getActiveRow();
    void setRangeSelection(String rc);
    void setRangeSelection(Area area);
    String getRangeSelection();

}
