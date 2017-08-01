/*
 * Created on Oct 15, 2004
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
 * Represents the WorksheetOptions element.
 */
public interface WorksheetOptions extends XLElement {
    
    public static final String SHEET_VISIBLE = "SheetVisible";
    public static final String SHEET_HIDDEN = "SheetHidden";
    public static final String SHEET_VERY_HIDDEN = "SheetVeryHidden";
    
    void setTopRowVisible(int tr);
    int getTopRowVisible();
    void setLeftColumnVisible(int lc);
    int getLeftColumnVisible();
    
    void setZoom(int z);
    int getZoom();
    
    void setGridlineColor(int r, int g, int b);
    String getGridlineColor();
    
    void setTabColorIndex(int ci);
    int getTabColorIndex();
    
    void setSelected(boolean s);
    boolean isSelected();
    
    void doNotDisplayHeadings(boolean b);
    boolean displaysNoHeadings();
    
    void doNotDisplayGridlines(boolean b);
    boolean displaysNoGridlines();
    
    void doDisplayFormulas(boolean f);
    boolean displaysFormulas();
    
    void setVisible(String wsoValue);
    String getVisible();
    
    void setActiveCell(int r, int c);
    void setActiveCell(int paneNumber, int r, int c);
    
    /**
     * Sets a range selection. The active cell must also be set within
     * the range in order for this method to have effect.
     * 
     * @param rcRange A String in R1C1 reference style.
     * @see #setActiveCell(int, int)
     */
    void setRangeSelection(String rcRange);
    void setRangeSelection(Area area);
    
    void setRangeSelection(int paneNumber, String rcRange);
    
    void splitHorizontal(int points, int topRow);
    void splitVertical(int points, int leftColumn);
    
    boolean hasHorizontalSplit();
    boolean hasVerticalSplit();
    boolean hasSplit();
    
    void freezePanesAt(int r, int c);
    boolean hasFrozenPanes();
    
}
