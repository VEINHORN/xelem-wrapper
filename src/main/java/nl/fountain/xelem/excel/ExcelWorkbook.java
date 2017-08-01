/*
 * Created on Sep 7, 2004
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


/**
 * Represents the ExcelWorkbook element.
 */
public interface ExcelWorkbook extends XLElement {
    
    void setWindowHeight(int height);
    int getWindowHeight();
    
    void setWindowWidth(int width);
    int getWindowWidth();
    
    void setWindowTopX(int x);
    int getWindowTopX();
    
    void setWindowTopY(int y);
    int getWindowTopY();
    
    /**
     * Set the active sheet. 0-based.
     * @param nr the index of the active sheet.
     */
    void setActiveSheet(int nr);
    
    /**
     * Gets the active sheet. 0 based.
     * @return an int >= 0
     */
    int getActiveSheet();
    
    void setProtectStructure(boolean protect);
    boolean getProtectStructure();
    
    void setProtectWindows(boolean protect);
    boolean getProtectWindows();
    
}
