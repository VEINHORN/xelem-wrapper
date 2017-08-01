/*
 * Created on 30-okt-2004
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

import java.lang.reflect.Method;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.excel.AbstractXLElement;
import nl.fountain.xelem.excel.ExcelWorkbook;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

/**
 * An implementation of the XLElement ExcelWorkbook.
 */
public class XExcelWorkbook extends AbstractXLElement implements ExcelWorkbook {
    
    private int windowHeight;
    private int windowWidth;
    private int windowTopX;
    private int windowTopY;
    private int activeSheet = -1;
    private boolean protectstructure;
    private boolean protectwindows;
    
    /**
     * Creates a new XExcelWorkbook.
     * 
     * @see nl.fountain.xelem.excel.Workbook#getExcelWorkbook()
     */
    public XExcelWorkbook() {}
    
    public void setWindowHeight(int height) {
        windowHeight = height;
    }
    
    // method called by ExcelReader
    private void setWindowHeight(String s) {
        windowHeight = Integer.parseInt(s);
    }

    public int getWindowHeight() {
        return windowHeight;
    }

    public void setWindowWidth(int width) {
        windowWidth = width;
    }
    
    //  method called by ExcelReader
    private void setWindowWidth(String s) {
       windowWidth = Integer.parseInt(s); 
    }

    public int getWindowWidth() {
        return windowWidth;
    }

    public void setWindowTopX(int x) {
        windowTopX = x;
    }
    
    //  method called by ExcelReader
    private void setWindowTopX(String s) {
       windowTopX = Integer.parseInt(s);
    }

    public int getWindowTopX() {
        return windowTopX;
    }

    public void setWindowTopY(int y) {
        windowTopY = y;
    }
    
    //  method called by ExcelReader
    private void setWindowTopY(String s) {
        windowTopY = Integer.parseInt(s);
    }

    public int getWindowTopY() {
        return windowTopY;
    }
    
    public void setActiveSheet(int nr) {
        activeSheet = nr;
    }
    
    //  method called by ExcelReader
    private void setActiveSheet(String s) {
        activeSheet = Integer.parseInt(s);
    }
    
	public int getActiveSheet() {
	    if (activeSheet < 0) return 0;
	    return activeSheet;
	}
	
	public void setProtectStructure(boolean protect) {
	    protectstructure = protect;
    }
	
	//	 method called by ExcelReader
	private void setProtectStructure(String s) {
	    protectstructure = s.equalsIgnoreCase("true");
	}
	
	public boolean getProtectStructure() {
        return protectstructure;
    }
	
	public void setProtectWindows(boolean protect) {
	    protectwindows = protect;
    }
	
	//	 method called by ExcelReader
	private void setProtectWindows(String s) {
	    protectwindows = s.equalsIgnoreCase("true");
	}
	
	public boolean getProtectWindows() {
        return protectwindows;
    }

    public String getTagName() {
        return "ExcelWorkbook";
    }
    
    public String getNameSpace() {
        return XMLNS_X;
    }
    
    public String getPrefix() {
        return PREFIX_X;
    }
    
    public Element assemble(Element parent, GIO gio) {
        Document doc = parent.getOwnerDocument();
        Element ewbe = assemble(doc, gio);
        
        if (windowHeight > 0)
            ewbe.appendChild(createElementNS(doc, "WindowHeight", windowHeight));
        if (windowWidth > 0)
            ewbe.appendChild(createElementNS(doc, "WindowWidth", windowWidth));
        if (windowTopX != 0)
            ewbe.appendChild(createElementNS(doc, "WindowTopX", windowTopX));
        if (windowTopY != 0)
            ewbe.appendChild(createElementNS(doc, "WindowTopY", windowTopY));
        if (activeSheet > -1)
            ewbe.appendChild(createElementNS(doc, "ActiveSheet", activeSheet));
        ewbe.appendChild(createElementNS(doc, "ProtectStructure", protectstructure));
        ewbe.appendChild(createElementNS(doc, "ProtectWindows", protectwindows));
        
        parent.appendChild(ewbe);
        return ewbe;
    }
    
    public void setChildElement(String localName, String content) {
        invokeMethod(localName, content);
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
