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

import java.util.ArrayList;
import java.util.List;

import nl.fountain.xelem.Area;
import nl.fountain.xelem.excel.Worksheet;



/**
 * Supervises the process of reading Excel workbooks.
 * 
 * @since xelem.2.0
 */
class Director {
    
    private XLWorkbookBuilder xlworkbookbuilder;
    private SSWorksheetBuilder ssworksheetbuilder;
    private SSRowBuilder ssrowbuilder;
    private SSCellBuilder sscellbuilder;
    private List<AnonymousBuilder> anonymousBuilders;
    private Area buildArea;
    private List<ExcelReaderListener> listeners;
    
    private int currentSheetIndex;
    private String currentSheetName;
    private int currentRowIndex;
    
    public void setBuildArea(Area area) {
        buildArea = area;
    }
    
    public Area getBuildArea() {
        if (buildArea == null) {
            buildArea = new Area(Worksheet.firstRow, Worksheet.firstColumn,
                    Worksheet.lastRow, Worksheet.lastColumn);
        }
        return buildArea;
    }
    
    public boolean hasBuildArea() {
        return buildArea != null;
    }
    
    public List<ExcelReaderListener> getListeners() {
        if (listeners == null) {
            listeners = new ArrayList<ExcelReaderListener>();
        }
        return listeners;
    }
    
    public void addExcelReaderListener(ExcelReaderListener l) {
        if (!getListeners().contains(l)) {
            getListeners().add(l);
        }
    }
    
    public boolean removeExcelReaderListener(ExcelReaderListener l) {
        return getListeners().remove(l);
    }
    
    public void clearExcelReaderListeners() {
        getListeners().clear();
    }
    
    public XLWorkbookBuilder getXLWorkbookBuilder() {
        if (xlworkbookbuilder == null) {
            xlworkbookbuilder = new XLWorkbookBuilder(this);
        }
        return xlworkbookbuilder;
    }
    
    public SSWorksheetBuilder getSSWorksheetBuilder() {
        if (ssworksheetbuilder == null) {
            ssworksheetbuilder = new SSWorksheetBuilder(this);
        }
        return ssworksheetbuilder;
    }
    
    public SSRowBuilder getSSRowBuilder() {
        if (ssrowbuilder == null) {
            ssrowbuilder = new SSRowBuilder(this);
        }
        return ssrowbuilder;
    }
    
    public SSCellBuilder getSSCellBuilder() {
        if (sscellbuilder == null) {
            sscellbuilder = new SSCellBuilder(this);
        }
        return sscellbuilder; 
    }
    
    public AnonymousBuilder getAnonymousBuilder() {
        AnonymousBuilder aBuilder = null;
        for (AnonymousBuilder builder : getBuilders()) {
            if (!builder.isOccupied()) {
                aBuilder = builder;
                break;
            }
        }
        if (aBuilder == null) {
            aBuilder = new AnonymousBuilder(this);
            getBuilders().add(aBuilder);
        }
        aBuilder.setOccupied(true);
        return aBuilder;
    }
    
    void setCurrentSheetIndex(int index) {
        currentSheetIndex = index;
    }
    
    int getCurrentSheetIndex() {
        return currentSheetIndex;
    }
    
    void setCurrentSheetName(String name) {
        currentSheetName = name;
    }
    
    String getCurrentSheetName() {
        return currentSheetName;
    }
    
    void setCurrentRowIndex(int index) {
        currentRowIndex = index; 
    }
    
    int getCurrentRowIndex() {
        return currentRowIndex;
    }
    
    private List<AnonymousBuilder> getBuilders() {
        if (anonymousBuilders == null) {
            anonymousBuilders = new ArrayList<AnonymousBuilder>();
        }
        return anonymousBuilders;
    }
    
    

}
