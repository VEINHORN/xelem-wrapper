/*
 * Created on Sep 8, 2004
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
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import nl.fountain.xelem.Address;
import nl.fountain.xelem.CellPointer;
import nl.fountain.xelem.GIO;
import nl.fountain.xelem.excel.AbstractXLElement;
import nl.fountain.xelem.excel.AutoFilter;
import nl.fountain.xelem.excel.Cell;
import nl.fountain.xelem.excel.Column;
import nl.fountain.xelem.excel.NamedRange;
import nl.fountain.xelem.excel.Row;
import nl.fountain.xelem.excel.Table;
import nl.fountain.xelem.excel.Worksheet;
import nl.fountain.xelem.excel.WorksheetOptions;
import nl.fountain.xelem.excel.x.XAutoFilter;
import nl.fountain.xelem.excel.x.XWorksheetOptions;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.xml.sax.Attributes;

/**
 * An implementation of the XLElement Worksheet.
 * <P>
 * This class can be used to add rows, columns and cells to the worksheet.
 * The addCell-methods add a cell at the current position of the
 * worksheet's cellPointer (see {@link #getCellPointer()}) and move the 
 * cellPointer. If you whish to obtain the address of the newly added cell,
 * it should be obtained from the cellPointer prior to the add:
 * <PRE>
 *          Address adr = sheet.getCellPointer().getAddress();
 *          Cell cell = sheet.addCell();
 * </PRE>
 * The addCellAt-methods, which take an address, a string or a 
 * row/column pair of ints as
 * parameter, move the cellPointer to the specified position, add a cell and
 * leave the cellPointer pointing to the next position.
 */
public class SSWorksheet extends AbstractXLElement implements Worksheet {
    
    private CellPointer cellPointer;
    private String name;
    private boolean protect;
    private boolean righttoleft;
    private Map<String, NamedRange> namedRanges;
    private Table table;
    private WorksheetOptions options;
    private AutoFilter autoFilter;
    
    /**
     * Constructs a new SSWorksheet with the given name.
     * The worksheets cellpointer will be at position row 1, column 1.
     * 
     * @see nl.fountain.xelem.excel.Workbook#addSheet()
     */
    public SSWorksheet(String name) {
        this.name = name;
        cellPointer = new CellPointer();
    }

    public String getName() {
        return name;
    }
    
    public String getReferenceName() {
        if (name.indexOf(" ") > -1) {
            return "'" + name + "'";
        }
        return name;
    }
    
    public void setProtected(boolean p) {
        protect = p;
    }
    
    // method called by ExcelReader
    private void setProtected(String s) {
        protect = s.equals("1");
    }
    
    public boolean isProtected() {
        return protect;
    }
    
    public void setRightToLeft(boolean r) {
        righttoleft = r;
    }
    
    //  method called by ExcelReader
    private void setRightToLeft(String s) {
        righttoleft = s.equals("1");
    }
    
    public boolean isRightToLeft() {
        return righttoleft;
    }
    
    public NamedRange addNamedRange(NamedRange nr) {
        if (namedRanges == null) {
            namedRanges = new HashMap<String, NamedRange>();
        }
        namedRanges.put(nr.getName(), nr);
        return nr;
    }
    
    public NamedRange addNamedRange(String name, String refersTo) {
        return addNamedRange(new SSNamedRange(name, refersTo));
    }
    
    public Map<String, NamedRange> getNamedRanges() {
        if (namedRanges == null) {
            return Collections.emptyMap();
        } else {
            return namedRanges;
        }
    }
    
    public void setWorksheetOptions(WorksheetOptions wso) {
        options = wso;
    }
    
    public boolean hasWorksheetOptions() {
        return options != null;
    }

    public WorksheetOptions getWorksheetOptions() {
        if (options == null) {
            options = new XWorksheetOptions();
        }
        return options;
    }
    
    public void setTable(Table table) {
        this.table = table;
    }
    
    public Table getTable() {
        if (table == null) {
            table = new SSTable();
        }
        return table;
    }
    
    public boolean hasTable() {
        return table != null;
    }
    
    public CellPointer getCellPointer() {
        return cellPointer;
    }

    public Cell addCell() {
        return addCell(new SSCell());
    }
    
    public Cell addCell(Cell cell) {       
        int prevRow = cellPointer.getRowIndex();
        int prevColumn = cellPointer.getColumnIndex();
        cellPointer.move();
        getTable().getRowAt(prevRow).addCellAt(prevColumn, cell);
        return cell;
    }
    
    public Cell addCell(Object data) {
        Cell cell = addCell();
        cell.setData(data);
        return cell;
    }
    
    public Cell addCell(Object data, String styleID) {
        Cell cell = addCell(data);
        cell.setStyleID(styleID);
        return cell;
    }
    
    public Cell addCell(double data) {
        Cell cell = addCell();
        cell.setData(data);
        return cell;
    }
    
    public Cell addCell(double data, String styleID) {
        Cell cell = addCell(data);
        cell.setStyleID(styleID);
        return cell;
    }
    
    public Cell addCell(int data) {
        Cell cell = addCell();
        cell.setData(data);
        return cell;
    }
    
    public Cell addCell(int data, String styleID) {
        Cell cell = addCell(data);
        cell.setStyleID(styleID);
        return cell;
    }

    public Cell addCellAt(Address address) {
        return addCellAt(address.getRowIndex(), address.getColumnIndex());
    }
    
    public Cell addCellAt(int rowIndex, int columnIndex) {
        cellPointer.moveTo(rowIndex, columnIndex);
        return addCell();
    }
    
    public Cell addCellAt(String a1_ref) {
        return addCellAt(
                Address.calculateRow(a1_ref), Address.calculateColumn(a1_ref));
    }
    
    public Cell addCellAt(Address address, Cell cell) {
        return addCellAt(address.getRowIndex(), address.getColumnIndex(), cell);
    }
    
    public Cell addCellAt(int rowIndex, int columnIndex, Cell cell) {
        cellPointer.moveTo(rowIndex, columnIndex);
        return addCell(cell);
    }
    
    public Cell addCellAt(String a1_ref, Cell cell) {
        return addCellAt(
                Address.calculateRow(a1_ref), Address.calculateColumn(a1_ref), cell);
    }
    
    public Cell getCellAt(Address address) {
        return getCellAt(address.getRowIndex(), address.getColumnIndex());
    }
    
    public Cell getCellAt(int rowIndex, int columnIndex) {
        Row row = getTable().getRowAt(rowIndex);
        return row.getCellAt(columnIndex);
    }
    
    public Cell getCellAt(String a1_ref) {
        return getCellAt(
                Address.calculateRow(a1_ref), Address.calculateColumn(a1_ref));
    }
    
    public Cell removeCellAt(Address address) {
        return removeCellAt(address.getRowIndex(), address.getColumnIndex());
    }
    
    public Cell removeCellAt(int rowIndex, int columnIndex) {
        Row row = getTable().getRowAt(rowIndex);
        if (row == null) return null;
        return row.removeCellAt(columnIndex);
    }
    
    public Cell removeCellAt(String a1_ref) {
        return removeCellAt(
                Address.calculateRow(a1_ref), Address.calculateColumn(a1_ref));
    }
    
    public boolean hasCellAt(Address address) {
        return hasCellAt(address.getRowIndex(), address.getColumnIndex());
    }
    
    public boolean hasCellAt(int rowIndex, int columnIndex) {
        if (!hasTable()) {
            return false;
        } else if (!getTable().hasRowAt(rowIndex)) {
            return false;
        }
        return getTable().getRowAt(rowIndex).getCellAt(columnIndex) != null;
    }
    
    public boolean hasCellAt(String a1_ref) {
        return hasCellAt(
                Address.calculateRow(a1_ref), Address.calculateColumn(a1_ref));
    }

    public Row addRow() {
        return getTable().addRow();
    }

    public Row addRowAt(int rowIndex) {
        return getTable().addRowAt(rowIndex);
    }

    public Row addRow(Row row) {
        return getTable().addRow(row);
    }
    
    public Row addRowAt(int index, Row row) {
        return getTable().addRowAt(index, row);
    }

    public Row removeRowAt(int rowIndex) {
        return getTable().removeRowAt(rowIndex);
    }

    public Collection<Row> getRows() {
        return getTable().getRows();
    }

    public Row getRowAt(int rowIndex) {
        return getTable().getRowAt(rowIndex);
    }
    
    public boolean hasRowAt(int rowIndex) {
        if (!hasTable()) {
            return false;
        } else {
            return getTable().hasRowAt(rowIndex);
        }
    }
    
    public Column addColumn() {
        return getTable().addColumn();
    }
    
    public Column addColumnAt(int index) {
        return getTable().addColumnAt(index);
    }
    
    public Column addColumnAt(String label) {
        return getTable().addColumnAt(Address.calculateColumn(label));
    }
    
    public Column addColumn(Column column) {
        return getTable().addColumn(column);
    }
    
    public Column addColumnAt(int index, Column column) {
        return getTable().addColumnAt(index, column);
    }
    
    public Column addColumnAt(String label, Column column) {
        return getTable().addColumnAt(Address.calculateColumn(label));
    }
    
    public Column removeColumnAt(int columnIndex) {
        return getTable().removeColumnAt(columnIndex);
    }
    
    public Column removeColumnAt(String label) {
        return getTable().removeColumnAt(Address.calculateColumn(label));
    }
    
    public Collection<Column> getColumns() {
        return getTable().getColumns();
    }
    
    public Column getColumnAt(int columnIndex) {
        return getTable().getColumnAt(columnIndex);
    }
    
    public Column getColumnAt(String label) {
        return getTable().getColumnAt(Address.calculateColumn(label));
    }
    
    public boolean hasColumnAt(int columnIndex) {
        if (!hasTable()) {
            return false;
        } else {
            return getTable().hasColumnAt(columnIndex);
        }
    }
    
    public boolean hasColumnAt(String label) {
        return getTable().hasColumnAt(Address.calculateColumn(label));
    }
    
    public String getTagName() {
        return "Worksheet";
    }
    
    public String getNameSpace() {
        return XMLNS_SS;
    }
    
    public String getPrefix() {
        return PREFIX_SS;
    }
    
    public void setAutoFilter(AutoFilter af) {
        autoFilter = af;
    }
    
    public void setAutoFilter(String rcString) {
        autoFilter = new XAutoFilter();
        autoFilter.setRange(rcString);
    }
    
    public boolean hasAutoFilter() {
        return autoFilter != null;
    }
    
    public void removeAutoFilter() {
        autoFilter = null;
    }
    
    public Element assemble(Element parent, GIO gio) {
        Document doc = parent.getOwnerDocument();
        Element wse = assemble(doc, gio);
        
        wse.setAttributeNodeNS(createAttributeNS(doc, "Name", getName()));
        if (protect) wse.setAttributeNodeNS(
                createAttributeNS(doc, "Protected", "1"));
        if (righttoleft) wse.setAttributeNodeNS(
                createAttributeNS(doc, "RightToLeft", "1"));
        
        parent.appendChild(wse);
        
        // Names
        if (namedRanges != null) {
            Element names = doc.createElement("Names");
            wse.appendChild(names);
            for (NamedRange nr : namedRanges.values()){
                nr.assemble(names, gio);
            }
        }
        
        if (hasTable()) {
            getTable().assemble(wse, gio);
        }
        if (hasWorksheetOptions()) {
            getWorksheetOptions().assemble(wse, gio);
        }
        if (hasAutoFilter()) {
            autoFilter.assemble(wse, gio);
        }
        return wse;
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

}
