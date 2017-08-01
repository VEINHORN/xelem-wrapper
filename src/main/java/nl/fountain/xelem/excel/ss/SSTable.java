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
import java.util.Iterator;
import java.util.TreeMap;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.excel.AbstractXLElement;
import nl.fountain.xelem.excel.Column;
import nl.fountain.xelem.excel.Row;
import nl.fountain.xelem.excel.Table;
import nl.fountain.xelem.excel.Worksheet;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.xml.sax.Attributes;

/**
 * An implementation of the XLElement Table.
 */
public class SSTable extends AbstractXLElement implements Table {
    
    TreeMap<Integer, Column> columns;
    TreeMap<Integer, Row> rows;
    private String styleID;
    private double rowheight;
    private double columnwidth;
    private int expandedcolumncount;
    private int expandedrowcount;

    /**
     * Constructs a new SSTable.
     * 
     * @see nl.fountain.xelem.excel.Worksheet#getTable()
     */
    public SSTable() {
        rows = new TreeMap<Integer, Row>();
        columns = new TreeMap<Integer, Column>();
    }
    
    public void setStyleID(String id) {
        styleID = id;
    }

    public String getStyleID() {
        return styleID;
    }
    
    public void setDefaultRowHeight(double points) {
        rowheight = points;
    }
    
    public void setDefaultColumnWidth(double points) {
        columnwidth = points;
    }
    
    public Column addColumn() {
        return addColumnAt(maxColumnIndex() + 1, new SSColumn());
    }
    
    public Column addColumnAt(int index) {
        return addColumnAt(index, new SSColumn());
    }
    
    public Column addColumn(Column column) {
        return addColumnAt(maxColumnIndex() + 1, column);
    }
    
    public Column addColumnAt(int index, Column column) {
        if (index < Worksheet.firstColumn || index > Worksheet.lastColumn) {
            throw new IndexOutOfBoundsException("columnIndex = " + index);
        }
        columns.put(new Integer(index), column);
        return column;
    }
    
    public Column removeColumnAt(int columnIndex) {
        return (Column) columns.remove(new Integer(columnIndex));
    }
    
    public Column getColumnAt(int columnIndex) {
        Column column = columns.get(new Integer(columnIndex));
        if (column == null) {
            column = addColumnAt(columnIndex);
        }
        return column;
    }
    
    public boolean hasColumnAt(int index) {
        return columns.get(new Integer(index)) != null;
    }
    
    public Collection<Column> getColumns() {
        return columns.values();
    }

    public Row addRow() {
        return addRowAt(maxRowIndex() + 1, new SSRow());
    }

    public Row addRowAt(int index) {
        return addRowAt(index, new SSRow());
    }

    public Row addRow(Row row) {
        return addRowAt(maxRowIndex() + 1, row);
    }
    
    public Row addRowAt(int index, Row row) {
        if (index < Worksheet.firstRow || index > Worksheet.lastRow) {
            throw new IndexOutOfBoundsException("rowIndex = " + index);
        }
        rows.put(new Integer(index), row);
        return row;
    }

    public Row removeRowAt(int rowIndex) {
        Row row = (Row) rows.remove(new Integer(rowIndex));
        return row;
    }

    public Collection<Row> getRows() {
        return rows.values();
    }
    
    public TreeMap<Integer, Row> getRowMap() {
        return rows;
    }

    public Row getRowAt(int rowIndex) {
        Row row = (Row) rows.get(new Integer(rowIndex));
        if (row == null) {
            row = addRowAt(rowIndex);
        }
        return row;
    }
    
    public boolean hasRowAt(int index) {
        return rows.get(new Integer(index)) != null;
    }

    public int rowCount() {
        return rows.size();
    }
    
    public int columnCount() {
        return columns.size();
    }
    
    public boolean hasChildren() {
        return (columns.size() + rows.size()) > 0;
    }
    
    public int maxCellIndex() {
        int max = 0;
        for (Row r : rows.values()) {
            if (r.maxCellIndex() > max) max = r.maxCellIndex();
        }
        return max;
    }
    
    public int maxRowIndex() {
        int lastKey;
        if (rows.size() == 0) {
            lastKey = 0;
        } else {
            lastKey = rows.lastKey().intValue();
        }
        return lastKey;
    }
    
    public int maxColumnIndex() {
        int lastKey;
        if (columns.size() == 0) {
            lastKey = 0;
        } else {
            lastKey = columns.lastKey().intValue();
        }
        return lastKey;
    }
    
    public Iterator<Row> rowIterator() {
        return new RowIterator();
    }
    
    public Iterator<Column> columnIterator() {
        return new ColumnIterator();
    }
    
    public String getTagName() {
        return "Table";
    }
    
    public String getNameSpace() {
        return XMLNS_SS;
    }
    
    public String getPrefix() {
        return PREFIX_SS;
    } 
    
    public Element assemble(Element parent, GIO gio) {
        Document doc = parent.getOwnerDocument();
        Element te = assemble(doc, gio);
        
        if (getStyleID() != null) {
            te.setAttributeNodeNS(createAttributeNS(doc, "StyleID", getStyleID()));
            gio.addStyleID(getStyleID());
        }
        if (rowheight > 0.0) te.setAttributeNodeNS(
                createAttributeNS(doc, "DefaultRowHeight", "" + rowheight));
        if (columnwidth > 0.0) te.setAttributeNodeNS(
                createAttributeNS(doc, "DefaultColumnWidth", "" + columnwidth));  
        
        parent.appendChild(te);
        
        Iterator<Column> iterC = columnIterator();
        while (iterC.hasNext()) {
            iterC.next().assemble(te, gio);
        }
        
        Iterator<Row> iterR = rowIterator();
        while (iterR.hasNext()) {
            iterR.next().assemble(te, gio);
        }
        return te;
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
	
    public int getExpandedColumnCount() {
        return expandedcolumncount;
    }
    
    // method called by ExcelReader
    private void setExpandedColumnCount(String s) {
        expandedcolumncount = Integer.parseInt(s);
    }
    
    public int getExpandedRowCount() {
        return expandedrowcount;
    }
    
    // method called by ExcelReader
    private void setExpandedRowCount(String s) {
        expandedrowcount = Integer.parseInt(s);
    }
    
    /////////////////////////////////////////////
    private class RowIterator implements Iterator<Row> {
        
        private Iterator<Integer> rit;
        private Integer current;
        private int prevIndex;
        
        protected RowIterator() {
            rit = rows.keySet().iterator();
        }

        public void remove() {
            throw new UnsupportedOperationException();
        }

        public boolean hasNext() {
            return rit.hasNext();
        }

        public Row next() {
            current = rit.next();  
            int curIndex = current.intValue();
            SSRow r = (SSRow) rows.get(current);
            if (prevIndex + 1 != curIndex) {
                r.setIndex(curIndex);
            } else {
                r.setIndex(0);
            }
            prevIndex = curIndex;
            return r;
        }
        
    }
    
    private class ColumnIterator implements Iterator<Column> {
        
        private Iterator<Integer> cit;
        private Integer current;
        private int prevIndex;
        private int maxSpan;
        
        protected ColumnIterator() {
            cit = columns.keySet().iterator();
        }
        
        // @see java.util.Iterator#remove()
        public void remove() {
            throw new UnsupportedOperationException();
        }

        // @see java.util.Iterator#hasNext()
        public boolean hasNext() {
            return cit.hasNext();
        }

        // @see java.util.Iterator#next()
        public Column next() {
            current = cit.next();  
            int curIndex = current.intValue();
            SSColumn c = (SSColumn) columns.get(current);
            if (prevIndex + 1 != curIndex) {
                c.setIndex(curIndex);
            } else {
                c.setIndex(0);
            }
            prevIndex = curIndex;
            
            return c;
        }
        
    }

}
