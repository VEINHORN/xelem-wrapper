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

import java.util.Collection;
import java.util.Iterator;
import java.util.TreeMap;



/**
 * Represents the Table element.
 */
public interface Table extends XLElement {
    
	/**
	 * Sets the ss:StyleID on this table. If no styleID is set on a table,
	 * the ss:StyleID-attribute is not deployed in the resulting xml and
	 * Excel employes the Default-style on the table.
	 * 
	 * @param 	id	the id of the style to employ on this table.
	 */
    void setStyleID(String id);
    
    /**
     * Gets the ss:StyleID which was set on this table.
     * 
     * @return 	The id of the style to employ on this table or
     * 			<code>null</code> if no styleID was previously set.
     */
    String getStyleID();
    
    /**
     * Sets the default row height.
     * 
     * @param points	The default row height.
     */
    void setDefaultRowHeight(double points);
    
    /**
     * Sets the default column width.
     * 
     * @param points	The default column width.
     */
    void setDefaultColumnWidth(double points);
    
    /**
     * Adds a new Column to this table. If no columns were previously added
     * the column will be added at index 1. Otherwise the column will be added
     * at {@link #maxColumnIndex()} + 1.
     * 
     * @return A new column.
     * @throws IndexOutOfBoundsException If the calculated index is less then
     * 			{@link nl.fountain.xelem.excel.Worksheet#firstColumn} or greater
     * 			then {@link nl.fountain.xelem.excel.Worksheet#lastColumn}.
     */
    Column addColumn();
    
    /**
     * Adds a new Column at the given index to this table. If the index 
     * was allready occupied by another column, replaces this column.
     * 
     * @param index The index (column number) of the column.
     * @return A new column.
     * @throws IndexOutOfBoundsException If the calculated index is less then
     * 			{@link nl.fountain.xelem.excel.Worksheet#firstColumn} or greater
     * 			then {@link nl.fountain.xelem.excel.Worksheet#lastColumn}.
     */
    Column addColumnAt(int index);
    
    /**
     * Adds the given column to this table. If no columns were previously added
     * the column will be added at index 1. Otherwise the ccolumn will be added
     * at {@link #maxColumnIndex()} + 1.
     * 
     * @param column	The column to be added.
     * @return The passed column.
     * @throws IndexOutOfBoundsException If the calculated index is less then
     * 			{@link nl.fountain.xelem.excel.Worksheet#firstColumn} or greater
     * 			then {@link nl.fountain.xelem.excel.Worksheet#lastColumn}.
     */
    Column addColumn(Column column);
    
    /**
     * Adds the given Column at the given index to this table. If the index 
     * was allready occupied by another column, replaces this column.
     * 
     * @param index The index (column number) of the column.
     * @param column	The column to be added.
     * @return The passed column.
     * @throws IndexOutOfBoundsException If the calculated index is less then
     * 			{@link nl.fountain.xelem.excel.Worksheet#firstColumn} or greater
     * 			then {@link nl.fountain.xelem.excel.Worksheet#lastColumn}.
     */
    Column addColumnAt(int index, Column column);
    
    /**
     * Removes the column at the given index.
     * 
     * @param index The index (column number) of the column.
     * 
     * @return The removed column or <code>null</code> if the index was not occupied
     * 			by a column.
     */
    Column removeColumnAt(int index);
    
    /**
     * Gets the column at the given index. If no column was at the given index,
     * returns a new column.
     * 
     * @param index The index (column number) of the column.
     * 
     * @return The column at the given index. Never <code>null</code>. 
     * @throws IndexOutOfBoundsException If the calculated index is less then
     * 			{@link nl.fountain.xelem.excel.Worksheet#firstColumn} or greater
     * 			then {@link nl.fountain.xelem.excel.Worksheet#lastColumn}.
     */
    Column getColumnAt(int index);
    
    /**
     * Specifies whether this table has a column at the given index.
     */
    boolean hasColumnAt(int index);
    
    /**
     * Gets all the columns of this table in the order of their index.
     * 
     * @return A collection of columns.
     */
    Collection<Column> getColumns();
    
    /**
     * Returns an iterator for the columns in this table. Columns are aware of their
     * index number when passed by this iterator.
     * 
     * @return An iterator for the columns in this table.
     */
    Iterator<Column> columnIterator();
    
    /**
     * Adds a new Row to this table. If no rows were previously added
     * the row will be added at index 1. Otherwise the row will be added
     * at {@link #maxRowIndex()} + 1.
     * 
     * @return A new Row.
     * @throws IndexOutOfBoundsException If the calculated index is less then
     * 			{@link nl.fountain.xelem.excel.Worksheet#firstRow} or greater
     * 			then {@link nl.fountain.xelem.excel.Worksheet#lastRow}.
     */
    Row addRow();
    
    /**
     * Adds a new Row at the given index to this table. If the index 
     * was allready occupied by another row, replaces this row.
     * 
     * @param index The index (row number) of the row.
     * @return A new Row.
     * @throws IndexOutOfBoundsException If the calculated index is less then
     * 			{@link nl.fountain.xelem.excel.Worksheet#firstRow} or greater
     * 			then {@link nl.fountain.xelem.excel.Worksheet#lastRow}.
     */
    Row addRowAt(int index);
    
    /**
     * Adds the given row to this table. If no rows were previously added
     * the row will be added at index 1. Otherwise the row will be added
     * at {@link #maxRowIndex()} + 1.
     * 
     * @param row	The row to be added.
     * @return The passed row.
     * @throws IndexOutOfBoundsException If the calculated index is less then
     * 			{@link nl.fountain.xelem.excel.Worksheet#firstRow} or greater
     * 			then {@link nl.fountain.xelem.excel.Worksheet#lastRow}.
     */
    Row addRow(Row row);
    
    /**
     * Adds the given Row at the given index to this table. If the index 
     * was allready occupied by another row, replaces this row.
     * 
     * @param index The index (row number) of the row.
     * @param row	The row to be added.
     * @return The passed row.
     * @throws IndexOutOfBoundsException If the calculated index is less then
     * 			{@link nl.fountain.xelem.excel.Worksheet#firstRow} or greater
     * 			then {@link nl.fountain.xelem.excel.Worksheet#lastRow}.
     */
    Row addRowAt(int index, Row row); 
    
    /**
     * Removes the row at the given index.
     * 
     * @param index The index (row number) of the row.
     * 
     * @return The removed row or <code>null</code> if the index was not occupied
     * 			by a row.
     */
    Row removeRowAt(int index); 
    
    /**
     * Gets the row at the given index. If no row was at the given index,
     * returns a new row at that index.
     * 
     * @param index The index (row number) of the row.
     * @return The row at the given index. Never <code>null</code>
     * @throws IndexOutOfBoundsException If the calculated index is less then
     * 			{@link nl.fountain.xelem.excel.Worksheet#firstRow} or greater
     * 			then {@link nl.fountain.xelem.excel.Worksheet#lastRow}.
     */
    Row getRowAt(int index); 
    
    /**
     * Specifies whether this table has a row at the given index.
     */
    boolean hasRowAt(int index);
    
    
    /**
     * Gets all the rows of this table in the order of their index.
     * 
     * @return A collection of rows.
     */
    Collection<Row> getRows(); 
    
    /**
     * Gets all the rows of this table. The key of the Map.Entrys is of type
     * Integer and stands for the index number of the row.
     * 
     * @return A TreeMap of rows.
     */
    TreeMap<Integer, Row> getRowMap();
    
    /**
     * Returns an iterator for the rows in this table. Rows are aware of their
     * index number when passed by this iterator.
     * 
     * @return An iterator for the rows in this table.
     */
    Iterator<Row> rowIterator();
    
    /**
     * Returns the number of rows in this table.
     * @return The number of rows in this table.
     */
    int rowCount();
    
    /**
     * Returns the number of columns in this table.
     * @return The number of columns in this table.
     */
    int columnCount();
    
    /**
     * Indicates whether this table has any rows or columns.
     */
    boolean hasChildren();
    
    /**
     * Gets the highest index number of all the cells of all the rows in this table:
     * the right-most cell.
     * 
     * @return The highest index number of cells in this table.
     */
    int maxCellIndex();
    
    /**
     * Gets the highest index number of the rows in this table.
     * 
     * @return The highest index number of the rows in this table.
     */
    int maxRowIndex();
    
    /**
     * Gets the highest index number of the columns in this table.
     * 
     * @return The highest index number of the columns in this table.
     */
    int maxColumnIndex();  
    
    /**
     * An indicator of the index of the right-most column on this table.
     * May have been set when reading workbooks.
     * @return an indicator of the index of the right-most column
     */
    int getExpandedColumnCount();
    
    /**
     * An indicator of the index of the bottom row on this table.
     * May have been set when reading workbooks.
     * @return an indicator of the index of the bottom row
     */
    int getExpandedRowCount();
    
}
