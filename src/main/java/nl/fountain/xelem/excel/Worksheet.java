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
import java.util.Map;

import nl.fountain.xelem.Address;
import nl.fountain.xelem.CellPointer;

/**
 * Represents the Worksheet element.
 * 
 */
public interface Worksheet extends XLElement {
    
    /**
     * The left-most limit of the worksheets columns.
     */
    public static int firstColumn = 1;
    /**
     * The top row of the worksheet.
     */
    public static int firstRow = 1;
    /**
     * The right-most limit of the worksheets columns.
     */
    public static int lastColumn = 256;
    /**
     * The bottom row of the worksheet.
     */
    public static int lastRow = 65536;
    
    /**
     * Returns the name of this worksheet.
     * @return the name of this worksheet
     */
    String getName();
    
    /**
     * Returns the name of this worksheet for purposes of reference.
     * @return the name of this worksheet or, if there are spaces in the name,
     * 		the name of this worksheet between single quoutes ('...')
     */
    String getReferenceName();
    
    /**
     * Sets whether protection will be applied to this worksheet.
     */
    void setProtected(boolean p);
    
    /**
     * Specifies whether protection is applied to this worksheet.
     * @return <code>true</code> if this worksheet is protected
     */
    boolean isProtected();
    
    /**
     * Sets whether this worksheet will be displayed from right to left.
     */
    void setRightToLeft(boolean r);
    
    /**
     * Specifies whether this worksheet is displayed from right to left.
     * @return <code>true</code> if this worksheet is displayed from right to left
     */
    boolean isRightToLeft();
    
    /**
     * Adds a new NamedRange to this worksheet. Named ranges are usefull when
     * working with formulas. If this workbook allready contained a NamedRange
     * with the same name, replaces that NamedRange.
     * <P>
     * If the workbook has a named range with the same name, the named range
     * on the worksheet level has precedence over the one defined on the
     * workbook level. 
     * <P>
     * The string <code>refersTo</code> should be in R1C1-reference style i.e.
     * in the format <code>[worksheet name]!R1C1:R1C1</code>. You may skip
     * the worksheet indicator: <code>R1C1:R1C1</code> will do as well.
     * 
     * @param name		The name to apply to the range.
     * @param refersTo	A String of R1C1-reference style.
     * 
     * @return New NamedRange.
     */
    NamedRange addNamedRange(String name, String refersTo);
    
    /**
     * Adds a new NamedRange to this worksheet.
     * If this workbook allready contained a NamedRange
     * with the same name, replaces that NamedRange.
     * <P>
     * If the workbook has a named range with the same name, the named range
     * on the worksheet level has precedence over the one defined on the
     * workbook level. 
     * 
     * @param namedRange	the NamedRange to be added to this worksheet
     * @return the given NamedRange
     */
    NamedRange addNamedRange(NamedRange namedRange);
    
    /**
     * Gets all the NamedRanges that were added to this worksheet. The map-keys
     * are equal to their names.
     * 
     * @return a map with NamedRanges
     */
    Map<String, NamedRange> getNamedRanges();
    
    /**
     * Sets the given worksheetOptions as the worksheetOptions of this worksheet.
     * @param wso the worksheetOptions to be set on this worksheet
     */
    void setWorksheetOptions(WorksheetOptions wso);
    
    /**
     * Indicates whether WorksheetOptions was added to this worksheet.
     * 
     * @return <code>true</code> if this worksheet has WorksheetOptions.
     */
    boolean hasWorksheetOptions();
    
    /**
     * Gets the WorksheetOptions of this worksheet.
     * 
     * @return The WorksheetOptions of this worksheet. Never <code>null</code>.
     */
    WorksheetOptions getWorksheetOptions();
    
    /**
     * Sets the given table as the table of this worksheet.
     * @param table the table to be set on this worksheet
     */
    void setTable(Table table);
    
    /**
     * Gets the table of this worksheet.
     * 
     * @return The table of this worksheet. Never <code>null</code>.
     */
    Table getTable();
    
    /**
     * Indicates whether this worksheet has a table.
     */
    boolean hasTable();
    
    /**
     * Gets the cellpointer of this worksheet.
     * 
     * @return This worksheets cellpointer.
     */
    CellPointer getCellPointer();
       
    /**
     * Adds a new cell at the current position of the worksheets
     * {@link nl.fountain.xelem.CellPointer} and moves the cellpointer.
     * <br>
     * If this worksheet did not have a table, 
     * a new table will be added.
     * 
     * @return A new cell with an initial datatype of "String" and an
     * 			empty ("") value.
     * 
     * @throws IndexOutOfBoundsException if the position of the cellpointer
     * 			is not within the limits of the worksheet.
     * 
     * @see nl.fountain.xelem.CellPointer#move()
     */
    Cell addCell();
    
    /**
     * Adds the given cell at the current position of the worksheets
     * {@link nl.fountain.xelem.CellPointer} and moves the cellpointer.
     * <br>
     * If this worksheet did not have a table, 
     * a new table will be added.
     * 
     * @return The passed cell.
     * 
     * @throws IndexOutOfBoundsException if the position of the cellpointer
     * 			is not within the limits of the worksheet.
     * 
     * @see nl.fountain.xelem.CellPointer#move()
     */
    Cell addCell(Cell cell);
    
    /**
     * Adds a new cell at the current position of the worksheets
     * {@link nl.fountain.xelem.CellPointer} and moves the cellpointer.
     *  <br>
     * If this worksheet did not have a table, 
     * a new table will be added.
     * <P>
     * Datatype and value of the new cell depend on the passed object.
     * 
     * @param data	The data to be displayed in this cell.
     * 
     * @return A new cell.
     * 
     * @throws IndexOutOfBoundsException if the position of the cellpointer
     * 			is not within the limits of the worksheet.
     * 
     * @see nl.fountain.xelem.excel.Cell#setData(Object)
     * @see nl.fountain.xelem.CellPointer#move()
     */
    Cell addCell(Object data);
    
    /**
     * Adds a new cell at the current position of the worksheets
     * {@link nl.fountain.xelem.CellPointer} and moves the cellpointer.
     *  <br>
     * If this worksheet did not have a table, 
     * a new table will be added.
     * <P>
     * Datatype and value of the new cell depend on the passed object.
     * The new cell will have it's styleID set to the given id.
     * 
     * @param data		The data to be displayed in this cell.
     * @param styleID	The id of the style to employ on this cell.
     * 
     * @return A new cell.
     * 
     * @throws IndexOutOfBoundsException if the position of the cellpointer
     * 			is not within the limits of the worksheet.
     * 
     * @see nl.fountain.xelem.excel.Cell#setData(Object)
     * @see nl.fountain.xelem.excel.Cell#setStyleID(String)
     * @see nl.fountain.xelem.CellPointer#move()
     */
    Cell addCell(Object data, String styleID);
    
    /**
     * Adds a new cell at the current position of the worksheets
     * {@link nl.fountain.xelem.CellPointer} and moves the cellpointer.
     *  <br>
     * If this worksheet did not have a table, 
     * a new table will be added.
     * 
     * @param data		The data to be displayed in this cell.
     * 
     * @return A new cell with a datatype "Number" and the given double as value. 
     * 
     * @throws IndexOutOfBoundsException if the position of the cellpointer
     * 			is not within the limits of the worksheet.
     * 
     * @see nl.fountain.xelem.excel.Cell#setData(double)
     * @see nl.fountain.xelem.CellPointer#move()
     */
    Cell addCell(double data);
    
    /**
     * Adds a new cell at the current position of the worksheets
     * {@link nl.fountain.xelem.CellPointer} and moves the cellpointer.
     *  <br>
     * If this worksheet did not have a table, 
     * a new table will be added.
     * The new cell will have it's styleID set to the given id.
     * 
     * @param data		The data to be displayed in this cell.
     * @param styleID	The id of the style to employ on this cell.
     * 
     * @return A new cell with a datatype "Number" and the given double as value. 
     * 
     * @throws IndexOutOfBoundsException if the position of the cellpointer
     * 			is not within the limits of the worksheet.
     * 
     * @see nl.fountain.xelem.excel.Cell#setData(double)
     * @see nl.fountain.xelem.excel.Cell#setStyleID(String)
     * @see nl.fountain.xelem.CellPointer#move()
     */
    Cell addCell(double data, String styleID);
    
    /**
     * Adds a new cell at the current position of the worksheets
     * {@link nl.fountain.xelem.CellPointer} and moves the cellpointer.
     *  <br>
     * If this worksheet did not have a table, 
     * a new table will be added.
     * 
     * @param data		The data to be displayed in this cell.
     * 
     * @return A new cell with a datatype "Number" and the given int as value.
     * 
     * @throws IndexOutOfBoundsException if the position of the cellpointer
     * 			is not within the limits of the worksheet.
     * 
     * @see nl.fountain.xelem.CellPointer#move()
     */
    Cell addCell(int data);
    
    /**
     * Adds a new cell at the current position of the worksheets
     * {@link nl.fountain.xelem.CellPointer} and moves the cellpointer.
     *  <br>
     * If this worksheet did not have a table, 
     * a new table will be added.
     * 
     * @param data		The data to be displayed in this cell.
     * @param styleID	The id of the style to employ on this cell.
     * 
     * @return A new cell with a datatype "Number" and the given int as value. 
     * 
     * @throws IndexOutOfBoundsException if the position of the cellpointer
     * 			is not within the limits of the worksheet.
     * 
     * @see nl.fountain.xelem.excel.Cell#setStyleID(String)
     * @see nl.fountain.xelem.CellPointer#move
     */
    Cell addCell(int data, String styleID);
    
    /**
     * Adds a new cell at the given address. If this worksheet
     * didn't have a row at the row index of the given address, 
     * a new row will be added. If there was a cell at the given address,
     * replaces this cell. If this worksheet did not have a table, 
     * a new table will be added.
     * <br>
     * Moves the cellpointer to a new position relative to
     * <code>address</code>.
     * 
     * @param address The address where the new cell should be added.
     * @return A new cell with an initial datatype of "String" and an
     * 			empty ("") value.
     * 
     * @throws IndexOutOfBoundsException if the address
     * 			is not within the limits of the worksheet.
     * 
     * @see nl.fountain.xelem.CellPointer#move()
     */
    Cell addCellAt(Address address);
    
    /**
     * Adds a new cell at the given row and column index. If this worksheet
     * didn't have a row at the given index, a new row will be added. If the
     * place at the given coordinates was allredy occupied by another cell,
     * replaces this cell. If this worksheet did not have a table, 
     * a new table will be added.
     * <br>
     * Moves the cellpointer to a new position relative to
     * <code>rowIndex</code> and <code>columnIndex</code>.
     * 
     * @param rowIndex		The row index (row number).
     * @param columnIndex	The column index (column number).
     * 
     * @return A new cell with an initial datatype of "String" and an
     * 			empty ("") value.
     * 
     * @throws IndexOutOfBoundsException if the position
     * 			is not within the limits of the worksheet.
     * 
     * @see nl.fountain.xelem.CellPointer#move()
     */
    Cell addCellAt(int rowIndex, int columnIndex);
    
    /**
     * Adds a new cell at the address specified by the A1-reference string. 
     * If this worksheet
     * didn't have a row at the row index specified, a new row will be added. If the
     * place at the given coordinates was allredy occupied by another cell,
     * replaces this cell. If this worksheet did not have a table, 
     * a new table will be added.
     * <br>
     * Moves the cellpointer to a new position relative to
     * the address specified by the A1-reference string.
     * 
     * @param a1_ref	a string in A1-reference style
     * 
     * @return A new cell with an initial datatype of "String" and an
     * 			empty ("") value.
     * 
     * @throws IndexOutOfBoundsException if the position
     * 			is not within the limits of the worksheet.
     * 
     * @see nl.fountain.xelem.CellPointer#move()
     */
    Cell addCellAt(String a1_ref);
    
    /**
     * Adds the given cell at the given address. If this worksheet
     * didn't have a row at the row index of the given address, 
     * a new row will be added. If there was a cell at the given address,
     * replaces this cell. If this worksheet did not have a table, 
     * a new table will be added.
     * <br>
     * Moves the cellpointer to a new position relative to
     * <code>address</code>.
     * 
     * @param address The address where the cell should be added.
     * @param cell			The cell to be added.
     * @return The passed cell.
     * 
     * @throws IndexOutOfBoundsException if the address
     * 			is not within the limits of the worksheet.
     * 
     * @see nl.fountain.xelem.CellPointer#move()
     */
    Cell addCellAt(Address address, Cell cell);
    
    /**
     * Adds the given cell at the given row and column index. If this worksheet
     * didn't have a row at the given index, a new row will be added. If the
     * place at the given coordinates was allredy occupied by another cell,
     * replaces this cell. If this worksheet did not have a table, 
     * a new table will be added. 
     * <br>
     * Moves the cellpointer to a new position relative to
     * <code>rowIndex</code> and <code>columnIndex</code>.
     * 
     * @param rowIndex		The row index (row number).
     * @param columnIndex	The column index (column number).
     * @param cell			The cell to be added.
     * 
     * @return The passed cell.
     * 
     * @throws IndexOutOfBoundsException if the position
     * 			is not within the limits of the worksheet.
     * 
     * @see nl.fountain.xelem.CellPointer#move()
     */
    Cell addCellAt(int rowIndex, int columnIndex, Cell cell);
    
    /**
     * Adds the given cell at the address specified by the A1-reference string.
     * If this worksheet
     * didn't have a row at the index specified, a new row will be added. If the
     * place at the given coordinates was allredy occupied by another cell,
     * replaces this cell. If this worksheet did not have a table, 
     * a new table will be added. 
     * <br>
     * Moves the cellpointer to a new position relative to
     * the address specified by the A1-reference string.
     * 
     * @param a1_ref		a string of A1-reference style
     * @param cell			the cell to be added.
     * 
     * @return The passed cell.
     * 
     * @throws IndexOutOfBoundsException if the position
     * 			is not within the limits of the worksheet.
     * 
     * @see nl.fountain.xelem.CellPointer#move()
     */
    Cell addCellAt(String a1_ref, Cell cell);
    
    /**
     * Removes the cell at the given address.
     * 
     * @param address	The position of the cell to remove.
     * @return The removed cell or <code>null</code>.
     */
    Cell removeCellAt(Address address);
    
    /**
     * Removes the cell at the given coordinates.
     * 
     * @param rowIndex		The row index (row number).
     * @param columnIndex	The column index (column number).
     * 
     * @return The removed cell or <code>null</code>.
     */
    Cell removeCellAt(int rowIndex, int columnIndex);
    
    /**
     * Removes the cell at the address specified by the A1-reference string.
     * 
     * @param a1_ref	a string of A1-reference style	
     * 
     * @return The removed cell or <code>null</code>.
     */
    Cell removeCellAt(String a1_ref);
    
    /**
     * Gets the cell at the given address.
     * 
     * @param address		The position of the cell.
     * @return The cell at the given address. Never <code>null</code> 
     */
    Cell getCellAt(Address address);
    
    /**
     * Gets the cell at the given coordinates.
     * 
     * @param rowIndex		The row index (row number).
     * @param columnIndex	The column index (column number).
     * 
     * @return The cell at the given coordinates. Never <code>null</code> 
     */
    Cell getCellAt(int rowIndex, int columnIndex);
    
    /**
     * Gets the cell at address specified by the given string in A1-reference style.
     * 
     * @param a1_ref	a string of A1-reference style
     * 
     * @return The cell at the given coordinates. Never <code>null</code>
     */
    Cell getCellAt(String a1_ref);
    
    /**
     * Specifies whether there is a cell at the given address.
     */
    boolean hasCellAt(Address address);
    
    /**
     * Specifies whether there is a cell at the intersection of the given
     * row and column index.
     */
    boolean hasCellAt(int rowIndex, int columnIndex);
    
    /**
     * Specifies whether there is a cell at the address specified by the 
     * given string in A1-reference style.
     */
    boolean hasCellAt(String a1_ref);
    
    /**
     * Adds a new Row to this worksheet. If no rows were previously added
     * the row will be added at index 1. Otherwise the row will be added at 
     * {@link nl.fountain.xelem.excel.Table#maxRowIndex() maxRowIndex} 
     * of the underlying Table + 1.
     * 
     * @return A new Row.
     * @throws IndexOutOfBoundsException If the calculated index is less then
     * 			{@link #firstRow} or greater
     * 			then {@link #lastRow}.
     */
    Row addRow();
    
    /**
     * Adds a new Row at the given index to this worksheet. If the index 
     * was allready occupied by another row, replaces this row.
     * 
     * @param index The index (row number) of the row.
     * @return A new Row.
     *  @throws IndexOutOfBoundsException If the given index is less then
     * 			{@link #firstRow} or greater
     * 			then {@link #lastRow}.
     */
    Row addRowAt(int index);
    
    /**
     * Adds the given row to this worksheet. If no rows were previously added
     * the row will be added at index 1. Otherwise the row will be added at
     * {@link nl.fountain.xelem.excel.Table#maxRowIndex() maxRowIndex} 
     * of the underlying Table + 1.
     * 
     * @param row	The row to be added.
     * @return The passed row.
     *  @throws IndexOutOfBoundsException If the calculated index is less then
     * 			{@link #firstRow} or greater
     * 			then {@link #lastRow}.
     */
    Row addRow(Row row);
    
    /**
     * Adds the given Row at the given index to this worksheet. If the index 
     * was allready occupied by another row, replaces this row.
     * 
     * @param index The index (row number) of the row.
     * @param row	The row to be added.
     * @return The passed row.
     *  @throws IndexOutOfBoundsException If the given index is less then
     * 			{@link #firstRow} or greater
     * 			then {@link #lastRow}.
     */
    Row addRowAt(int index, Row row);
    
    /**
     * Removes the row at the given index.
     * 
     * @param rowIndex The index (row number) of the row.
     * 
     * @return The removed row or <code>null</code> if the index was not occupied
     * 			by a row.
     */
    Row removeRowAt(int rowIndex);
    
    /**
     * Gets all the rows of this worksheet in the order of their index.
     * 
     * @return A collection of rows.
     */
    Collection<Row> getRows();
    
    /**
     * Gets the row at the given index.
     * 
     * @param rowIndex The index (row number) of the row.
     * 
     * @return The row at the given index, never <code>null</code>.
     * @throws IndexOutOfBoundsException If the calculated index is less then
     * 			{@link #firstRow} or greater
     * 			then {@link #lastRow}.
     */
    Row getRowAt(int rowIndex);
    
    /**
     * Specifies whether there is a row at the given row index.
     */
    boolean hasRowAt(int rowIndex);
    
    /**
     * Adds a new Column to this worksheet. If no columns were previously added
     * the column will be added at index 1. Otherwise the column will be added at 
     * {@link nl.fountain.xelem.excel.Table#maxColumnIndex() maxColumnIndex} 
     * of the underlying Table + 1.
     * 
     * @return A new column.
     * @throws IndexOutOfBoundsException If the calculated index is less then
     * 			{@link #firstColumn} or greater
     * 			then {@link #lastColumn}.
     */
    Column addColumn();
    
    /**
     * Adds a new Column at the given index to this worksheet. If the index 
     * was allready occupied by another column, replaces this column.
     * 
     * @param index The index (column number) of the column.
     * @return A new column.
     * @throws IndexOutOfBoundsException If the given index is less then
     * 			{@link #firstColumn} or greater
     * 			then {@link #lastColumn}.
     */
    Column addColumnAt(int index);
    
    /**
     * Adds a new Column at the index specified by the given label.
     * If the index 
     * was allready occupied by another column, replaces this column.
     * 
     * @param label a string ("A" through "IV" inclusive) 
     * 				corresponding to the column index
     * @return A new column.
     * @throws IndexOutOfBoundsException If the given index is less then
     * 			{@link #firstColumn} or greater
     * 			then {@link #lastColumn}.
     */
    Column addColumnAt(String label);
    
    /**
     * Adds the given column to this worksheet. If no columns were previously added
     * the column will be added at index 1. Otherwise the ccolumn will be added at 
     * {@link nl.fountain.xelem.excel.Table#maxColumnIndex() maxColumnIndex} 
     * of the underlying Table + 1.
     * 
     * @param column	The column to be added.
     * @return The passed column.
     * @throws IndexOutOfBoundsException If the calculated index is less then
     * 			{@link #firstColumn} or greater
     * 			then {@link #lastColumn}.
     */
    Column addColumn(Column column);
    
    /**
     * Adds the given Column at the given index to this worksheet. If the index 
     * was allready occupied by another column, replaces this column.
     * 
     * @param index The index (column number) of the column.
     * @param column	The column to be added.
     * @return The passed column.
     * @throws IndexOutOfBoundsException If the given index is less then
     * 			{@link #firstColumn} or greater
     * 			then {@link #lastColumn}.
     */
    Column addColumnAt(int index, Column column);
    
    /**
     * Adds the given Column at the given index specified by the given label.
     * If the index 
     * was allready occupied by another column, replaces this column.
     * 
     * @param label 	a string ("A" through "IV" inclusive) 
     * 					corresponding to the column index
     * @param column	The column to be added.
     * @return The passed column.
     * @throws IndexOutOfBoundsException If the given index is less then
     * 			{@link #firstColumn} or greater
     * 			then {@link #lastColumn}.
     */
    Column addColumnAt(String label, Column column);
    
    /**
     * Removes the column at the given index.
     * 
     * @param columnIndex The index (column number) of the column.
     * 
     * @return The removed column or <code>null</code> if the index was not occupied
     * 			by a column.
     */
    Column removeColumnAt(int columnIndex);
    
    /**
     * Removes the column at the index specified by the given label.
     * 
     * @param label 	a string ("A" through "IV" inclusive) 
     * 					corresponding to the column index
     * 
     * @return The removed column or <code>null</code> if the index was not occupied
     * 			by a column.
     */
    Column removeColumnAt(String label);
    
    /**
     * Gets all the columns of this worksheet in the order of their index.
     * 
     * @return A collection of columns.
     */
    Collection<Column> getColumns();
    
    /**
     * Gets the column at the given index. If no column was at the given index,
     * returns a new column.
     * 
     * @param columnIndex The index (column number) of the column.
     * 
     * @return The column at the given index. Never <code>null</code>.
     * @throws IndexOutOfBoundsException If the given index is less then
     * 			{@link #firstColumn} or greater
     * 			then {@link #lastColumn}.
     */
    Column getColumnAt(int columnIndex);
    
    /**
     * Gets the column at the index specified by the given label. 
     * If no column was at the index,
     * returns a new column.
     * 
     * @param label 	a string ("A" through "IV" inclusive) 
     * 					corresponding to the column index
     * 
     * @return The column at the given index. Never <code>null</code>.
     * @throws IndexOutOfBoundsException If the given index is less then
     * 			{@link #firstColumn} or greater
     * 			then {@link #lastColumn}.
     */
    Column getColumnAt(String label);
    
    /**
     * Specifies whether there is a column at the given column index.
     */
    boolean hasColumnAt(int columnIndex);
    
    /**
     * Specifies whether there is a column at the index specified by the given label.
     * 
     * @param label 	a string ("A" through "IV" inclusive) 
     * 					corresponding to the column index
     */
    boolean hasColumnAt(String label);
    
    /**
     * Sets the AutoFilter-option on the specified range.
     * <P>
     * The String <code>rcString</code> should include all of the
     * columns that are to be equiped with a drop-down button at their column-
     * headings, and should either include just the row of the column-headings
     * or the entire range on which the AutoFilter must be set.
     * <P>
     * Here are some examples:
     * <br>
     * Given a table of 3 columns and 11 rows. The column-headings are at row 10.
     * The table-data expands to row 21.
     * <UL>
     * <LI><code><b>R10C1:R10C3</b></code> - the AutoFilter includes the entire range
     * 		(R10C1:R21C3) and will dynamically expand if more rows of data are
     * 		added below.
     * <LI><code><b>R10C1:R21C3</b></code> - the AutoFilter includes the entire range
     * 		(R10C1:R21C3) and will dynamically expand if more rows of data are
     * 		added below.
     * <LI><code><b>R10C1:R10C2</b></code> - the AutoFilter includes the entire range
     * 		(R10C1:R21C3) and will dynamically expand if more rows of data are
     * 		added below, however, only the first two column-headings have drop
     * 		-down buttons.
     * <LI><code><b>R12C1:R12C3</b></code> - drop-down buttons manifest at row 12;
     * 		this is probably not what you wanted.
     * </UL>
     * 
     * @param rcString	A String of R1C1 reference style.
     */
    void setAutoFilter(String rcString);
    
    /**
     * Sets the AutoFilter on this worksheet.
     * @param af the AutoFilter to be set on this worksheet.
     */
    void setAutoFilter(AutoFilter af);
    
    /**
     * Specifies whether the {@link #setAutoFilter(String) setAutoFilter}-method 
     * was applied on this worksheet.
     *  
     * @return true if the {@link #setAutoFilter(String) setAutoFilter}-method 
     * 				was applied on this worksheet, false otherwise.
     */
    boolean hasAutoFilter();
    
    /**
     * Removes the Autofilter on this worksheet.
     * 
     * @since xelem.2.1
     */
    void removeAutoFilter();
    
    //int rowCount();
    //int columnCount();
    
}
