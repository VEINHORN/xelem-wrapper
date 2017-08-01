/*
 * Created on Oct 28, 2004
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
 * Represents the Column element. Only add columns
 * if formatting of columns is necessary.
 * <P>
 * The XLElement Column is not aware of it's parent,
 * nor of it's index-position in a table or worksheet. This makes it possible
 * to use the same instances of columns on different worksheets, or different
 * places in the worksheet. Only at the time of
 * assembly the column-index and the attribute ss:Index are automatically set, 
 * if necessary. See also: {@link nl.fountain.xelem.excel.Table#columnIterator()}.
 * (Also a Column that is read by the nl.fountain.xelem.lex-API has it's index set.)
 * 
 */
public interface Column extends XLElement {
    
    /**
     * Sets the ss:StyleID on this column. If no styleID is set on a column,
     * the ss:StyleID-attribute is not deployed in the resulting xml and
     * Excel employes the Default-style on the column.
     * 
     * @param 	id	the id of the style to employ on this column.
     */
    void setStyleID(String id);
    
    /**
     * Gets the ss:StyleID which was set on this column.
     * 
     * @return 	The id of the style to employ on this column or
     * 			<code>null</code> if no styleID was previously set.
     */
    String getStyleID();
    
    /**
     * Sets column-width to autofit for datatypes 'DateTime' and 'Number'. 
     * The default setting of Excel is 'true', so this method only has effect
     * when <code>false</code> is passed as a parameter.
     * 
     * @param autoFit	<code>false</code> if no autofit of column-width is
     * 					desired.
     */
    void setAutoFitWidth(boolean autoFit);
    
    /**
     * Specifies whether autofit was set on this column.
     * 
     * @return	<code>true</code> if column-width is set to autofit,
     * 		<code>false</code> otherwise 
     */
    boolean getAutoFitWith();
    
    /**
     * Sets the span of this column.
     * The value of 
     * <code>s</code> must be greater than 0 in order to have effect on
     * the resulting xml.
     * <P>
     * No other columns must be added within the span of this column:
     * <PRE>
     *         Column column = table.addColumn(5); // adds a column with index 5
     *         column.setWidth(25.2);              // do some formatting on the column
     *         column.setSpan(3);                  // span a total of 4 columns
     *         // illegal: first free index = 5 + 3 + 1 = 9
     *         // Column column2 = table.addColumn(8);
     * </PRE>
     * 
     * @param 	s	The number of additional columns to include in the span.
     */
    void setSpan(int s);
    
    /**
     * Gets the number of extra columns this column spans.
     * 
     * @return	the number of extra columns this column spans
     */
    int getSpan();
    
    /**
     * Sets the width of this column.
     * 
     * @param w	The width of the column (in points).
     */
    void setWidth(double w);
    
    /**
     * Gets the width of this column. A return value of <code>0.0</code>
     * may indicate this column has a default width.
     * 
     * @return	the width of this column
     */
    double getWidth();
    
    /**
     * Sets whether this column will be hidden.
     * 
     * @param hide	<code>true</code> if this column must not be displayed.
     */
    void setHidden(boolean hide);
    
    /**
     * Specifies whether this column is hidden.
     */
    boolean isHidden();
    
    /**
     * Sets the value of the ss:Index-attribute of this Column-element. 
     * Any value set may be overruled  
     * by {@link nl.fountain.xelem.excel.Table#columnIterator()}, which sets the
     * index of columns during assembly. 
     * <P>
     * <em>
     * If you want to place a column in any particular place, use the
     * addColumnAt-methods of Table or Worksheet.
     * </em>
     * <P>
     * 
     * @param index the index of this column
     */
    void setIndex(int index);
    
    /**
     * Gets the value of the ss:Index-attribute of this Column-element.
     * The returned value only makes sence if this column was read with the
     * reader-API in the nl.fountain.xelem.lex-package and before
     * any manipulation of this column or the workbook took place.
     * 
     * @return the value of the ss:Index-attribute of this Column-element
     * @since xelem.2.0
     */
    int getIndex();
}
