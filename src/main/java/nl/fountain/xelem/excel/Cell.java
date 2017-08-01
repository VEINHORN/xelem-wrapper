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

import java.util.Date;

import org.w3c.dom.Document;
import org.w3c.dom.Element;


/**
 * Represents the Cell element. The XLElement Cell is not aware of it's parent,
 * nor of it's index-position in a row or worksheet. This makes it possible
 * to use the same instances of Cells on different worksheets, or different
 * places in the worksheet. Only at the time of
 * assembly the cell-index and the attribute ss:Index are automatically set, 
 * if necessary. See also: {@link nl.fountain.xelem.excel.Row#cellIterator()}.
 * (Also a Cell that is read by the nl.fountain.xelem.lex-API has it's index set.)
 * 
 * <P>
 * <h3>The setData-methods</h3>
 * The overloaded setData-methods will set the data displayed in the cell
 * when the xml spreadsheet is opened. These methods will set the Excel
 * datatype according to the java-type of the passed parameter. 
 * 
 * <P id="nullvalues">
 * <b>Null values</b><br>
 * If the passed parameter has a value of <code>null</code> the resulting
 * xml will have a datatype set to "Error",
 * the formula of the cell will be set to "<code>=#N/A</code>" and the cell will 
 * display "<code>#N/A</code>" when the spreadsheet is opened.
 * 
 * <P id="infinitevalues">
 * <b>Infinite values</b><br>
 * If the passed parameter is of type Double, Float or the primitive representation 
 * of these objects and the method {@link java.lang.Double#isInfinite() isInfinite}
 * results to <code>true</code> the xml will have a datatype set to
 * "String" and the cell will display "Infinite" when the spreadsheet is opened.
 * 
 * <P id="nanvalues">
 * <b>NaN values</b><br>
 * If the passed parameter is of type Double, Float or the primitive representation 
 * of these objects and the method {@link java.lang.Double#isNaN() isNaN}
 * results to <code>true</code> the resulting xml will have a datatype set to
 * "String" and the cell will display "NaN" when the spreadsheet is opened.
 */
public interface Cell extends XLElement {
    
    /**
     * The Excel datatype for number values: {@value}.
     */
    public static final String DATATYPE_NUMBER = "Number";
    
    /**
     * The Excel datatype for date and time values: {@value}.
     */
    public static final String DATATYPE_DATE_TIME = "DateTime";
    
    /**
     * The Excel datatype for boolean values: {@value}.
     */
    public static final String DATATYPE_BOOLEAN = "Boolean";
    
    /**
     * The Excel datatype for string values: {@value}.
     */
    public static final String DATATYPE_STRING = "String";
    
    /**
     * The Excel datatype for error values: {@value}.
     */
    public static final String DATATYPE_ERROR = "Error";
    
    /**
     * The Excel error value indicating division by zero: {@value}.
     */
    public static final String ERRORVALUE_DIV_BY_0 = "#DIV/0";
    
    /**
     * The Excel error value indicating text in a formula is not being 
     * recognized: {@value}.
     */
    public static final String ERRORVALUE_NAME = "#NAME?";
    
    /**
     * The Excel error value indicating 
     * intersection of two cell ranges is empty {@value}. 
     */
    public static final String ERRORVALUE_NULL = "#NULL!";
    
    /**
     * The Excel error value indicating problems with a number in a
     * formula or function: {@value}.
     */
    public static final String ERRORVALUE_NUM = "#NUM!";
    
    /**
     * The Excel error value indicating a value is not available: {@value}.
     */
    public static final String ERRORVALUE_NA = "#N/A";
    
    /**
     * The Excel error value indicating that a cell reference is not valid:
     * {@value}.
     */
    public static final String ERRORVALUE_REF = "#REF!";
    
    /**
     * The Excel error value indicating the wrong type of argument or operand
     * is being used: {@value}.
     */
    public static final String ERRORVALUE_VALUE = "#VALUE!";
       
    
    /**
     * Sets the ss:StyleID on this cell. If no styleID is set on a cell,
     * the ss:StyleID-attribute is not deployed in the resulting xml and
     * Excel employes the Default-style on the cell.
     * 
     * @param 	styleID	the id of the style to employ on this cell.
     */
    void setStyleID(String styleID);
    
    /**
     * Gets the ss:StyleID which was set on this cell.
     * 
     * @return 	The id of the style to employ on this cell or
     * 			<code>null</code> if no styleID was previously set.
     */
    String getStyleID();
    
    /**
     * Sets the ss:Formula-attribute on this cell. When opening a XML spreadsheet
     * the formula-result goes before any data-value that was set on the cell.
     * <PRE>
     *         cell.setData(3);
     *         cell.setFormula("=1+1");
     * </PRE>will display <code>2</code> in the opened spreadsheet.
     * 
     * @param 	formula	The formula for this cell.
     */
    void setFormula(String formula);
    
    /**
     * Gets the formula that was set on this cell.
     * 
     * @return 	The formula that was set on this cell or
     * 			<code>null</code> if the cell has no formula.
     */
    String getFormula();
    
    /**
     * Sets the ss:HRef-attribute on this cell. Any special html-characters
     * will be escaped in the reulting xml.
     * 
     * @param href	the ss:HRef-attribute on this cell.
     */
    void setHRef(String href);
    
    /**
     * Gets the ss:HRef-attribute value that was set on this cell.
     * 
     * @return 	The HRef that was set on this cell or
     * 			<code>null</code> if this cell has no HRef.
     */
    String getHRef();
    
    /**
     * Adds a comment to this cell. Initially the comment has no text
     * (no data and no author) and is not showing. If this cell allready
     * had a comment, replaces this comment.
     * 
     * @return the comment added to this cell
     * @since xelem.2.0
     */
    Comment addComment();
    
    /**
     * Adds the given comment to this cell. If this cell allready had a comment,
     * replaces this comment.
     * 
     * @param comment	the comment to be added to this cell
     * @return the newly added comment
     * @since xelem.2.0
     */
    Comment addComment(Comment comment);
    
    /**
     * Adds a comment to this cell and sets the data-attribute of the comment
     * to the given text. Initially the comment will have no author 
     * and is not showing. If this cell allready
     * had a comment, replaces this comment.
     * 
     * @param text the string to be displayed in the newly added comment
     * @return the newly added comment
     * @since xelem.2.0
     */
    Comment addComment(String text);
    
    /**
     * Specifies whether this cell has a comment.
     * 
     * @return <code>true</code> if this cell has a comment, false otherwise.
     * @since xelem.2.0
     */
    boolean hasComment();
    
    /**
     * Gets the comment of this cell.
     * 
     * @return the comment of this cell. May be <code>null</code>.
     * @since xelem.2.0
     */
    Comment getComment();
    
    /**
     * Merges this cell with adjacent cells on the same row. 
     * <P>
     * The value of 
     * <code>m</code> must be greater than 0 in order to have effect on
     * the resulting xml. No logic is included to prevent the merging of
     * cells that goes beyond the width of the spreadsheet. If such is the case,
     * the resulting xml spreadsheet will not open.
     * <P>
     * No other cells must be added within the scope of the merged cell:
     * <PRE>
     *         Cell c1 = sheet.addCellAt(1, 3);  // adds a cell at row 1, column 3
     *         c1.setMergeAcross(5);             // merges c1 across columns 3 - 8
     *         // illegal: first free index = 3 + 5 + 1 = 9
     *         // Cell c2 = sheet.addCellAt(1, 7);
     * </PRE>
     * 
     * @param 	m	The number of cells to include in the merge.
     */
    void setMergeAcross(int m);
    
    /**
     * Gets the number of cells over which this cell has been merged horizontally.
     * A return value of <code>0</code> indicates no horizontal merge 
     * was applied on this cell.
     * 
     * @return the number of cells that partake in the horizontal merge of this cell
     * @since xelem.2.0
     */
    int getMergeAcross();
    
    /**
     * Merges this cell with adjacent cells on the same column. 
     * <P>
     * The value of 
     * <code>m</code> must be greater than 0 in order to have effect on
     * the resulting xml.
     * <P>
     * No other rows or cells must be added within the scope of the merged cell:
     * <PRE>
     *         Cell c1 = sheet.addCellAt(1, 3);  // adds a cell at row 1, column 3
     *         c1.setMergeDown(5);             // merges c1 across rows 1 - 6
     *         // illegal: first free index = 1 + 5 + 1 = 7
     *         // Cell c2 = sheet.addCellAt(5, 3);
     * </PRE>
     * 
     * @param 	m	The number of cells to include in the merge.
     */
    void setMergeDown(int m);
    
    /**
     * Gets the number of cells over which this cell has been merged vertically.
     * A return value of <code>0</code> indicates no vertical merge 
     * was applied on this cell.
     * 
     * @return the number of cells that partake in the vertical merge of this cell
     * @since xelem.2.0
     */
    int getMergeDown();
    
    /**
     * Sets the given Number as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "Number". 
     * 
     * @param 	data The data to be displayed in this cell. 
     * 
     * @see <a href="#nullvalues">Null values</a>
     */
    void setData(Number data);
    
    /**
     * Sets the given Integer as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "Number". 
     * 
     * @param 	data The data to be displayed in this cell. 
     * 
     * @see <a href="#nullvalues">Null values</a>
     */
    void setData(Integer data);
    
    /**
     * Sets the given Double as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "Number". 
     * 
     * @param 	data The data to be displayed in this cell. 
     * 
     * @see <a href="#nullvalues">Null values</a>
     * @see <a href="#infinitevalues">Infinite values</a>
     * @see <a href="#nanvalues">NaN values</a>
     */
    void setData(Double data);
    
    /**
     * Sets the given Long as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "Number". 
     * 
     * @param 	data The data to be displayed in this cell. 
     * 
     * @see <a href="#nullvalues">Null values</a>
     */
    void setData(Long data);
    
    /**
     * Sets the given Float as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "Number". 
     * 
     * @param 	data The data to be displayed in this cell. 
     * 
     * @see <a href="#nullvalues">Null values</a>
     * @see <a href="#infinitevalues">Infinite values</a>
     * @see <a href="#nanvalues">NaN values</a>
     */
    void setData(Float data);
    
    /**
     * Sets the given Date as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "DateTime". 
     * 
     * @param 	data The data to be displayed in this cell. 
     * 
     * @see <a href="#nullvalues">Null values</a>
     */
    void setData(Date data);
    
    /**
     * Sets the given Boolean as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "Boolean". 
     * 
     * @param 	data The data to be displayed in this cell. 
     * 
     * @see <a href="#nullvalues">Null values</a>
     */
    void setData(Boolean data);
    
    /**
     * Sets the given String as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "String". 
     * 
     * @param 	data The data to be displayed in this cell. 
     * 
     * @see <a href="#nullvalues">Null values</a>
     */
    void setData(String data);
    
    /**
     * Sets the given object as the data of this cell. This method reflects
     * upon the class of the given object and will delegate to a corresponding
     * setData-method if such a method is available. If no corresponding
     * method is found this method sets the data of this cell to the
     * <code>toString</code>-value of the given object and the datatype to
     * "String".
     * 
     * @param 	data The data to be displayed in this cell.
     * 
     * @see <a href="#nullvalues">Null values</a>  
     */
    void setData(Object data);
    
    /**
     * Sets the given error value as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "Error". The formula
     * of this cell will be set to the corresponding error value.
     * 
     * @param error_value	Must be one of Cell's ERRORVALUE_XXX values.
     */
    void setError(String error_value);
    
    /**
     * Sets the given byte as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "Number". 
     * 
     * @param 	data The data to be displayed in this cell. 
     */
    void setData(byte data);
    
    /**
     * Sets the given short as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "Number". 
     * 
     * @param 	data The data to be displayed in this cell. 
     */
    void setData(short data);
    
    /**
     * Sets the given int as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "Number". 
     * 
     * @param 	data The data to be displayed in this cell. 
     */
    void setData(int data);
    
    /**
     * Sets the given long as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "Number". 
     * 
     * @param 	data The data to be displayed in this cell. 
     */
    void setData(long data);
    
    /**
     * Sets the given float as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "Number". 
     * 
     * @param 	data The data to be displayed in this cell. 
     * 
     * @see <a href="#infinitevalues">Infinite values</a>
     * @see <a href="#nanvalues">NaN values</a>
     */
    void setData(float data);
    
    /**
     * Sets the given double as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "Number". 
     * 
     * @param 	data The data to be displayed in this cell. 
     * 
     * @see <a href="#infinitevalues">Infinite values</a>
     * @see <a href="#nanvalues">NaN values</a>
     */
    void setData(double data);
    
    /**
     * Sets the data of this cell to the value
     * {@link java.lang.String#valueOf(char)}. The attribute
     * ss:Type of the Data-element will be set to "String". 
     * 
     * @param 	data The data to be displayed in this cell. 
     * 
     */
    void setData(char data); 
    
    /**
     * Sets the given boolean as the data of this cell. The attribute
     * ss:Type of the Data-element will be set to "Boolean". 
     * 
     * @param 	data The data to be displayed in this cell. 
     */
    void setData(boolean data);
    
    /**
     * Gets the NodeValue of the Data-tag as a String.
     * 
     * @return The NodeValue of the Data-tag as a String.
     */
    String getData$();
    
    /**
     * Gets the value of the ss:Type-attribute of the Data-element.
     * 
     * @return The value of the ss:Type-attribute of the Data-element.
     */
    String getXLDataType();
    
    /**
     * Specifies whether this cell has data.
     * 
     * @return <code>true</code> if this cell has a data element,
     * 		<code>false</code> otherwise
     * @since xelem.2.0
     */
    boolean hasData();
    
    /**
     * Specifies whether this cell has an error.
     * 
     * @return <code>true</code> if the 'Type' attribute of this cells
     * 	data element equals DATATYPE_ERROR, <code>false</code> otherwise
     */
    boolean hasError();
    
    /**
     * Gets the value of the data element of this cell as an object.
     * The type of the returned object is determined by the value of the
     * 'Type' attribute of this cells data element.
     * <P>
     * <TABLE border="1" cellpadding="5" cellspacing="4">
     * <TR><TH>xlDataType</TH><TH>returned type</TH></TR>
     * <TR><TD>Number</TD><TD>{@link java.lang.Double}</TD></TR>
     * <TR><TD>DateTime</TD><TD>{@link java.util.Date}</TD></TR>
     * <TR><TD>Boolean</TD><TD>{@link java.lang.Boolean}</TD></TR>
     * <TR><TD>String</TD><TD>{@link java.lang.String}</TD></TR>
     * <TR><TD>Error</TD><TD>{@link java.lang.String}</TD></TR>
     * <TR><TD>(not specified)</TD><TD>{@link java.lang.String}</TD></TR>
     * </TABLE>
     * 
     * @return	the value of the data element of this cell as an object
     * @since xelem.2.0
     */
    Object getData();
    
    /**
     * Convenience method for reading the value of the data element 
     * of this cell as an <code>int</code>. Returns <code>0</code>
     * if the value could not be converted.
     * 
     * @return the value of the data element
     * @since xelem.2.0
     */
    int intValue();
    
    /**
     * Convenience method for reading the value of the data element 
     * of this cell as a <code>double</code>. Returns <code>0.0</code>
     * if the value could not be converted.
     * 
     * @return the value of the data element
     * @since xelem.2.0
     */
    double doubleValue();
    
    /**
     * Convenience method for reading the value of the data element 
     * of this cell as a <code>boolean</code>. Returns <code>false</code>
     * if the value is not equal to "1".
     * 
     * @return the value of the data element
     * @since xelem.2.0
     */
    boolean booleanValue();
    
    /**
     * Gets the Data-element of this cell as an
     * {@link org.w3c.dom.Element org.w3c.dom.Element}.
     * @param doc the document to which the returned element may be appended
     * @return the Data-element
     */
    Element getDataElement(Document doc);
    
    /**
     * Sets the value of the ss:Index-attribute of this Cell-element. 
     * Any value set may be overruled  
     * by {@link nl.fountain.xelem.excel.Row#cellIterator()}, which sets the
     * index of cells during assembly. 
     * <P>
     * <em>
     * If you want to place a cell in any particular place, use the
     * addCellAt-methods of Row, Table or Worksheet, or use Worksheets
     * CellPointer.
     * </em>
     * <P>
     * 
     * @param index the index of this cell
     */
    void setIndex(int index);
    
    /**
     * Gets the value of the ss:Index-attribute of this Cell-element.
     * The returned value only makes sence if this cell was read with the
     * reader-API in the nl.fountain.xelem.lex-package and before
     * any manipulation of this cell or the workbook took place.
     * 
     * @return the value of the ss:Index-attribute of this Cell-element
     * @since xelem.2.0
     */
    int getIndex();
}
