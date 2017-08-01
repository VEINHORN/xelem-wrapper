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

import java.util.List;
import java.util.Map;

import javax.xml.parsers.ParserConfigurationException;

import nl.fountain.xelem.UnsupportedStyleException;

import org.w3c.dom.Document;

/**
 * Represents the Workbook element, the root of a SpreadsheetML document.
 * <P>
 * Creating SpreadsheetML and Excel workbooks from within Java starts here.
 * Besides adding child XLElements and setting other features,
 * the Workbook can be used to merge Styles.
 * After finishing the Workbook you'll probably want to create the 
 * {@link org.w3c.dom.Document}, the structure that holds this Workbook
 * and all of it's children. This is done with the 
 * {@link #createDocument()}-method.
 * <P>
 * A Workbook might also be obtained from an
 * {@link nl.fountain.xelem.lex.ExcelReader},
 * an object that reads Excel xml-files or xml-InputSources.
 * 
 * 
 * 
 * @see <a href="../../../../overview-summary.html#overview_description">overview</a>
 */
public interface Workbook extends XLElement {
    
    /**
     * Sets the name of this Workbook.
     */
    void setName(String name);
    
    /**
     * Gets the name of this Workbook.
     */
    String getName();
    
    /**
     * Sets the filename of this Workbook. 
     */
    void setFileName(String filename);
    
    /**
     * Gets the filename of this Workbook. If no filename was supplied with
     * {@link #setFileName(String)} the name of the workbook is returned
     * with the extension '.xls'.
     */
    String getFileName();
    
    /**
     * Merges two SpreadsheetML Style elements.
     * 
     * @param 	newID	the ss:ID of the new merged style
     * @param	id1		the ss:ID of the copied style
     * @param 	id2		the ss:ID of the appended style
     * 
     * @throws 	UnsupportedStyleException if either of the styles mentioned with 
     * 			id1 or id2 are unknown to the factory.
     * 
     * @see nl.fountain.xelem.XFactory#mergeStyles(String, String, String)
     */
    void mergeStyles(String newID, String id1, String id2) 
    	throws UnsupportedStyleException;
    
    /**
     * Sets whether this Workbook's {@link #createDocument()}-method
     * will print comments which were set with
     * {@link nl.fountain.xelem.excel.XLElement#addElementComment(String)}.
     * <P>
     * The default is true.
     * 
     * @see nl.fountain.xelem.excel.XLElement#addElementComment(String)
     */
    void setPrintElementComments(boolean print);
    
    /**
     * Specifies whether this Workbook's {@link #createDocument()}-method
     * will include comments.
     */
    boolean isPrintingElementComments();
    
    /**
     * Sets whether this Workbook's {@link #createDocument()}-method
     * will include docComments. DocComments are the comments that will 
     * be passed in at the start of the Workbook document, 
     * just after the processing instruction.
     * <P>
     * DocComments can be configured in the configuration file, standard at
     * config/xelem.xml.
     * <P>
     * The default is true.
     * 
     * @see nl.fountain.xelem.XFactory#getDocComments()
     */
    void setPrintDocComments(boolean print);
    
    /**
     * Specifies whether this Workbook's {@link #createDocument()}-method
     * will include docComments. 
     */
    boolean isPrintingDocComments();
    
    /**
     * Appends a Worksheet with general information to this Workbook. 
     * The infoSheet is allways displayed as the last worksheet.
     */
    void appendInfoSheet();
    
    /**
     * Set the DocumentProperties of this workbook or, if this workbook
     * allready had DocumentProperties, replace it.
     * 
     * @param docProps	the DocumentProperties for this workbook
     */
    void setDocumentProperties(DocumentProperties docProps);
    
    /**
     * Gets the DocumentProperties of this workbook. 
     * 
     * @return The DocumentProperties of this workbook. Never <code>null</code>.
     */
    DocumentProperties getDocumentProperties();
    
    /**
     * Specifies whether this workbook has DocumentProperties.
     */
    boolean hasDocumentProperties();
    
    /**
     * Set the ExcelWorkbook of this workbook or, if this workbook allready had
     * an ExcelWorkbook, replace it.
     * 
     * @param excelWb	The ExcelWorkbook for this workbook
     */
    void setExcelWorkbook(ExcelWorkbook excelWb);
    
    /**
     * Gets the ExcelWorkbook of this workbook. 
     * 
     * @return The ExcelWorkbook of this workbook. Never <code>null</code>.
     */
    ExcelWorkbook getExcelWorkbook();
    
    /**
     * Specifies whether this workbook has an ExcelWorkbook.
     */
    boolean hasExcelWorkbook();
    
    /**
     * Adds a new NamedRange to this workbook. Named ranges are usefull when
     * working with formulas. 
     * <P>
     * The string <code>refersTo</code> should be of format
     * <code>[worksheet name]!R1C1:R1C1</code>.
     * If this workbook allready contained a NamedRange with the given name,
     * replaces it.
     * 
     * @param name		The name to apply to the range.
     * @param refersTo	A String of R1C1 reference style, including worksheet name.
     * 
     * @return New NamedRange.
     */
    NamedRange addNamedRange(String name, String refersTo);
    
    /**
     * Adds a new NamedRange to this workbook. 
     * If this workbook allready contained a NamedRange with an equal name,
     * replaces it.
     * 
     * @param namedRange	the NamedRange to be added to this workbook
     * @return	the given NamedRange
     */
    NamedRange addNamedRange(NamedRange namedRange);
    
    /**
     * Gets all the NamedRanges that were added to this workbook. The map-keys
     * are equal to their names.
     * @return 	a map with NamedRanges
     */
    Map<String, NamedRange> getNamedRanges();
    
    /**
     * Adds a new Worksheet to this workbook. The new sheet will have 
     * the name "Sheet[x]", where [x] is 
     * a sequential number to prevent the duplication of sheet names.
     * The worksheets cellpointer will be at position row 1, column 1.
     * 
     * @return A new Worksheet.
     */
    Worksheet addSheet();
    
    /**
     * Adds a new Worksheet to this workbook with the given name as name.
     * The worksheets cellpointer will be at position row 1, column 1.
     * 
     * @param name	The name that is to be given to the new worksheet.
     * 
     * @return A new Worksheet.
     * 
     * @throws DuplicateNameException if a sheet with this name is allready 
     * 			present in this workbook.
     */
    Worksheet addSheet(String name);
    
    /**
     * Adds the given Worksheet to this workbook.
     * 
     * @param sheet	The worksheet to be added.
     * 
     * @return The passed Worksheet.
     * 
     * @throws DuplicateNameException if a sheet with such a name is allready 
     * 			present in this workbook.
     */
    Worksheet addSheet(Worksheet sheet); 
    
    /**
     * Removes the sheet with the given name.
     * 
     * @param name	The name of the worksheet to remove.
     * 
     * @return The removed worksheet or <code>null</code> if a worksheet with
     * 			the given name was not in this workbook.
     */
    Worksheet removeSheet(String name); 
    
    /**
     * Gets a list of all the worksheets in this workbook in display order.
     * 
     * @return A list of worksheets.
     */
    List<Worksheet> getWorksheets();
    
    /**
     * Gets a list of all the names of worksheets in this workbook in
     * display order. This list may be used to manipulate the display
     * order of sheets.
     * For instance:
     * <PRE>              Collections.swap(workbook.getSheetNames(), 0, 2);</PRE>
     * will display the first worksheet added to 
     * this workbook at position 3.
     * <P>
     * No objects should be added to this list.
     * 
     * @return A list of sheet names.
     */
    List<String> getSheetNames();
    
    /**
     * Gets the worksheet with the given name.
     * 
     * @param name	The name of the worksheet.
     * 
     * @return The worksheet with the given name or <code>null</code> if
     * 			no worksheet with that name was in this workbook.
     */
    Worksheet getWorksheet(String name);
    
    /**
     * Gets the worksheet at the given index. The worksheet index is 0-based.
     * 
     * @param index the worksheet index number
     * @return	the worksheet at the given index or <code>null</code> if
     * 			no worksheet was at that index
     */
    Worksheet getWorksheetAt(int index);
    
    /**
     * Creates a {@link org.w3c.dom.Document}, the structure that holds this Workbook
     * and all of it's children. 
     * <P>
     * During assembly
     * non fatal-errors will be registered as warnings. See {@link #getWarnings()}.
     * A new call to createDocument erases all previous warnings.
     * <P>
     * Non fatal-errors could be:
     * <UL>
     * <LI>The {@link nl.fountain.xelem.XFactory} could not be configured.
     * <LI>There was no definition for a 'Default' style in the XFactory. The
     * 	workbook will create a standard 'Default' style.
     * <LI>A cell, column, row or table was given a styleID, but the corresponding
     * 	definition of that style was not found in the XFactory. The workbook will
     * 	create an empty definition of that style 
     * (i.e. <code>&lt;Style ss:ID="foo"/&gt;</code>).
     * <LI>The method {@link #appendInfoSheet()} was called, but an infoSheet-Node
     * 	could not be created.
     * </UL>
     * 
     * 
     * @return A {@link org.w3c.dom.Document}, representing this workbook.
     * 
     * @throws ParserConfigurationException if the document could not be created.
     */
    Document createDocument() throws ParserConfigurationException;
    
    /**
     * Gets a list of warnings (Strings) that were registered during assembly
     * of this workbook after a call to {@link #createDocument()}.
     * 
     * @return A list of warnings. The list may be empty if no warnings were
     * 			generated.
     */
    List<String> getWarnings();
    

}
