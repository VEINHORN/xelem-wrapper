/*
 * Created on 6-apr-2005
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

import java.util.Map;

import nl.fountain.xelem.excel.AutoFilter;
import nl.fountain.xelem.excel.Cell;
import nl.fountain.xelem.excel.Column;
import nl.fountain.xelem.excel.DocumentProperties;
import nl.fountain.xelem.excel.ExcelWorkbook;
import nl.fountain.xelem.excel.NamedRange;
import nl.fountain.xelem.excel.Row;
import nl.fountain.xelem.excel.Table;
import nl.fountain.xelem.excel.Worksheet;
import nl.fountain.xelem.excel.WorksheetOptions;

/**
 * Recieve notification of parsing events and the construction of 
 * {@link nl.fountain.xelem.excel.XLElement XLElements}.
 * A registered ExcelReaderListener listens to events fired by {@link ExcelReader}
 * while the ExcelReader is reading xml-spreadsheets of type spreadsheetML.
 * 
 * @see <a href="package-summary.html#package_description">package overview</a>
 * @since xelem.2.0
 */
public interface ExcelReaderListener {
    
    /**
     * Recieve notification of the start of the document.
     */
    void startDocument();
    
    /**
     * Recieve notification of processing instruction.
     * 
     * @param target 	the target of the processing instruction
     * @param data		the data of the processing instruction
     */
    void processingInstruction(String target, String data);
    
    /**
     * Recieve notification of the start of the Workbook tag.
     * 
     * @param systemID the systemID or "source" if no systemID was encountered
     */
    void startWorkbook(String systemID);
    
    /**
     * Recieve notification of the the construction of DocumentProperties.
     * 
     * @param docProps	fully populated instance of DocumentProperties
     */
    void setDocumentProperties(DocumentProperties docProps);
    
    /**
     * Recieve notification of the the construction of ExcelWorkbook.
     * 
     * @param excelWb	fully populated instance of ExcelWorkbook
     */
    void setExcelWorkbook(ExcelWorkbook excelWb);
    
    /**
     * Recieve notification of the the construction of a NamedRange on the 
     * workbook level.
     * 
     * @param namedRange	fully populated instance of NamedRange
     */
    void setNamedRange(NamedRange namedRange);
    
    /**
     * Recieve notification of the the start of a Worksheet tag.
     * 
     * @param sheetIndex 	the index of the encountered sheet. 0-based.
     * @param sheet			dummy instance of Worksheet. Only the (xml-)attributes
     * 	of the Worksheet-element have been set on the instance
     * 
     */
    void startWorksheet(int sheetIndex, Worksheet sheet);
    
    /**
     * Recieve notification of the the construction of a NamedRange on the 
     * worksheet level.
     * 
     * @param sheetIndex 	the index of the worksheet (0-based)
     * @param sheetName 	the name of the worksheet where the namedRange was found
     * @param namedRange	fully populated instance of NamedRange
     */
    void setNamedRange(int sheetIndex, String sheetName, NamedRange namedRange);
    
    /**
     * Recieve notification of the start of a Table tag.
     * 
     * @param sheetIndex 	the index of the worksheet (0-based)
     * @param sheetName 	the name of the worksheet where the table was found
     * @param table			dummy instance of Table. Only the (xml-)attributes
     * 	of the Table-element have been set on the instance
     */
    void startTable(int sheetIndex, String sheetName, Table table);
    
    /**
     * Recieve notification of the the construction of a Column. The column is fully
     * populated, it's column index has been set and can be obtained by
     * {@link nl.fountain.xelem.excel.Column#getIndex() column.getIndex()}.
     * 
     * @param sheetIndex 	the index of the worksheet (0-based)
     * @param sheetName 	the name of the worksheet where the column was found
     * @param column		fully populated instance of Column
     */
    void setColumn(int sheetIndex, String sheetName, Column column);
    
    /**
     * Recieve notification of the the construction of a Row. The row is fully
     * populated, it's row index has been set and can be obtained by
     * {@link nl.fountain.xelem.excel.Row#getIndex() row.getIndex()}.
     * 
     * @param sheetIndex	the index of the worksheet (0-based)
     * @param sheetName		the name of the worksheet where the row was found
     * @param row 			fully populated instance of Row
     */
    void setRow(int sheetIndex, String sheetName, Row row);
    
    /**
     * Recieve notification of the the construction of a Cell. The cell is fully
     * populated, it's cell index has been set and can be obtained by
     * {@link nl.fountain.xelem.excel.Cell#getIndex() cell.getIndex()}.
     * 
     * @param sheetIndex	the index of the worksheet (0-based)
     * @param sheetName		the name of the worksheet where the cell was found
     * @param rowIndex		the index of the row
     * @param cell 			fully populated instance of Cell
     */
    void setCell(int sheetIndex, String sheetName, int rowIndex, Cell cell);
    
    /**
     * Recieve notification of the the construction of WorksheetOptions.
     * 
     * @param sheetIndex 	the index of the worksheet (0-based)
     * @param sheetName 	the name of the worksheet where the worksheetOptions was found
     * @param wsOptions		fully populated instance of WorksheetOptions
     */
    void setWorksheetOptions(int sheetIndex, String sheetName, WorksheetOptions wsOptions);
    
    /**
     * Recieve notification of the the construction of AutoFilter.
     * 
     * @param sheetIndex 	the index of the worksheet (0-based)
     * @param sheetName 	the name of the worksheet where the autoFilter was found
     * @param autoFilter	fully populated instance of AutoFilter
     */
    void setAutoFilter(int sheetIndex, String sheetName, AutoFilter autoFilter);
    
    /**
     * Recieve notification of the end of a worksheet.
     * 
     * @param sheetIndex	the index of the worksheet (0-based)
     * @param sheetName		the name of the worksheet
     */
    void endWorksheet(int sheetIndex, String sheetName);
    
    /**
     * Recieve notification of the end of the document.
     * 
     * @param prefixMap a map of prefixes (keys) and uri's recieved while reading
     */
    void endDocument(Map<String, String> prefixMap);

}
