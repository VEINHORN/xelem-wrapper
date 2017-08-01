/*
 * Created on 8-apr-2005
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

import java.io.File;

import nl.fountain.xelem.excel.AutoFilter;
import nl.fountain.xelem.excel.Column;
import nl.fountain.xelem.excel.DocumentProperties;
import nl.fountain.xelem.excel.ExcelWorkbook;
import nl.fountain.xelem.excel.NamedRange;
import nl.fountain.xelem.excel.Row;
import nl.fountain.xelem.excel.Table;
import nl.fountain.xelem.excel.Workbook;
import nl.fountain.xelem.excel.Worksheet;
import nl.fountain.xelem.excel.WorksheetOptions;
import nl.fountain.xelem.excel.ss.XLWorkbook;

/**
 * A concrete implementation of {@link ExcelReaderListener}. Listens to
 * events, values and instances fired by {@link ExcelReader} and (re)constructs
 * the {@link nl.fountain.xelem.excel.Workbook}. After the read has completed
 * the current workbook may be obtained by {@link #getWorkbook()}.
 * 
 * @since xelem.2.0
 */
public class WorkbookListener extends DefaultExcelReaderListener {
    
    private Workbook currentWorkbook;
    private Worksheet currentWorksheet;
    private Table currentTable;
    
    /**
     * Gets the current workbook.
     * 
     * @return the workbook that was read
     * 
     */
    public Workbook getWorkbook() {
        return currentWorkbook;
    }

    /**
     * Called by event dispatcher.
     * Recieve notification of the start of the Workbook tag.
     * Creates a new {@link nl.fountain.xelem.excel.ss.XLWorkbook}.
     * The workbookName may be set to a relevant portion of the passed 
     * <code>systemID</code>.
     * The fileName of the workbook is set to <code>systemID</code>.
     * The newly constructed Workbook is made the current workbook.
     * 
     * @param systemID the systemID or "source" if no systemID was encountered
     */
    public void startWorkbook(String systemID) {
        currentWorkbook = new XLWorkbook(getWorkbookName(systemID));
        currentWorkbook.setFileName(systemID);
    }
    
    /**
     * Called by event dispatcher.
     * Recieve notification of the the construction of DocumentProperties.
     * The passed docProps is added to the current workbook.
     * 
     * @param docProps	fully populated instance of DocumentProperties
     */
    public void setDocumentProperties(DocumentProperties docProps) {
        currentWorkbook.setDocumentProperties(docProps);
    }

    /**
     * Called by event dispatcher.
     * Recieve notification of the the construction of ExcelWorkbook.
     * The passed excelWb is added to the current workbook.
     * @param excelWb	fully populated instance of ExcelWorkbook
     */
    public void setExcelWorkbook(ExcelWorkbook excelWb) {
        currentWorkbook.setExcelWorkbook(excelWb);
    }

    /**
     * Called by event dispatcher.
     * Recieve notification of the the construction of a NamedRange on the 
     * workbook level.
     * The passed namedRange is added to the current workbook.
     * @param namedRange	fully populated instance of NamedRange
     */
    public void setNamedRange(NamedRange namedRange) {
        currentWorkbook.addNamedRange(namedRange);
    }

    /**
     * Called by event dispatcher.
     * Recieve notification of the the start of a Worksheet tag.
     * The passed sheet is added to the current workbook and is made the current
     * worksheet.
     * 
     * @param sheetIndex 	the index of the encountered sheet. 0-based.
     * @param sheet			dummy instance of Worksheet. Only the (xml-)attributes
     * 	of the Worksheet-element have been set on the instance
     * 
     */
    public void startWorksheet(int sheetIndex, Worksheet sheet) {
        currentWorksheet = sheet;
        currentWorkbook.addSheet(currentWorksheet);
    }

    /**
     * Called by event dispatcher.
     * Recieve notification of the the construction of a NamedRange on the 
     * worksheet level.
     * The passed namedRange is added to the current worksheet.
     * 
     * @param sheetIndex 	the index of the worksheet (0-based)
     * @param sheetName 	the name of the worksheet where the namedRange was found
     * @param namedRange	fully populated instance of NamedRange
     */
    public void setNamedRange(int sheetIndex, String sheetName, NamedRange namedRange) {
        currentWorksheet.addNamedRange(namedRange);
    }

    /**
     * Called by event dispatcher.
     * Recieve notification of the start of a Table tag.
     * The passed table is added to the current worksheet and is made the current
     * table.
     * 
     * @param sheetIndex 	the index of the worksheet (0-based)
     * @param sheetName 	the name of the worksheet where the table was found
     * @param table			dummy instance of Table. Only the (xml-)attributes
     * 	of the Table-element have been set on the instance
     */
    public void startTable(int sheetIndex, String sheetName, Table table) {
        currentTable = table;
        currentWorksheet.setTable(table);
    }

    /**
     * Called by event dispatcher.
     * Recieve notification of the the construction of a Column. The column is fully
     * populated, it's column index has been set and can be obtained by
     * {@link nl.fountain.xelem.excel.Column#getIndex() column.getIndex()}.
     * The passed column is added to the current table.
     * 
     * @param sheetIndex 	the index of the worksheet (0-based)
     * @param sheetName 	the name of the worksheet where the column was found
     * @param column		fully populated instance of Column
     */
    public void setColumn(int sheetIndex, String sheetName, Column column) {
        currentTable.addColumnAt(column.getIndex(), column);
    }

    /**
     * Called by event dispatcher.
     * Recieve notification of the the construction of a Row. The row is fully
     * populated, it's row index has been set and can be obtained by
     * {@link nl.fountain.xelem.excel.Row#getIndex() row.getIndex()}.
     * The passed row is added to the current table.
     * 
     * @param sheetIndex	the index of the worksheet (0-based)
     * @param sheetName		the name of the worksheet where the row was found
     * @param row 			fully populated instance of Row
     */
    public void setRow(int sheetIndex, String sheetName, Row row) {
        currentTable.addRowAt(row.getIndex(), row);
    }

    /**
     * Called by event dispatcher.
     * Recieve notification of the the construction of WorksheetOptions.
     * The passed wsOptions is added to the current worksheet.
     * 
     * @param sheetIndex 	the index of the worksheet (0-based)
     * @param sheetName 	the name of the worksheet where the worksheetOptions was found
     * @param wsOptions		fully populated instance of WorksheetOptions
     */
    public void setWorksheetOptions(int sheetIndex, String sheetName,
            WorksheetOptions wsOptions) {
        currentWorksheet.setWorksheetOptions(wsOptions);
    }

    /**
     * Called by event dispatcher.
     * Recieve notification of the the construction of AutoFilter.
     * The passed autoFilter is added to the current worksheet.
     * 
     * @param sheetIndex 	the index of the worksheet (0-based)
     * @param sheetName 	the name of the worksheet where the autoFilter was found
     * @param autoFilter	fully populated instance of AutoFilter
     */
    public void setAutoFilter(int sheetIndex, String sheetName,
            AutoFilter autoFilter) {
        currentWorksheet.setAutoFilter(autoFilter);
    }
    
    private String getWorkbookName(String systemId) {
        File file = new File(systemId);
        String[] s = file.getName().split("\\.");
        if (s.length > 0) {
            return s[0];
        } else {
            return "";
        }
    }

}
