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
 * Does effectively nothing. It's up to the subclasses of this class to do
 * effectively something.
 * 
 * @since xelem.2.0
 */
public class DefaultExcelReaderListener implements ExcelReaderListener {
    
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void startDocument() {
    }
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void processingInstruction(String target, String data) {
    }
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void startWorkbook(String systemID) {
    }
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void setDocumentProperties(DocumentProperties docProps) {
    }
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void setExcelWorkbook(ExcelWorkbook excelWb) {
    }
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void setNamedRange(NamedRange namedRange) {
    }
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void startWorksheet(int sheetIndex, Worksheet sheet) {
    }
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void setNamedRange(int sheetIndex, String sheetName,
            NamedRange namedRange) {
    }
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void startTable(int sheetIndex, String sheetName, Table table) {
    }
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void setColumn(int sheetIndex, String sheetName, Column column) {
    }
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void setRow(int sheetIndex, String sheetName, Row row) {
    }
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void setCell(int sheetIndex, String sheetName, int rowIndex,
            Cell cell) {
    }
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void setWorksheetOptions(int sheetIndex, String sheetName,
            WorksheetOptions wsOptions) {
    }
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void setAutoFilter(int sheetIndex, String sheetName,
            AutoFilter autoFilter) {
    }
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void endWorksheet(int sheetIndex, String sheetName) {
    }
    /**
     * Does effectively nothing. Subclasses can override this method
     * to do effectively something.
     */
    public void endDocument(Map<String, String> prefixMap) {
    }

}
