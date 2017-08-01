/*
 * Created on 16-apr-2005
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

import java.util.List;

/**
 * Recieve notification of parsing events and the construction of 
 * {@link nl.fountain.xelem.excel.XLElement XLElements} and transmit these
 * events, values and instances to the listeners registered with this
 * ExcelReaderFilter.
 * 
 * @see <a href="package-summary.html#package_description">package overview</a>
 * @since xelem.2.0
 * 
 */
public interface ExcelReaderFilter extends ExcelReaderListener {
    
    /**
     * Gets a list of registered listeners.
     * 
     * @return a list of registered listeners
     */
    List<ExcelReaderListener> getListeners();
    
    /**
     * Registers the given listener.
     * @param listener the listener to be registered
     */
    void addExcelReaderListener(ExcelReaderListener listener);
    
    /**
     * Remove the specified listener.
     * @param listener the listener to be removed
     * @return <code>true</code> if the given listener was registered 
     * 		with this ExcelReaderFilter, <code>false</code> otherwise.
     */
    boolean removeExcelReaderListener(ExcelReaderListener listener);
    
    /**
     * Remove all registered listeners.
     *
     */
    void clearExcelReaderListeners();
}
