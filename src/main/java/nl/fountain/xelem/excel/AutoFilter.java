/*
 * Created on Nov 4, 2004
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
 * Represents the AutoFilter element. 
 * <P>
 * There's a lot more to autofiltering
 * than this interface provides. However, if it suffices to just set the
 * autofiltering-mode to 'on' for a certain range, here's all you need.
 * <P>
 * Usually you don't have to deal with implementations of this interface.
 * You just call the Worksheet-method 
 * {@link nl.fountain.xelem.excel.Worksheet#setAutoFilter(String)}.
 */
public interface AutoFilter extends XLElement {
    
    
    void setRange(String rcString);
    String getRange();

}
