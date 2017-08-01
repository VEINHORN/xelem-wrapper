/*
 * Created on 31-mrt-2005
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
 *
 */
package nl.fountain.xelem.excel;

/**
 * Represents the Comment element.
 */
public interface Comment extends XLElement {
    
    void setAuthor(String author);
    String getAuthor();
    
    void setShowAlways(boolean show);
    boolean showsAlways();
    
    void setData(String data);
    String getData();
    
    /**
     * Gets the content of the data element stripped of the author (if
     * there was any).
     * 
     * @return the content of the data element stripped of the author
     */
    String getDataClean();

}
