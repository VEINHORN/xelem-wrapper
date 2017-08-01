/*
 * Created on 5-apr-2005
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
package nl.fountain.xelem;

/**
 * A reference to a rectangular range.
 * @since xelem.2.0
 */
public class Area {
    
    /**
     * The top-row of this Area.
     */
    protected int r1;
    
    /**
     * The bottom-row of this Area.
     */
    protected int r2;
    
    /**
     * The left-most column of this Area.
     */
    protected int c1;
    
    /**
     * The right-most column of this Area.
     */
    protected int c2;
    
    /**
     * Constructs a new Area.
     * @param row1		the index of the top row
     * @param column1	the index of the left-most column
     * @param row2		the index of the bottom row
     * @param column2	the index of the right-most column
     */
    public Area(int row1, int column1, int row2, int column2) {
        setDimensions(row1, column1, row2, column2);
    }
    
    /**
     * Constructs a new Area.
     * @param address1	the address in one corner of the area
     * @param address2	the address in the opposite corner
     */
    public Area(Address address1, Address address2) {
        setDimensions(address1.r, address1.c, address2.r, address2.c);
    }
    
    /**
     * Constructs a new Area. This constructor takes an A1-reference string 
     * as a parameter. For instance "B4:E6" will construct an area delimmited
     * by rows 4 and 6 and columns 2 and 5. Treats the passed string case-insensitive.
     * <P>
     * Splits the string in the parameter on ":" and calculates row and
     * column numbers with the 2 resulting strings.
     * 
     * @param a1_ref a string of A1-reference style
     * @throws IllegalArgumentException if a1_ref cannot be split in two strings
     * 		using ":" as breakpoint.
     * @see Address#calculateColumn(String)
     * @see Address#calculateRow(String)
     */
    public Area(String a1_ref) {
        String[] ar = a1_ref.split(":");
        if (ar.length < 2) {
            throw new IllegalArgumentException("use format 'A1:A1'.");
        }
        setDimensions(Address.calculateRow(ar[0]), Address.calculateColumn(ar[0]),
                Address.calculateRow(ar[1]), Address.calculateColumn(ar[1]));
    }
    
    /**
     * Gets the index of the top row of this area.
     * @return	the index of the top row.
     */
    public int getFirstRow() {
        return r1;
    }
    
    /**
     * Gets the index of the bottom row of this area.
     * @return the index of the bottom row.
     */
    public int getLastRow() {
        return r2;
    }
    
    /**
     * Gets the index of the left-most column of this area.
     * @return the index of the left-most column.
     */
    public int getFirstColumn() {
        return c1;
    }
    
    /**
     * Gets the index of the right-most column of this area.
     * @return the index of the right-most column
     */
    public int getLastColumn() {
        return c2;
    }
    
    /**
     * Gets a string in A1-reference style denoting the range of this area.
     * @return a string in A1-reference style
     */
    public String getA1Reference() {
        StringBuffer sb = new StringBuffer();
        sb.append(Address.calculateColumn(c1));
        sb.append(r1);
        sb.append(":");
        sb.append(Address.calculateColumn(c2));
        sb.append(r2);
        return sb.toString();
    }
    
    /**
     * Gets a string in R1C1-reference style denoting the range of this area.
     * @return a string in R1C1-reference style
     */
    public String getAbsoluteRange() {
        StringBuffer sb = new StringBuffer("R");
        sb.append(r1);
        sb.append("C");
        sb.append(c1);
        sb.append(":R");
        sb.append(r2);
        sb.append("C");
        sb.append(c2);
        return sb.toString();
    }
    
    /**
     * Specifies whether the intersection of the given row and column index
     * is within this area.
     * @param rowIndex		the rownumber
     * @param columnIndex	the columnnumber
     * @return <code>true</code> if the intersection is within this area,
     * 		<code>fals</code> otherwise.
     */
    public boolean isWithinArea(int rowIndex, int columnIndex) {
        return rowIndex >= r1 && rowIndex <= r2 
        	&& columnIndex >= c1 && columnIndex <= c2;
    }
    
    /**
     * Specifies whether the given address is within this area.
     * @param address the address to be investigated
     * @return <code>true</code> if the address is within this area,
     * 		<code>false</code> otherwise.
     */
    public boolean isWithinArea(Address address) {
        return isWithinArea(address.r, address.c);
    }
    
    /**
     * Specifies whether the given row is within this area.
     * @param rowIndex	the rownumber
     * @return	<code>true</code> if the row index is part of this area,
     * 		<code>false</code> otherwise.
     */
    public boolean isRowPartOfArea(int rowIndex) {
        return rowIndex >= r1 && rowIndex <= r2;
    }
    
    /**
     * Specifies whether the given column is within this area.
     * @param columnIndex	the columnnumber
     * @return	<code>true</code> if the column index is part of this area,
     * 		<code>false</code> otherwise.
     */
    public boolean isColumnPartOfArea(int columnIndex) {
        return columnIndex >= c1 && columnIndex <= c2;
    }
    

    private void setDimensions(int row1, int column1, int row2, int column2) {
        if (row2 > row1) {
            r1 = row1;
            r2 = row2;
        } else {
            r1 = row2;
            r2 = row1;
        }
        if (column2 > column1) {
            c1 = column1;
            c2 = column2;
        } else {
            c1 = column2;
            c2 = column1;
        }
    }
    

}
