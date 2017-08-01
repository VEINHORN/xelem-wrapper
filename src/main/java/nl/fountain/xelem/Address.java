/*
 * Created on 30-nov-2004
 * Copyright (C) 2004, 2005  Henk van den Berg
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

import java.util.Collection;
import java.util.Set;
import java.util.TreeSet;

import nl.fountain.xelem.excel.Worksheet;

/**
 * A reference to the intersection of a row and a column. Besides that Address can
 * be used to get R1C1-reference strings, which can be used in formulas and 
 * NamedRanges.
 */
public class Address implements Comparable {
    
     /**
      * The row index of this address.
      */
    protected int r;
    
    /**
     * The column index of this address.
     */
    protected int c;
    
    private static int ABC_RADIX = 26;
    
    /**
     * This constructor is protected.
     */
    protected Address() {}
    
    /**
     * Constructs a new Address.
     * @param rowIndex		The row index of this Address.
     * @param columnIndex	The column index of this Address.
     */
    public Address(int rowIndex, int columnIndex) {
        r = rowIndex;
        c = columnIndex;
    }
    
    /**
     * Constructs a new Address.
     * The given string can be of A1-reference style, as used in Excel.
     * The parameter "A1" constructs an address at row 1, column 1;
     * "Z23" leads to an address pointing to row 23, column 26.
     * <P>
     * This constructor treats the passed string case-insensitive,
     * row indicators (digits) and column indicators (letters) may be intermingled. 
     * The next equations all evaluate as <code>true</code>.
     * <PRE>
     *   new Address("BQ65").equals(new Address("65bq"))
     *   new Address("BQ65").equals(new Address("6B5q"))
     *   etc.
     * </PRE>
     * 
     * @param a1_ref	a String in A1-reference style
     * @since xelem.2.0
     */
    public Address(String a1_ref) {
        r = calculateRow(a1_ref);
        c = calculateColumn(a1_ref);
    }
    
    /**
     * Calculates the column number of a given string in A1-reference style.
     * Column indicator (letters) and row indicator (digits) may be intermingled.
     * Treats the passed string case-insensitive. The maximum column notation
     * that can be calculated is "FXSHRXW" and this string returns
     * {@link java.lang.Integer#MAX_VALUE}.
     *  If no letters are present in the given string this method returns 0.
     * 
     * @param s a string in A1-reference style
     * @return the column number of the given string
     * @since xelem.2.0
     */
    public static int calculateColumn(String s) {
        String su = s.toUpperCase();
        int colnr = 0;
        int factor = 1;
        for (int i = su.length() - 1; i >= 0; i--) {
            char ch = su.charAt(i);
            if (Character.isLetter(ch)) {
                colnr += (ch - 64) * factor;
                factor *= ABC_RADIX;
            }
        }
        return colnr;
    }
    
    /**
     * Calculates the column notation in A1-reference style of a given column number.
     * A column number of 0 or less returns as an empty string ("").
     * A column number equal to {@link java.lang.Integer#MAX_VALUE} returns
     * "FXSHRXW".
     * <P>
     * It may be interesting to know that a parameter of <code>1000</code>
     * returns "ALL" and that multiplying this parameter with a factor
     * of <code>676.149</code> yields the dutch translation: "ALLES".
     * Unfortunately not all words can be translated using the same factor and
     * a more practical use then of this method is to feed it column numbers
     * of 1 to 256 inclusive and get the Excel label of the column, 
     * "A" to "IV" inclusive, in return.
     * 
     * @param columnNumber the column number to be calculated
     * @return the column notation in A1-reference style
     * @since xelem.2.0
     */
    public static String calculateColumn(int columnNumber) {
        int div = 1;
        int af = 0;
        StringBuffer sb = new StringBuffer();
        int q;
        while ((q = (columnNumber-af)/div) > 0) {
            sb.insert(0, getDigit(q));
            af += div;
            div *= ABC_RADIX;            
        }
        return sb.toString();
    }
    
    private static char getDigit(int q) {
        int r = q % ABC_RADIX;
        if (r == 0) r = ABC_RADIX;
        return (char) (r + 64);
    }
    
    /**
     * Calculates the row number of a given string in A1-reference style.
     * Column indicator (letters) and row indicator (digits) may be intermingled.
     * If no digits are present in the given string this method returns 0.
     * 
     * @param s a string in A1-reference style
     * @return the row number of the given string
     * @since xelem.2.0
     */
    public static int calculateRow(String s) {
        int rownr = 0;
        int factor = 1;
        for (int i = s.length() -1; i >=0; i--) {
            char ch = s.charAt(i);
            if (Character.isDigit(ch)) {
                rownr += (ch - 48) * factor;
                factor *= 10;
            }
        }
        return rownr;
    }

    /**
     * Gets the index of the row of this address.
     * 
     * @return the index of the row of this address
     */
    public int getRowIndex() {
        return r;
    }

    /**
     * Gets the index of the column of this address.
     * 
     * @return the index of the column of this address
     */
    public int getColumnIndex() {
        return c;
    }

    /**
     * Specifies whether this address is within the bounds of the
     * spreadsheet.
     * @return <code>true</code> if this address is within the bounds of
     * 	the sheet, <code>false</code> otherwise
     */
    public boolean isWithinSheet() {
        return c >= Worksheet.firstColumn && c <= Worksheet.lastColumn 
        	&& r >= Worksheet.firstRow && r <= Worksheet.lastRow;
    }
    
    /**
     * Translates the position of this address into an
     * A1-reference string.
     * 
     * @return the position of this address in A1-reference style
     * @since xelem.2.0
     */
    public String getA1Reference() {
        StringBuffer sb = new StringBuffer(calculateColumn(c));
        sb.append(r);
        return sb.toString();
    }

    /**
     * Translates the position of this address into an 
     * absolute R1C1-reference string.
     * 
     * @return The position of this address as an absolute R1C1-reference string.
     */
    public String getAbsoluteAddress() {
        StringBuffer sb = new StringBuffer("R");
        sb.append(r);
        sb.append("C");
        sb.append(c);
        return sb.toString();
    }
    
    /**
     * Gets the absolute range-address of a rectangular range 
     * in R1C1-reference style. The rectangle is
     * delimited by this address and <code>otherAddress</code>. 
     * This address can be in
     * any of the four corners of the rectangle, as long as 
     * <code>otherAddress</code> is in the opposite corner.
     * <P>
     * <img src="doc-files/getAbsoluteRange.gif">
     * <P>
     * @param otherAddress The address in the opposite corner.
     * @return A string in R1C1-reference style.
     */
    public String getAbsoluteRange(Address otherAddress) {
        return getAbsoluteRange(otherAddress.r, otherAddress.c);
    }
    
    /**
     * Gets the absolute range-address of a rectangular range 
     * in R1C1-reference style. The rectangle is
     * delimited by this address and the cell at the intersection of
     * <code>row</code> and <code>column</code>.
     * This address can be in
     * any of the four corners of the rectangle, as long as 
     * the intersection of
     * <code>row</code> and <code>column</code> is in the opposite corner.
     * <P>
     * <img src="doc-files/getAbsoluteRange_int_int.gif">
     * <P>
     * @param row		The row of the cell in the opposite corner.
     * @param column	The column of the cell in the opposite corner.
     * @return A string in R1C1-reference style.
     */
    public String getAbsoluteRange(int row, int column) {
        StringBuffer sb = new StringBuffer("R");
        int minR = r;
        int maxR = row;
        if (minR > maxR) {
            minR = row;
            maxR = r;
        }
        sb.append(minR);
        sb.append("C");
        int minC = c;
        int maxC = column;
        if (minC > maxC) {
            minC = column;
            maxC = c;
        }
        sb.append(minC);
        if (minR != maxR || minC != maxC) {
            sb.append(":R");
            sb.append(maxR);
            sb.append("C");
            sb.append(maxC);
        }
        return sb.toString();
    }
    
    /**
     * Gets the absolute range-address of a collection of addresses
     * in R1C1-reference style.
     * <P>
     * <img src="doc-files/getAbsoluteRange_Set.gif">
     * <P>
     * @param addresses A collection of addresses.
     * @return A string in R1C1-reference style or <code>null</code> if
     * 			the list is empty.
     * @throws ClassCastException If the addresses provided are not of equal class.
     */
    public static String getAbsoluteRange(Collection<Address> addresses) {
        if (addresses.size() == 0) {
            return null;
        }
        StringBuffer sb = new StringBuffer();
        Set<Address> adrs = new TreeSet<Address>(addresses);
        for (Address a : adrs) {
            if (sb.length() > 0) sb.append(","); 
            sb.append(a.getAbsoluteAddress());
        }
        return sb.toString();
    }
    
    /**
     * Gets a relative reference from this address to another address 
     * in R1C1-reference style.
     * <P>
     * <img src="doc-files/getRefTo.gif">
     * <P>
     * @param otherAddress The referenced address.
     * @return A string in R1C1-reference style.
     */
    public String getRefTo(Address otherAddress) {
        return getRefTo(otherAddress.r, otherAddress.c);
    }
    
    /**
     * Gets a relative reference from this address to a cell at the 
     * intersection of row and column in R1C1-reference style.
     * <P>
     * <img src="doc-files/getRefTo_int_int.gif">
     * <P>
     * @param row 		The row to be referenced.
     * @param column	The column to be referenced.
     * @return A string in R1C1-reference style.
     */
    public String getRefTo(int row, int column) {
        StringBuffer sb = new StringBuffer("R");
        int rO = row - r;
        if (rO != 0) {
            sb.append("[");
            sb.append(rO);
            sb.append("]");
        }
        sb.append("C");
        int cO = column - c;
        if (cO != 0) {
            sb.append("[");
            sb.append(cO);
            sb.append("]");
        }
        return sb.toString();
    }
    
    /**
     * Gets a relative reference from this address to a rectanglular range
     * in R1C1-reference style.
     * The rectangle is delimited by address1 and address2.
     * Address1 can be in any of the four corners of the rectangle, as long as
     * address2 is in the opposite corner.
     * <P>
     * <img src="doc-files/getRefTo_AdrAdr.gif">
     * <P>
     * @param address1 	The address in one corner of the range to be referenced.
     * @param address2	The address in the opposite corner of the range
     * 					to be referenced.
     * @return A string in R1C1-reference style.
     */
    public String getRefTo(Address address1, Address address2) {
        return getRefTo(address1.r, address1.c, address2.r, address2.c);
    }
    
    /**
     * Gets a relative reference from this address to a rectanglular range
     * in R1C1-reference style.
     * The rectangle is delimited by cells at the
     * intersection of r1, c1 and r2, c2 respectively.
     * The intersection of r1, c1 can be in any of the four corners of the 
     * rectangle, as long as the intersection of r2, c2
     * is in the opposite corner.
     * <P>
     * <img src="doc-files/getRefTo_4int.gif">
     * <P>
     * @param r1 	The row at intersection 1.
     * @param c1	The column at intersection 1.
     * @param r2 	The row at intersection 2.
     * @param c2	The column at intersection 2.
     * @return A string in R1C1-reference style.
     */
    public String getRefTo(int r1, int c1, int r2, int c2) {
        StringBuffer sb = new StringBuffer();
        String ref1 = getRefTo(r1, c1);
        sb.append(ref1);
        String ref2 = getRefTo(r2, c2);
        if (!ref1.equals(ref2)) {
            sb.append(":");
            sb.append(ref2);
        }
        return sb.toString();
    }
    
    /**
     * Gets a relative reference from this address to an area.
     * 
     * @param area	the area to reference
     * @return A string in R1C1-reference style.
     */
    public String getRefTo(Area area) {
        return getRefTo(area.r1, area.c1, area.r2, area.c2);
    }
    
    /**
     * Gets a relative reference from this address to a collection of
     * addresses in R1C1-reference style.
     * <P>
     * <img src="doc-files/getRefTo_Set.gif">
     * <P>
     * @param addresses	A collection of addresses.
     * @return A string in R1C1-reference style or <code>null</code> if
     * 			the list is empty.
     * @throws ClassCastException If the addresses provided are not of equal class.
     */
    public String getRefTo(Collection<Address> addresses) {
        if (addresses.size() == 0) {
            return null;
        }   
        StringBuffer sb = new StringBuffer();
        Set<Address> adrs = new TreeSet<Address>(addresses);
        for (Address a : adrs) {
            if (sb.length() > 0) sb.append(","); 
            sb.append(getRefTo(a));
        }
        return sb.toString();
    }
    
    /**
     * Returns a string representation of this address.
     * Composed as:
     * <PRE>
     *          this.getClass().getName() + "[row=x,column=y]"
     * </PRE>
     * where x and y stand for row- and column index of this address.
     * 
     * @return A string representation of this address.
     */
    public String toString() {
        StringBuffer sb = new StringBuffer(this.getClass().getName());
        sb.append("[row=");
        sb.append(r);
        sb.append(",column=");
        sb.append(c);
        sb.append("]");
        return  sb.toString();
    }
    
    /**
     * Specifies whether the object in the parameter is equal to this address.
     * Another object is equal to this address if
     * <UL>
     * <LI>their classes are equal and
     * <LI>their row indexes are equal and
     * <LI>their column indexes are equal.
     * </UL>
     * In fact this method invokes the toString-method on both the object and
     * this address and performs the equals-test on these strings.
     * 
     * @param obj An object.
     * @return <code>true</code> if this address equals <code>obj</code>,
     * 			<code>false</code> otherwise.
     */
    public boolean equals(Object obj) {
        if (obj == null) return false;
        return this.toString().equals(obj.toString());
    }
    
    /**
     * Compare this addres with the specified object for order. 
     * The specified object is cast to an address. Returns
     * <UL>
     * <LI>Negative if the row index of this address is less then the row index
     * 	of <code>o</code> or, if both have the same row index, if the column
     *  index of this address is less then the column index of of <code>o</code>;
     * <LI>Zero if both row and column index of this address and <code>o</code>
     *  are the same;
     * <LI>Positive if the row index of this address is greater then the row index
     * 	of <code>o</code> or, if both have the same row index, if the column
     *  index of this address is greater then the column index of of <code>o</code>
     * </UL>
     * 
     * @param o the object to be compared
     * @return A negative integer, zero, or a positive integer as this address is
     * 			less then, equal to or greater then the specified object.
     * @throws ClassCastException If the class of the specified object is not
     * 			equal to the class of this object.
     */
    public int compareTo(Object o) {
        if (!o.getClass().equals(this.getClass())) {
            throw new ClassCastException();
        }
        Address a = (Address) o;
        //return Worksheet.lastColumn * (r - a.r) + (c - a.c);
        if (r == a.r) {
            return c - a.c;
        }
        return r - a.r;
    }
    

}
