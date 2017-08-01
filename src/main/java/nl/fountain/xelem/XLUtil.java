/*
 * Created on Sep 8, 2004
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
package nl.fountain.xelem;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * A utility class for xelem.
 */
public class XLUtil {
       
    private static SimpleDateFormat xldf;
    
    // this class has only static methods.
    private XLUtil() {}
    
    /**
     * Converts rgb to a string. Like <code><i>#ff00ba</i></code>
     * 
     * @param 	r	the value of red (0 - 255)
     * @param 	g	the value of green (0 - 255)
     * @param 	b	the value of blue (0 - 255)
     * 
     * @return 	A String starting with <code>#</code>
     * 			+ the two-digit hex-code for 
     * 			<code>r</code>, <code>g</code> and <code>b</code>.
     */
    public static String convertToHex(int r, int g, int b) {
        StringBuffer sb = new StringBuffer("#");
        String red = Integer.toHexString(toByte(r));
        sb.append(red.length() < 2 ? "0" + red : red);
        String green = Integer.toHexString(toByte(g));
        sb.append(green.length() < 2 ? "0" + green : green);
        String blue = Integer.toHexString(toByte(b));
        sb.append(blue.length() < 2 ? "0" + blue : blue);
        return sb.toString();
    }
    
    private static int toByte(int i) {
        if (i < 0) i = 0;
        if (i > 255) i = 255;
        return i;
    }
    
    /**
     * Formats a Date into a Date-Time value used by SpreadsheetML.
     * 
     * @param 	date 	The date to be formatted.
     * 
     * @return A string of format <code>yyyy-MM-ddTHH:mm:ss.SSS</code>.
     */
    public static String format(Date date) {
        StringBuffer sb = new StringBuffer(getDateFormat().format(date));
        sb.append("T");
        sb.append(getTimeFormat().format(date));
        return sb.toString();
    }
    
    /**
     * Parses a string in the DateTime format used by SpreadsheetML 
     * to a Date.
     * 
     * @param dateString	the string to be parsed
     * @return				the corresponding Date
     * @throws				java.text.ParseException 
     * 						if the string could not be parsed
     */
    public static Date parse(String dateString) {
        String datum = dateString.substring(0, 10);
        String tijd = dateString.substring(11, 19);
        Date date = null;
        try {
            date = getExcelFormat().parse(datum + " " + tijd);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return date;
    }
    
    private static DateFormat getDateFormat() {
        if (xldf == null) {
            xldf = new SimpleDateFormat();
        }
        xldf.applyPattern("yyyy-MM-dd");
        return xldf;
    }
    
    private static DateFormat getTimeFormat() {
        if (xldf == null) {
           xldf = new SimpleDateFormat(); 
        }
        xldf.applyPattern("HH:mm:ss.SSS");
        return xldf;
    }
    
    private static DateFormat getExcelFormat() {
        if (xldf == null) {
           xldf = new SimpleDateFormat(); 
        }
        xldf.applyPattern("yyyy-MM-dd HH:mm:ss");
        return xldf;
    }

}
