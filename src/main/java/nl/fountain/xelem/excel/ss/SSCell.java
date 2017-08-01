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
package nl.fountain.xelem.excel.ss;

import java.lang.reflect.Method;
import java.util.Date;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.XLUtil;
import nl.fountain.xelem.excel.AbstractXLElement;
import nl.fountain.xelem.excel.Cell;
import nl.fountain.xelem.excel.Comment;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.xml.sax.Attributes;

/**
 * An implementation of the XLElement Cell.
 * 
 * <P>
 * <h3>The setData-methods</h3>
 * The overloaded setData-methods will set the data displayed in the cell
 * when the xml spreadsheet is opened. These methods will set the Excel
 * datatype according to the java-type of the passed parameter. 
 * <P>
 * The {@link #setData(Object)}-method reflects
 * upon the class of the given object and will delegate to a corresponding
 * setData-method if such a method is available. If no corresponding
 * method is found this method sets the data of this cell to the
 * <code>toString</code>-value of the given object and the datatype to
 * "String".
 * This class may be extended to accommodate setData-methods for java-objects 
 * that might otherwise
 * be displayed as the <code>toString</code>-value of that object.
 * 
 * <P id="nullvalues">
 * <b>Null values</b><br>
 * If the passed parameter has a value of <code>null</code> the resulting
 * xml will have a datatype set to "Error",
 * the formula of the cell will be set to "<code>=#N/A</code>" and the cell will 
 * display "<code>#N/A</code>" when the spreadsheet is opened.
 * 
 * <P id="infinitevalues">
 * <b>Infinite values</b><br>
 * If the passed parameter is of type Double, Float or the primitive representation 
 * of these objects and the method {@link java.lang.Double#isInfinite() isInfinite}
 * results to <code>true</code> the resulting xml will have a datatype set to
 * "String" and the cell will display "Infinite" when the spreadsheet is opened.
 * 
 * <P id="nanvalues">
 * <b>NaN values</b><br>
 * If the passed parameter is of type Double, Float or the primitive representation 
 * of these objects and the method {@link java.lang.Double#isNaN() isNaN}
 * results to <code>true</code> the resulting xml will have a datatype set to
 * "String" and the cell will display "NaN" when the spreadsheet is opened.
 * 
 * @see nl.fountain.xelem.excel.Worksheet#addCell()
 * @see nl.fountain.xelem.excel.Row#addCell()
 */
public class SSCell extends AbstractXLElement implements Cell {
    
    private int idx;
    private boolean hasdata;
    private String styleID;
    private String formula;
    private String href;
    private String data$;
    private String datatype;
    private int mergeacross;
    private int mergedown;
    private Comment comment;
    
    /**
     * Creates a new SSCell with an initial datatype of "String" and an
     * empty ("") value.
     * 
     * @see nl.fountain.xelem.excel.Worksheet#addCell()
     */
    public SSCell() {
        datatype = DATATYPE_STRING;
        data$ = "";
    }
    
    /**
     * Sets the ss:StyleID on this cell. If no styleID is set on a cell,
     * the ss:StyleID-attribute is not deployed in the resulting xml and
     * the Default-style will be employed on the cell.
     * <P>
     * The refered style-definition must be available from the 
     * {@link nl.fountain.xelem.XFactory}.
     * If the definition was not found,
     * the {@link XLWorkbook}-implementation will create an empty ss:Style-definition
     * and adds a UnsupportedStyleException-warning.
     * 
     * @param 	id	the id of the style to employ on this cell.
     * 
     * @see XLWorkbook#getWarnings()
     */
    public void setStyleID(String id) {
        styleID = id;
    }
    
    public String getStyleID() {
        return styleID;
    }
    
    public void setFormula(String formula) {
        //this.formula = XLUtil.escapeHTML(formula);
        this.formula = formula;
    }

    public String getFormula() {
        return formula;
    }
    
    public void setHRef(String href) {
        this.href = href;
    }

    public String getHRef() {
        return href;
    }
    
    public Comment addComment() {
        comment = new SSComment();
        return comment;
    }
    
    public Comment addComment(Comment comment) {
        this.comment = comment;
        return comment;
    }
    
    public Comment addComment(String text) {
        comment = new SSComment();
        comment.setData(text);
        return comment;
    }
    
    public boolean hasComment() {
        return comment != null;
    }
    
    public Comment getComment() {
        return comment;
    }
    
    public void setMergeAcross(int m) {
        mergeacross = m;
    }
    
    private void setMergeAcross(String s) {
        mergeacross = Integer.parseInt(s);
    }
    
    public int getMergeAcross() {
        return mergeacross;
    }
    
    public void setMergeDown(int m) {
        mergedown = m;
    }
    
    private void setMergeDown(String s) {
        mergedown = Integer.parseInt(s);
    }
    
    public int getMergeDown() {
        return mergedown;
    }

    public String getXLDataType() { 
        return datatype;
    }

    public void setData(Number data) {
        if (data == null) {
            setError(ERRORVALUE_NA);
            return;
        }
        setData(data.doubleValue());
    }
    
    public void setData(Integer data) {
        if (data == null) {
            setError(ERRORVALUE_NA);
            return;
        }
        datatype = DATATYPE_NUMBER;
        setData$(data.toString());
    }
    
    public void setData(Double data) {
        if (data == null) {
            setError(ERRORVALUE_NA);
            return;
        }
        if (data.isInfinite() || data.isNaN()) {
            datatype = DATATYPE_STRING;
        } else {
            datatype = DATATYPE_NUMBER;
        }
        setData$(data.toString());
    }
    
    public void setData(Long data) {
        if (data == null) {
            setError(ERRORVALUE_NA);
            return;
        }
        datatype = DATATYPE_NUMBER;
        setData$(data.toString());
    }
    
    public void setData(Float data) {
        if (data == null) {
            setError(ERRORVALUE_NA);
            return;
        }
        if (data.isInfinite() || data.isNaN()) {
            datatype = DATATYPE_STRING;
        } else {
            datatype = DATATYPE_NUMBER;
        }
        setData$(data.toString());
    }

    public void setData(Date data) {
        if (data == null) {
            setError(ERRORVALUE_NA);
            return;
        }
        datatype = DATATYPE_DATE_TIME;
        setData$(XLUtil.format(data));
    }

    public void setData(Boolean data) {
        if (data == null) {
            setError(ERRORVALUE_NA);
            return;
        }
        setData(data.booleanValue());
    }

    public void setData(String data) {
        if (data == null) {
            setError(ERRORVALUE_NA);
            return;
        }
        datatype = DATATYPE_STRING;
        setData$(data);
    }
    
    public void setData(Object data) {
        if (data == null) {
            setError(ERRORVALUE_NA);
            return;
        }
        Class[] types = new Class[] {data.getClass()};
        Method method = null;
        try {
            method = this.getClass().getMethod("setData", types);
            method.invoke(this, new Object[]{data});
        } catch (NoSuchMethodException e) {
            setData(data.toString());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void setError(String error_value) {
        datatype = DATATYPE_ERROR;
        setData$(error_value);
        setFormula("=" + error_value);
    }
    
    public void setData(byte data) {
        datatype = DATATYPE_NUMBER;
        setData$(String.valueOf(data));
    }
    
    public void setData(short data) {
        datatype = DATATYPE_NUMBER;
        setData$(String.valueOf(data));
    }

    public void setData(int data) {
        datatype = DATATYPE_NUMBER;
        setData$(String.valueOf(data));
    }

    public void setData(long data) {
        datatype = DATATYPE_NUMBER;
        setData$(String.valueOf(data));
    }
    
    public void setData(float data) {
        if (Float.isInfinite(data) || Float.isNaN(data)) {
            datatype = DATATYPE_STRING;
        } else {
            datatype = DATATYPE_NUMBER;
        }
        setData$(String.valueOf(data));
    }

    public void setData(double data) {
        if (Double.isInfinite(data) || Double.isNaN(data)) {
            datatype = DATATYPE_STRING;
        } else {
            datatype = DATATYPE_NUMBER;
        }
        setData$(String.valueOf(data));       
    }
    
    public void setData(char data) {
        setData(String.valueOf(data));
    }
    
    public void setData(boolean data) {
        datatype = DATATYPE_BOOLEAN;
        if (data) {
            setData$("1");
        } else {
            setData$("0");
        }
    }
    
    private void setData$(String s) {
        data$ = s;
        hasdata = true;
        //System.out.println(data$);
    }
    
    public boolean hasData() {
        return hasdata;
    }
    
    public boolean hasError() {
        return DATATYPE_ERROR.equals(getXLDataType());
    }

    public String getData$() {
        return data$;
    }
    
    public Object getData() {
        if (DATATYPE_NUMBER.equals(datatype)) {
            return new Double(data$);
        } else if (DATATYPE_DATE_TIME.equals(datatype)) {
            return XLUtil.parse(data$);
        } else if (DATATYPE_BOOLEAN.equals(datatype)) {
            return new Boolean("1".equals(data$));
        }
        return data$;
    }
    
    public int intValue() {
        try {
            return new Double(data$).intValue();
        } catch (NumberFormatException e) {
            return 0;
        }
    }
    
    public double doubleValue() {
        try {
            return new Double(data$).doubleValue();
        } catch (NumberFormatException e) {
            return 0.0D;
        }
    }
    
    public boolean booleanValue() {
        return "1".equals(data$);
    }

    public String getTagName() {
        return "Cell";
    }
    
    public String getNameSpace() {
        return XMLNS_SS;
    }
    
    public String getPrefix() {
        return PREFIX_SS;
    }

    public Element assemble(Element parent, GIO gio) {
        Document doc = parent.getOwnerDocument();
        Element ce = assemble(doc, gio);
        
        if (idx != 0) ce.setAttributeNodeNS(
                createAttributeNS(doc, "Index", idx));
        if (getStyleID() != null) {
            ce.setAttributeNodeNS(createAttributeNS(doc, "StyleID", getStyleID()));
            gio.addStyleID(getStyleID());
        }
        if (formula != null) ce.setAttributeNodeNS(
                createAttributeNS(doc, "Formula", formula));
        if (href != null) ce.setAttributeNodeNS(
                createAttributeNS(doc, "HRef", href));
        if (mergeacross > 0) ce.setAttributeNodeNS(
                createAttributeNS(doc, "MergeAcross", mergeacross));
        if (mergedown > 0) ce.setAttributeNodeNS(
                createAttributeNS(doc, "MergeDown", mergedown));
        
        parent.appendChild(ce);
        
        if (!"".equals(getData$())) {
            Element data = getDataElement(doc);
            ce.appendChild(data);
        }
        
        if (comment != null) {
            comment.assemble(ce, gio);
        }
        
        return ce;
    }
    
    public void setAttributes(Attributes attrs) {
        for (int i = 0; i < attrs.getLength(); i++) {
            invokeMethod(attrs.getLocalName(i), attrs.getValue(i));
        }
    }
    
    public void setChildElement(String localName, String content) {
        if ("Data".equals(localName)) {
            setData$(content);
        }
    }
    
	private void invokeMethod(String name, Object value) {
        Class[] types = new Class[] { value.getClass() };
        Method method = null;
        try {
            method = this.getClass().getDeclaredMethod("set" + name, types);
            method.invoke(this, new Object[] { value });
        } catch (NoSuchMethodException e) {
            // no big deal
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    public Element getDataElement(Document doc) {
        Element data = doc.createElement("Data");
        data.setAttributeNodeNS(
                createAttributeNS(doc, "Type", getXLDataType()));
        data.appendChild(doc.createTextNode(getData$()));
        return data;
    }
    
    /**
     * @deprecated	as of xelem.2.0 use {@link #setType(String)}
     * @param type	one of Cell's DATATYPE_XXX values
     */
    protected void setXLDataType(String type) {
        setType(type);
    }

    /**
     * Sets the value of the ss:Type-attribute of the Data-element.
     * 
     * @param type	Must be one of Cell's DATATYPE_XXX values.
     */
    protected void setType(String type) {
        if (DATATYPE_BOOLEAN.equals(type)
                || DATATYPE_DATE_TIME.equals(type)
                || DATATYPE_ERROR.equals(type)
                || DATATYPE_NUMBER.equals(type)
                || DATATYPE_STRING.equals(type)) {
            datatype = type;
        } else {
            throw new IllegalArgumentException(type + " is not a valid datatype.");
        }
    }
    
    public void setIndex(int index) {
        idx = index;
    }
    
    public int getIndex() {
        return idx;
    }

}
