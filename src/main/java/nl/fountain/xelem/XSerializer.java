/*
 * Created on 9-nov-2004
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

import java.io.File;
import java.io.OutputStream;
import java.io.StringWriter;
import java.io.Writer;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import nl.fountain.xelem.excel.Workbook;

import org.w3c.dom.Document;

/**
 * A conveniance class for serializing Workbooks.
 * 
 */
public class XSerializer {
    
    private Transformer xformer;
    private String encoding;
    
    public static final String US_ASCII = "US-ASCII";
    
    public XSerializer() {}
    
    public XSerializer(String encoding) {
        this.encoding = encoding;
    }
    
    /**
     * 
     */
    public String serializeToString(Workbook wb) throws XelemException {
        StringWriter out = new StringWriter();
        serialize(wb, out);
        return out.toString();
    }
    
    public String serializeToString(Document doc) throws XelemException {
        StringWriter out = new StringWriter();
        serialize(doc, out);
        return out.toString();
    }
    
    /**
     * Serializes the Workbook to the file specified with the Workbook's
     * {@link nl.fountain.xelem.excel.Workbook#getFileName() getFileName}-method.
     * 
     */
    public void serialize(Workbook wb) throws XelemException {
        File out = new File(wb.getFileName());
        serialize(wb, out);
    }
    
    public void serialize(Workbook wb, File out) throws XelemException {
        Result result = new StreamResult(out);
        transform(wb, result); 
    }
    
    public void serialize(Document doc, File out) throws XelemException {
        Result result = new StreamResult(out);
        transform(doc, result); 
    }
    
    public void serialize(Workbook wb, OutputStream out) throws XelemException {
        Result result = new StreamResult(out);
        transform(wb, result);
    }
    
    public void serialize(Document doc, OutputStream out) throws XelemException {
        Result result = new StreamResult(out);
        transform(doc, result);
    }
    
    public void serialize(Workbook wb, Writer out) throws XelemException {
        Result result = new StreamResult(out);
        transform(wb, result);
    }
    
    public void serialize(Document doc, Writer out) throws XelemException {
        Result result = new StreamResult(out);
        transform(doc, result);
    }
    
    private void transform(Workbook wb, Result result) throws XelemException {
        try {
            Document doc = wb.createDocument();
            transform(doc, result);
        } catch (ParserConfigurationException e) {
            throw new XelemException(e.fillInStackTrace());
        }
    }
    
    private void transform(Document doc, Result result) throws XelemException {
        try {
            Transformer xformer = getTransformer();          
            Source source = new DOMSource(doc);
            xformer.transform(source, result);
        } catch (TransformerException e) {
            throw new XelemException(e.fillInStackTrace());
        }
    }

    private Transformer getTransformer() throws XelemException {
        if (xformer == null) {
	        TransformerFactory tFactory = TransformerFactory.newInstance();
	        try {
	            xformer = tFactory.newTransformer();
	        } catch (TransformerConfigurationException e) {
	            throw new XelemException(e.fillInStackTrace());
	        }
	        xformer.setOutputProperty(OutputKeys.METHOD, "xml");
	        xformer.setOutputProperty(OutputKeys.INDENT, "yes");
	        if (encoding != null) {
	            xformer.setOutputProperty(OutputKeys.ENCODING, encoding);
	        }
        }
        return xformer;
    }

}
