/*
 * Created on 15-mrt-2005
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

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import nl.fountain.xelem.Area;
import nl.fountain.xelem.excel.Workbook;
import nl.fountain.xelem.excel.XLElement;

import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.Locator;
import org.xml.sax.SAXException;
import org.xml.sax.SAXParseException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

/**
 * Basic class for reading xml-spreadsheets of type spreadsheetML. 
 * <P>
 * An ExcelReader
 * can deliver the contents of an xml-file or an xml-InputSource as a
 * fully populated {@link nl.fountain.xelem.excel.Workbook}.
 * <P>
 * Furthermore, it can dispatch events, values and instances of 
 * {@link nl.fountain.xelem.excel.XLElement XLElements} to listeners
 * registered at this ExcelReader while parsing an xml-file or
 * xml-InputSource.
 * 
 * @see <a href="package-summary.html#package_description">package overview</a>
 * 
 * @since xelem.2.0
 */
public class ExcelReader {
    
    private SAXParser parser;
    XMLReader reader;
    
    Director director;
    private Handler handler;
    private Map<String, String> uris;
    
    /**
     * Constructs a new ExcelReader.
     * Obtains a {@link javax.xml.parsers.SAXParser} from an instance of
     * {@link javax.xml.parsers.SAXParserFactory} to do the parsing.
     * There is no need to configure factory parameters.
     * 
     * @throws ParserConfigurationException if a parser cannot be created
     * 		 which satisfies the current configuration
     * @throws SAXException	for SAX errors
     */
    public ExcelReader() throws ParserConfigurationException, SAXException {
        SAXParserFactory spf = SAXParserFactory.newInstance();
        spf.setNamespaceAware(true);
        parser = spf.newSAXParser();
        director = new Director();
    }
    
    /**
     * Constructs a new ExcelReader. Uses the given parser. The parser
     * should be namespace aware.
     * 
     * @param parser a SAXParser that is namespace aware
     * @throws ParserConfigurationException	if the parser is not namespace aware
     */
    public ExcelReader(SAXParser parser) throws ParserConfigurationException {
        if (!parser.isNamespaceAware()) {
            throw new ParserConfigurationException(
                "cannot read with a parser that is unaware of namespaces.");
        }
        this.parser = parser;
        director = new Director();
    }
    
    /**
     * Returns the SAXParser that is used by this ExcelReader. (Just to see
     * what we've got under the hood.)
     * <P>
     * Under Java 1.4 we might see a org.apache.crimson.jaxp.SAXParserImpl,
     * it could be a com.sun.org.apache.xerces.internal.jaxp.SAXParserImpl
     * under Java 1.5. 
     * 
     * @return	the SAXParser that is used by this ExcelReader
     */
    public SAXParser getSaxParser() {
        return parser;
    }
    
    /**
     * Sets the area that is to be read. 
     * Reading will be restricted to the specified area until another
     * area is set or until the read area has been cleared.
     * 
     * @param area the area to read
     * @see <a href="package-summary.html#areas">package overview</a>
     */
    public void setReadArea(Area area) {
        director.setBuildArea(area);
    }
    
    /**
     * Clears the read area. Reading will not be rerstricted after a call
     * to this method.
     *
     */
    public void clearReadArea() {
        director.setBuildArea(null);
    }
    
    /**
     * Gets the area that restricts reading on this ExcelReader. May be null if
     * no read area was set or if the read area was cleared.
     * 
     * @return  the read area
     */
    public Area getReadArea() {
        if (director.hasBuildArea()) {
            return director.getBuildArea();
        } else {
            return null;
        }
    }
    
    /**
     * Specifies whether reading is restricted on this ExcelReader.
     * 
     */
    public boolean hasReadArea() {
        return director.hasBuildArea();
    }
    
    /**
     * Gets a list of registered listeners on this ExcelReader.
     * 
     * @return a list of registered listeners
     */
    public List<ExcelReaderListener> getListeners() {
        return director.getListeners();
    }
    
    /**
     * Registers the given listener on this ExcelReader.
     * 
     * @param listener the ExcelReaderListener to be registered
     * @see <a href="package-summary.html#eventbasedmodel">package overview</a>
     */
    public void addExcelReaderListener(ExcelReaderListener listener) {
        director.addExcelReaderListener(listener);
    }
    
    /**
     * Removes the passed listener on this ExcelReader.
     * 
     * @param listener the ExcelReaderListener to be removed
     * @return <code>true</code> if the passed listener was registered,
     * 		<code>false</code> otherwise
     */
    public boolean removeExcelReaderListener(ExcelReaderListener listener) {
        return director.removeExcelReaderListener(listener);
    }
    
    /**
     * Remove all listeners on this ExcelReader
     *
     */
    public void clearExcelReaderListeners() {
        director.clearExcelReaderListeners();
    }
    
    /**
     * Delivers the contents of the specified file as a fully populated Workbook.
     * If a read area was set on this ExcelReader its worksheets
     * are only populated in the specified area. If listeners are registered,
     * this ExcelReader dispatches events to these listeners during the read.
     * Performs a read. 
     * 
     * @param fileName 		the name of the file to be read
     * @return				a fully populated Workbook
     * @throws IOException 	signals a failed or interrupted I/O operation
     * @throws SAXException	signals a general SAX error or warning
     * 
     * @see #read(String)
     */
    public Workbook getWorkbook(String fileName) throws IOException, SAXException {
        InputSource in = new InputSource(fileName);
        return getWorkbook(in);
    }
    
    /**
     * Delivers the contents of the given InputSource as a fully populated Workbook.
     * If a read area was set on this ExcelReader its worksheets
     * are only populated in the specified area. If listeners are registered,
     * this ExcelReader dispatches events to these listeners during the read.
     * Performs a read. 
     * 
     * @param source		the Inputsource streaming spreadsheetML	
     * @return				a fully populated Workbook
     * @throws IOException	signals a failed or interrupted I/O operation
     * @throws SAXException	signals a general SAX error or warning
     * 
     * @see	#read(InputSource)
     */
    public Workbook getWorkbook(InputSource source) throws IOException, SAXException {
        WorkbookListener wbl = new WorkbookListener();
        addExcelReaderListener(wbl);
        read(source);
        removeExcelReaderListener(wbl);
        return wbl.getWorkbook();
    }
    
    /**
     * Reads the file with the given name and dispatches events to registered
     * {@link ExcelReaderListener ExcelReaderListeners}. If a read area was set
     * on this ExcelReader, only events occuring within the limits of the
     * area are dispatched.
     * <P>
     * When trying to read non-xml files, under java 1.5 a <code>
     * com.sun.org.apache.xerces.internal.impl.io.MalformedByteSequenceException
     * </code> might be thrown. The MalformedByteSequenceException is unknown
     * under java 1.4 and previous releases. Releases prior to java 1.5
     * will throw a <code>org.xml.sax.SAXParseException</code> under these
     * adverse conditions.
     * 
     * @param fileName		the name of the file to be read
     * @throws IOException	signals a failed or interrupted I/O operation
     * @throws SAXException	signals a general SAX error or warning
     */
    public void read(String fileName) throws IOException, SAXException {
        InputSource in = new InputSource(fileName);
        read(in);
    }
    
    /**
     * Reads the stream of the given InputSource and dispatches events to registered
     * {@link ExcelReaderListener ExcelReaderListeners}. If a read area was set
     * on this ExcelReader, only events occuring within the limits of the
     * area are dispatched.
     * <P>
     * When trying to read non-xml streams, under java 1.5 a <code>
     * com.sun.org.apache.xerces.internal.impl.io.MalformedByteSequenceException
     * </code> might be thrown. The MalformedByteSequenceException is unknown
     * under java 1.4 and previous releases. Releases prior to java 1.5
     * will throw a <code>org.xml.sax.SAXParseException</code> under these
     * adverse conditions.
     * 
     * @param source		the Inputsource streaming spreadsheetML
     * @throws IOException	signals a failed or interrupted I/O operation
     * @throws SAXException	signals a general SAX error or warning
     */
    public void read(InputSource source) throws IOException, SAXException {
        getPrefixMap().clear();
        reader = parser.getXMLReader();
        reader.setContentHandler(getHandler());
        reader.setErrorHandler(getHandler());
        reader.parse(source);
    }

    /**
     * Gets a map of prefixes (keys) and uri's recieved while reading. 
     * May be obtained after a read. Performing
     * a new read will clear any previously recorded entries from the map.
     * 
     * @return a map of prefixes (keys) and uri's recieved while reading
     */
    public Map<String, String> getPrefixMap() {
        if (uris == null) {
            uris = new HashMap<String, String>();
        }
        return uris;
    }
    
    private Handler getHandler() {
        if (handler == null) {
            handler = new Handler();
        }
        return handler;
    }
      
    
    //////////////////////////////////////////////////////////////////////////////
    
    private class Handler extends DefaultHandler {
        
        private Locator locator;
        
        public void setDocumentLocator(Locator locator) {
            this.locator = locator;
        }
        
        public void processingInstruction(String target, String data) throws SAXException {
        	for (ExcelReaderListener listener : director.getListeners()) {
                listener.processingInstruction(target, data);
            }
        }
        
        public void startPrefixMapping(String prefix, String uri) throws SAXException {
            getPrefixMap().put(prefix, uri);
        }
        
        public void startElement(String uri, String localName, String qName,
                Attributes attributes) throws SAXException {
            if (XLElement.XMLNS_SS.equals(uri) && "Workbook".equals(localName)) {
                String systemId = getSystemId();
                Builder builder = director.getXLWorkbookBuilder(); 
                for (ExcelReaderListener listener : director.getListeners()) {
                    listener.startWorkbook(systemId);
                }               
                builder.build(reader, this);               
            }       
        }
        
        public void startDocument() throws SAXException {        
        	for (ExcelReaderListener listener : director.getListeners()) {
                listener.startDocument();
            }
        }
        
        public void endDocument() throws SAXException {
        	for (ExcelReaderListener listener : director.getListeners()) {
                listener.endDocument(getPrefixMap());
            }
        }
        
        public void fatalError(SAXParseException e) throws SAXException {
            //System.out.println("fatal error detected: " + e.getMessage());
            throw e;
        }
        
        public void error(SAXParseException e) throws SAXException {
            //System.out.println("error detected: " + e.getMessage());
            throw e;
        }
        
        public void warning(SAXParseException e) throws SAXException {
            //System.out.println("warning detected: " + e.getMessage());
            throw e;
        }
        
        
        
        private String getSystemId() {
            String systemId = null;
            if (locator != null) {
                systemId = locator.getSystemId();
            }
            if (systemId == null) {
                systemId = "source";
            }
            return systemId;
        }
        

        

    }
    

}
