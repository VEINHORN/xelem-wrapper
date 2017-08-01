/*
 * Created on 24-okt-2004
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
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.FactoryConfigurationError;
import javax.xml.parsers.ParserConfigurationException;

import nl.fountain.xelem.excel.XLElement;

import org.w3c.dom.Attr;
import org.w3c.dom.DOMException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

/**
 * An intermediary to the configuration file, it's main productline being 
 * SpreadsheetML Style elements. An instance of this class may be obtained by calling
 * the {@link #newInstance()}-method. 
 * <P>
 * Normally this class expects a configuration file at the location
 * <code>config/xelem.xml</code>, relative to the classloader of the 
 * main application. If desired a different location may be set by calling
 * {@link #setConfigurationFileName(String)} prior to obtaining a new instance.
 * <P>
 * At the first call to {@link #newInstance()} the XFactory parses the
 * configuration file and loads comments and styles from it. The XFactory
 * can be used to merge styles.
 * 
 * @see <a href="../../../overview-summary.html#xfactory">overview - the XFactory</a>
 */
public class XFactory {
    
    private static final String XMLNS_F="urn:schemas-fountain-nl:xelem:config";
    private static final String CONFIGURATION_FILE = "config/xelem.xml";
    
    private static String configFileName;
    private static boolean loaded = false;
    private static List<String> doc_comments;
    private static Map<String, Node> styles;
    private static Set<String> issStyles;
    private static Node infoSheet;
    
    private XFactory() throws XelemException {
        if (!loaded) loadConfiguration(getConfigurationFileName());
    }
    
    /**
     * Creates a new, but empty, XFactory.
     * @param empty	has no effect
     */
    private XFactory(boolean empty) {
        if (empty) {};
        reset();
        init();
    }
    
    /**
     * Gets an instance of this class.
     * 
     * @return	an instance of this class
     * @throws XelemException could be caused by:
     * <UL>
     * <LI>{@link org.w3c.dom.DOMException}
     * <LI>{@link javax.xml.parsers.FactoryConfigurationError}
     * <LI>{@link javax.xml.parsers.ParserConfigurationException}
     * <LI>{@link org.xml.sax.SAXException}
     * <LI>{@link java.io.IOException}
     * </UL>
     */
    public static XFactory newInstance() throws XelemException {
        return new XFactory();
    }
    
    /**
     * Gets a new empty instance of this class; all existing instances of
     * this class will be empty after calling this method. If a configuration
     * file name was previously set, this will no longer hold.
     * 
     * @return new empty instance of this class
     */
    public static XFactory emptyFactory() {
        return new XFactory(true);
    }
    
    /**
     * Sets path and file name of the configuration file. This method may be invoked
     * if a configuration file is used other than the one at 
     * <code>config/xelem.xml</code>. The file name must be absolute or
     * relative to the classloader of the main application.
     * 
     * @param fileName	The file name (may include path) of the configuration file.
     */
    public static void setConfigurationFileName(String fileName) {
        reset();
        configFileName = fileName;
    }
    
    /**
     * Gets the file name of the configuration file. If no configuration
     * file was set previously, returns the default configuration file name:
     * <code>config/xelem.xml</code>.
     * 
     * @return the file name (may include path) of the configuration file
     */
    public static String getConfigurationFileName() {
        if (configFileName == null) {
            return CONFIGURATION_FILE;
        } else {
            return configFileName;
        }
    }
    
    /**
     * Resets the static configuration and configuration file name.
     * Changes in the configuration file
     * <code>config/xelem.xml</code> will be reflected in a new instance.
     */
    public static void reset() {
        loaded = false;
        configFileName = null;
    }
    
    /**
     * Gets a list of Strings which represent the node values of the tag
     * <code>&lt;f:comment&gt;</code> in the configuration file. 
     * If no comments are present the list may be empty.
     * 
     * @return	a list of Strings
     */
    public List<String> getDocComments() {
        return doc_comments;
    }
    
    /**
     * Gets the SpreadsheetML Style element with the ss:StyleID <code>id</code>.
     * The Style element was either present in the configuration file or was
     * previously merged with the method 
     * {@link #mergeStyles(String newID, String id1, String id2) mergeStyles}.
     * 
     * @param id	The ss:ID attribute-value of the wanted style.
     * 
     * @return 	A Style element. May be <code>null</code> if the wanted style
     * 			is unknown to the factory.
     */
    public Element getStyle(String id) {
        return (Element) styles.get(id);
    }
    
    /**
     * Gets the number of available styles in the factory.
     * 
     * @return 	the number of available styles.
     */
    public int getStylesCount() {
        return styles.size();
    }
    
    /**
     * Gets a set of ss:ID's (String) of all the available styles in the factory.
     * 
     * @return 	all available style id's.
     */
    public Set<String> getStyleIDs() {
        return styles.keySet();
    }
    
    /**
     * Merges two SpreadsheetML Style elements. If a style element with the
     * ss:ID <code>newID</code> allready was present in the factory, nothing happens.
     * <br>
     * Creates a copy of style 1
     * and appends those child elements of style 2 that were previously not
     * defined in style 1. Any ss:Name attribute is removed from the copy.
     * The copy is given the ss:ID <code>newID</code>
     * and is from now on availlable by the method 
     * {@link #getStyle(String id) getStyle(newID)}.
     * 
     * @param 	newID	the ss:ID of the new merged style
     * @param	id1		the ss:ID of the copied style
     * @param 	id2		the ss:ID of the appended style
     * 
     * @throws 	UnsupportedStyleException if either of the styles mentioned with 
     * 			id1 or id2 are unknown to the factory.
     */
    public void mergeStyles(String newID, String id1, String id2) throws UnsupportedStyleException {
        Element ne = getStyle(newID);
        if (ne != null) return; // we allready have a style with this ID
        
        Element s1 = getStyle(id1);
        Element s2 = getStyle(id2);
        String notFoundId = "";
        if (s1 == null) notFoundId = " '" + id1 + "'";
        if (s2 == null) notFoundId += " '" + id2 + "'";
        if (s1 == null || s2 == null) {
            throw new UnsupportedStyleException(
                    "Style(s)" + notFoundId + " not found.");
        }
        
        ne = (Element) s1.cloneNode(true);
        ne.setAttributeNS(XLElement.XMLNS_SS, "ID", newID);
        Attr id = s1.getOwnerDocument().createAttributeNS(XLElement.XMLNS_SS, "ID");
        id.setPrefix(XLElement.PREFIX_SS);
        id.setValue(newID);
        ne.setAttributeNodeNS(id);
        ne.removeAttributeNS(XLElement.XMLNS_SS, "Name");
        
        Element clonedS2 = (Element) s2.cloneNode(true);
        
        String[] elements = new String[] { "Alignment", "Borders",
                "Font",  "Interior", "NumberFormat", "Protection" };
        for (String s : elements) {
	        NodeList nl2 = clonedS2.getElementsByTagName(s);
	        if (nl2.getLength() > 0) {
	            NodeList nl1 = ne.getElementsByTagName(s);
	            if (nl1.getLength() == 0) {
	                ne.appendChild(nl2.item(0));
	            }
	        }
        }
        
        styles.put(newID, ne);
    }
    
    /**
     * Add a new SpreadsheetML Style element to the factory. If the
     * factory allready contains an element with the same ss:ID as the
     * passed element, no element will be added and the method returns
     * <code>false</code>.
     * 
     * @param style	a SpreadsheetML style element
     * @return 	<code>true</code> if the style element was succesfully added,
     * 	<code>false</code> otherwise
     * @throws 	java.lang.NullPointerException if the passed style did not
     * 	have an attribute ss:ID
     */
    public boolean addStyle(Element style) {
        String id = style.getAttributes().getNamedItemNS(
                XLElement.XMLNS_SS, "ID").getNodeValue();
        if (styles.containsKey(id)) {
            return false;
        }
        styles.put(id, style);
        return true;
    }
    
    /**
     * Appends a Worksheet element with general information to the root element. 
     * Adds all used styles in the info sheet to the factory 
     * and their id's to the gio.
     * 
     * @param 	root 	the root element.
     * @param 	gio		the GIO used while assembling the Workbook.
     * 
     * @throws 	XelemException	if the info sheet could not be loaded.
     * 
     * @see 	nl.fountain.xelem.excel.Workbook#appendInfoSheet()
     */
    public void appendInfoSheet(Element root, GIO gio) throws XelemException {
        if (infoSheet == null) {
            infoSheet = loadInfoSheet();
        }
        root.appendChild(root.getOwnerDocument().importNode(infoSheet, true));
        gio.getStyleIDSet().addAll(issStyles);
    }
    
    private void init() {
        doc_comments = new ArrayList<String>();
        styles = new HashMap<String, Node>();
    }
    
    private void loadConfiguration(String fileName) throws XelemException {
        init();
        
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            File configFile = new File(fileName);
            Document config = builder.parse(configFile);
            
            NodeList comments = config.getElementsByTagNameNS(XMLNS_F, "comment");
            for (int i = 0; i < comments.getLength(); i++) {
                String doc_comment = comments.item(i).getFirstChild().getNodeValue();
                doc_comments.add(doc_comment);
            }
            
            NodeList styleList = config.getElementsByTagName("Style");
            for (int i = 0; i < styleList.getLength(); i++) {
                Node style = styleList.item(i);
                String id = style.getAttributes().getNamedItemNS(
                        XLElement.XMLNS_SS, "ID").getNodeValue();
                styles.put(id, style);
            }
            
        } catch (DOMException e) {
            throw new XelemException(e.fillInStackTrace());
        } catch (FactoryConfigurationError e) {
            throw new XelemException(e.fillInStackTrace());
        } catch (ParserConfigurationException e) {
            throw new XelemException(e.fillInStackTrace());
        } catch (SAXException e) {
            throw new XelemException(e.fillInStackTrace());
        } catch (IOException e) {
            throw new XelemException(e.fillInStackTrace());
        }
        loaded = true;
    }

    Node loadInfoSheet() throws XelemException {
        issStyles = new TreeSet<String>();
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            InputStream is = this.getClass().getResourceAsStream("infoSheet.xml");
            Document infosh = builder.parse(is);
            
            NodeList styleList = infosh.getElementsByTagName("Style");
            for (int i = 0; i < styleList.getLength(); i++) {
                Node style = styleList.item(i);
                String id = style.getAttributes().getNamedItemNS(
                        XLElement.XMLNS_SS, "ID").getNodeValue();
                if (!"Default".equals(id)) {
                    issStyles.add(id);
                    styles.put(id, style);
                }
            }
            NodeList sheets = 
                infosh.getElementsByTagNameNS(XLElement.XMLNS_SS, "Worksheet");
            infoSheet = sheets.item(0);
            
        } catch (ParserConfigurationException e) {
            throw new XelemException(e.fillInStackTrace());
        } catch (SAXException e) {
            throw new XelemException(e.fillInStackTrace());
        } catch (IOException e) {
            throw new XelemException(e.fillInStackTrace());
        }
        return infoSheet;
    }


}
