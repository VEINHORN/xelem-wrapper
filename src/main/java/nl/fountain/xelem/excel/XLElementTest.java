/*
 * Created on Nov 2, 2004
 *
 */
package nl.fountain.xelem.excel;

import java.io.StringWriter;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.TransformerFactoryConfigurationError;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import junit.framework.TestCase;

import nl.fountain.xelem.GIO;

import org.w3c.dom.DOMImplementation;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

/**
 *
 */
public abstract class XLElementTest extends TestCase {
    
    protected Document getDoc() throws ParserConfigurationException {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        DocumentBuilder builder = factory.newDocumentBuilder();
        DOMImplementation domImpl = builder.getDOMImplementation();
        Document doc = domImpl.createDocument(XLElement.XMLNS, "Test", null);
        Element root = doc.getDocumentElement();
        root.setAttribute("xmlns", XLElement.XMLNS);
        root.setAttribute("xmlns:o", XLElement.XMLNS_O);
        root.setAttribute("xmlns:x", XLElement.XMLNS_X);
        root.setAttribute("xmlns:ss", XLElement.XMLNS_SS);
        root.setAttribute("xmlns:html", XLElement.XMLNS_HTML);
        return doc;
    }
    
    private void transform(Document doc, Result result) throws TransformerFactoryConfigurationError, TransformerConfigurationException, TransformerException {       
        TransformerFactory tFactory = TransformerFactory.newInstance();
        Transformer xformer = tFactory.newTransformer();
        xformer.setOutputProperty(OutputKeys.METHOD, "xml");
        xformer.setOutputProperty(OutputKeys.INDENT, "yes");
        Source source = new DOMSource(doc);
        xformer.transform(source, result);
    }
    
    private String xmlToString(Document doc) throws TransformerConfigurationException, TransformerFactoryConfigurationError, TransformerException {
        StringWriter sw = new StringWriter();
        Result result = new StreamResult(sw);
        transform(doc, result);
        return sw.toString();
    }
    
    protected String xmlToString(XLElement xle, GIO gio) {
        String xml = null;
        try {
            Document doc = getDoc();
            Element parent = doc.getDocumentElement();
            xle.assemble(parent, gio);
            xml = xmlToString(doc);
        } catch (TransformerConfigurationException e) {
            e.printStackTrace();
        } catch (ParserConfigurationException e) {
            e.printStackTrace();
        } catch (TransformerFactoryConfigurationError e) {
            e.printStackTrace();
        } catch (TransformerException e) {
            e.printStackTrace();
        } catch (SecurityException e) {
            e.printStackTrace();
        } catch (IllegalArgumentException e) {
            e.printStackTrace();
        }
        return xml;
    }
    
    protected String xmlToString(Node n) {
        String xml = null;
        try {
            Document doc = getDoc();
            doc.getDocumentElement().appendChild(doc.importNode(n, true));
            xml = xmlToString(doc);
        } catch (ParserConfigurationException e) {
            e.printStackTrace();
        } catch (TransformerConfigurationException e) {
            e.printStackTrace();
        } catch (TransformerFactoryConfigurationError e) {
            e.printStackTrace();
        } catch (TransformerException e) {
            e.printStackTrace();
        }
        return xml;
    }
    

}
