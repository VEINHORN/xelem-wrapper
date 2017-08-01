package nl.fountain.xelem.ztest;

import javax.xml.parsers.ParserConfigurationException;

import nl.fountain.xelem.XFactory;
import nl.fountain.xelem.XSerializer;
import nl.fountain.xelem.XelemException;
import nl.fountain.xelem.excel.Cell;
import nl.fountain.xelem.excel.XLElement;
import nl.fountain.xelem.excel.ss.XLWorkbook;

import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class StyleFactory {
	
	private static Document document;

	// main-method is just for showing..
	public static void main(String[] args) throws ParserConfigurationException, XelemException {
		StyleFactoryWorkbook wb = new StyleFactoryWorkbook("styleFactoryTest");
		Cell cell = wb.addSheet().addCell("tekst should be ");
		cell.setStyleID("TYOtest");
		
		XFactory xfactory = wb.getFactory();
		
		System.out.println("there are " + xfactory.getStylesCount() + " styles in the factory");
		
		Element style = createStyle("TYOtest");		
		Element font = createFont();
		style.appendChild(font);
		
		xfactory.addStyle(style);
		
		System.out.println("now there is " + xfactory.getStylesCount() + " style in the factory");
		
		
		new XSerializer().serialize(wb);
		for (String warning : wb.getWarnings()) {
			System.out.println(warning);
		}
	}
	
	
	public static Element createStyle(String id) throws ParserConfigurationException {
		Element style = getDoc().createElement("Style");
		Attr attr = getDoc().createAttributeNS(XLElement.XMLNS_SS, "ID");
		attr.setPrefix(XLElement.PREFIX_SS);
		attr.setNodeValue(id);
		style.setAttributeNodeNS(attr);
		return style;
	}

	public static Element createFont() throws ParserConfigurationException {
		Element font = getDoc().createElement("Font");
		Attr attr = getDoc().createAttributeNS(XLElement.XMLNS_SS, "Color");
		attr.setPrefix(XLElement.PREFIX_SS);
		attr.setNodeValue("#00FF00");
		font.setAttributeNodeNS(attr);
		return font;
	}
	
	private static Document getDoc() throws ParserConfigurationException {
		if (document == null) {
			document = new XLWorkbook().createDocument();
		}
		return document;
	}

}
