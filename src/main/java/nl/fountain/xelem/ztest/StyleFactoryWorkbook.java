package nl.fountain.xelem.ztest;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.UnsupportedStyleException;
import nl.fountain.xelem.XFactory;
import nl.fountain.xelem.XelemException;
import nl.fountain.xelem.excel.NamedRange;
import nl.fountain.xelem.excel.Worksheet;
import nl.fountain.xelem.excel.XLElement;
import nl.fountain.xelem.excel.ss.XLWorkbook;

import org.w3c.dom.Attr;
import org.w3c.dom.DOMImplementation;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

// usefull when no configuration file is available.
public class StyleFactoryWorkbook extends XLWorkbook {

	protected XFactory xFactory;
	private Document doc;
	
	public StyleFactoryWorkbook() {
		super("xelem");
	}

	public StyleFactoryWorkbook(String name) {
		super(name);
	}
	
	// creates an empty style element
	public Element createStyle(String id) throws ParserConfigurationException {
		Element style = getDoc().createElement("Style");
		Attr attr = getDoc().createAttributeNS(XLElement.XMLNS_SS, "ID");
		attr.setPrefix(XLElement.PREFIX_SS);
		attr.setNodeValue(id);
		style.setAttributeNodeNS(attr);
		return style;
	}
	
	// Creates a font element with attribute ss:Color and value #00FF00.
	//
	// Elaborate on this method and write other more versatile methods to
	// create other style details.
	public Element createFont() throws ParserConfigurationException {
		Element font = getDoc().createElement("Font");
		Attr attr = getDoc().createAttributeNS(XLElement.XMLNS_SS, "Color");
		attr.setPrefix(XLElement.PREFIX_SS);
		attr.setNodeValue("#00FF00");
		font.setAttributeNodeNS(attr);
		return font;
	}
	
	
	////////////// overriding methods from XLWorkbook /////////////////////

	public Document createDocument() throws ParserConfigurationException {
		GIO gio = new GIO();
		Document doc = getDoc();
		Element root = doc.getDocumentElement();
		assemble(root, gio);
		return doc;
	}

	public XFactory getFactory() {
		if (xFactory == null) {
			try {
				xFactory = XFactory.newInstance();
			} catch (XelemException e) {
				// return an empty factory
				xFactory = XFactory.emptyFactory();
			}
		}
		return xFactory;
	}

	public void mergeStyles(String newID, String id1, String id2)
			throws UnsupportedStyleException {
		getFactory().mergeStyles(newID, id1, id2);
	}

	public Element assemble(Element root, GIO gio) {
		Document doc = root.getOwnerDocument();
		gio.setPrintComments(isPrintingElementComments());
		doc.insertBefore(doc.createProcessingInstruction("mso-application",
				"progid=\"Excel.Sheet\""), root);

		if (isPrintingDocComments()) {
			for (String s : getFactory().getDocComments()) {
				doc.insertBefore(doc.createComment(s), root);
			}
		}

		root.setAttribute("xmlns", XMLNS);
		root.setAttribute("xmlns:o", XMLNS_O);
		root.setAttribute("xmlns:x", XMLNS_X);
		root.setAttribute("xmlns:ss", XMLNS_SS);
		root.setAttribute("xmlns:html", XMLNS_HTML);

		if (isPrintingElementComments() && getElementComments() != null) {
			for (String s : getElementComments()) {
				root.appendChild(doc.createComment(s));
			}
		}

		//o:DocumentProperties
		if (hasDocumentProperties()) {
			super.getDocumentProperties().assemble(root, gio);
		}

		//x:ExcelWorkbook
		Element xlwbe = getExcelWorkbook().assemble(root, gio);

		//Styles
		Element styles = doc.createElement("Styles");
		root.appendChild(styles);
		appendDefaultStyle(doc, styles);

		// Names
		if (getNamedRanges().size() > 0) {
			Element names = doc.createElement("Names");
			root.appendChild(names);
			for (NamedRange nr : getNamedRanges().values()) {
				nr.assemble(names, gio);
			}
		}

		// Worksheets
		if (getWorksheets().size() < 1) {
			addSheet();
		}
		for (Worksheet ws : getWorksheets()) {
			ws.assemble(root, gio);
		}

		// append Global Information
		int selectedSheets = gio.getSelectedSheetsCount();
		if (selectedSheets > 1) {
			Element n = doc.createElementNS(XMLNS_X, "SelectedSheets");
			n.setPrefix(PREFIX_X);
			n.appendChild(doc.createTextNode("" + selectedSheets));
			xlwbe.appendChild(n);
		}
		appendStyles(doc, styles, gio);

		return root;
	}

	private Document getDoc() throws ParserConfigurationException {
		if (doc == null) {
			DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
			factory.setNamespaceAware(true);
			DocumentBuilder builder = factory.newDocumentBuilder();
			DOMImplementation domImpl = builder.getDOMImplementation();
			doc = domImpl.createDocument(XMLNS, getTagName(), null);
		}
		return doc;
	}

	private void appendDefaultStyle(Document doc, Element styles) {
		Element dse = getFactory().getStyle("Default");
		if (dse == null) {
			dse = doc.createElement("Style");

			dse.setAttributeNodeNS(createAttributeNS(doc, "ID", "Default"));
			dse.setAttributeNodeNS(createAttributeNS(doc, "Name", "Normal"));

			Element alignment = doc.createElement("Alignment");
			dse.appendChild(alignment);
			alignment.setAttributeNodeNS(createAttributeNS(doc, "Vertical",
					"Bottom"));

			dse.appendChild(doc.createElement("Borders"));
			dse.appendChild(doc.createElement("Font"));
			dse.appendChild(doc.createElement("Interior"));
			dse.appendChild(doc.createElement("NumberFormat"));
			dse.appendChild(doc.createElement("Protection"));
		} else {
			dse = (Element) doc.importNode(dse, true);
		}
		styles.appendChild(dse);
	}

	private void appendStyles(Document doc, Element styles, GIO gio) {
		for (String id : gio.getStyleIDSet()) {
			Element style = getFactory().getStyle(id);
			if (style == null) {
				// last resort: create one on the spot
				style = doc.createElement("Style");
				style.setAttributeNodeNS(createAttributeNS(doc, "ID", id));
			} else {
				// we have a style from the XFactory
				style = (Element) doc.importNode(style, true);
			}
			styles.appendChild(style);
		}
	}

}
