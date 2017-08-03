package nl.fountain.xelem.ztest;

import javax.xml.parsers.ParserConfigurationException;

import nl.fountain.xelem.XFactory;
import nl.fountain.xelem.XSerializer;
import nl.fountain.xelem.XelemException;
import nl.fountain.xelem.excel.Cell;

import org.w3c.dom.Element;

public class StyleFactoryWorkbookTest {

	public static void main(String[] args)  {
		StyleFactoryWorkbook sfwb = new StyleFactoryWorkbook("test");
		Cell cell = sfwb.addSheet().addCell("text should be green");
		cell.setStyleID("TYOtest");
		
		// before serializing, add styles:
		XFactory xfactory = sfwb.getFactory();		
		try {			
			Element style = sfwb.createStyle("TYOtest");		
			Element font = sfwb.createFont();
			style.appendChild(font);		
			xfactory.addStyle(style);
			
			// now serialize it
			new XSerializer().serialize(sfwb);
			
			System.out.println("created the workbook");
			
		} catch (ParserConfigurationException e) {
			e.printStackTrace();
		} catch (XelemException e) {
			e.printStackTrace();
		}
	}

}
