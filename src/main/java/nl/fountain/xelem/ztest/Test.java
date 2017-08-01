/*
 * Created on 22-mrt-2005
 *
 */
package nl.fountain.xelem.ztest;

import java.io.IOException;

import javax.xml.parsers.ParserConfigurationException;

import nl.fountain.xelem.Area;
import nl.fountain.xelem.excel.Row;
import nl.fountain.xelem.excel.Worksheet;
import nl.fountain.xelem.lex.DefaultExcelReaderListener;
import nl.fountain.xelem.lex.ExcelReader;

import org.xml.sax.SAXException;



/**
 *
 */
public class Test {

    public static void main(String[] args) throws ParserConfigurationException, IOException, SAXException {
        new Test().listen();
    }
    
    private void simple() throws ParserConfigurationException, IOException, SAXException {
        ExcelReader reader = new ExcelReader();
        /*Workbook wb = */reader.getWorkbook("D:/test/big1879.xml");
    }

    private void listen() throws ParserConfigurationException, SAXException, IOException {
        ExcelReader reader = new ExcelReader();
        //reader.setReadArea(new Area("E11:M16"));
        reader.addExcelReaderListener(new DefaultExcelReaderListener() {
            public void setRow(int sheetIndex, String sheetName, Row row) {
                // process row and/or it's cells
                System.out.println(sheetIndex + " \t" + sheetName + " \t" + row.getIndex());
            }
        });
        reader.read("testsuitefiles/ReaderTest/reader.xml");
    }
    
    class SwappingAreaListener extends DefaultExcelReaderListener {
        
       private ExcelReader reader;

       public SwappingAreaListener(ExcelReader reader) {
          this.reader = reader;
       }
       
       // override method in DefaultExcelReaderListener
       public void startWorksheet(int sheetIndex, Worksheet sheet) {
          switch (sheetIndex) {
             case 1:
                reader.setReadArea(new Area("A1:C6"));
                break;
             case 2:
                reader.setReadArea(new Area("G11:G11"));
                break;
             default:
                reader.clearReadArea();
         }        
       }  
    }

    
}
