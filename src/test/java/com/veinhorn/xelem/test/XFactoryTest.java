package com.veinhorn.xelem.test;

import nl.fountain.xelem.XFactory;
import nl.fountain.xelem.XSerializer;
import nl.fountain.xelem.XelemException;
import nl.fountain.xelem.excel.Row;
import nl.fountain.xelem.excel.Workbook;
import nl.fountain.xelem.excel.Worksheet;
import nl.fountain.xelem.excel.ss.XLWorkbook;
import org.junit.Test;

import java.net.URL;

/**
 * Created by Boris Korogvich on 28.07.2017.
 */
public class XFactoryTest {
    @Test
    public void testXFactory() throws XelemException {
        URL resource = getClass().getResource("/xelem-config.xml");
        XFactory.setConfigurationFileName(resource.getPath());

        Workbook workbook = new XLWorkbook();

        workbook.setFileName("imported.xml");

        Worksheet sheet = workbook.addSheet("My Sheet");

        Row row = sheet.addRow();
        row.setIndex(1);
        row.addCell("Cell data", "sHeader");

        /*row = sheet.addRowAt(2);
        row.setIndex(2);
        row.addCell("cell data 2", "sNum2Fract_T");

        row = sheet.addRowAt(3);
        row.setIndex(3);
        row.addCell("cell data 3", "sStr_T");*/

        new XSerializer().serialize(workbook);

        for (String warning : workbook.getWarnings()) {
            System.out.println(warning);
        }

        String result = "";
    }
}
