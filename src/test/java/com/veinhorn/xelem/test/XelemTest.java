package com.veinhorn.xelem.test;

import nl.fountain.xelem.XFactory;
import nl.fountain.xelem.XSerializer;
import nl.fountain.xelem.XelemException;
import nl.fountain.xelem.excel.Workbook;
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


        /** 1st impl */
        XelemBuilder.with(workbook)
                .newSheet()
                  .addRow()
                    .addCell().data("cell 1")
                    .addCell().data("cell 2")
                    .build()
                  .addRow()
                    .addCell().data("cell 3")
                    .build()
                  .addRow()
                    .addCell().data("cell 4")
                    .addCell().data("cell 5")
                    .build().build()
                  .build();

        new XSerializer().serialize(workbook);
    }
}
