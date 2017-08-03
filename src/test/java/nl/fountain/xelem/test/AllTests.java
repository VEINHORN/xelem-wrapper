/*
 * Created on Sep 8, 2004
 *
 */
package nl.fountain.xelem.test;

import java.io.File;

import junit.framework.TestCase;
import junit.framework.TestSuite;
import junit.textui.TestRunner;

/**
 *
 */
public class AllTests extends TestCase {
	
	private static final String[] T_O_DIRS = {"testoutput/CreateDocumentTest", 
												"testoutput/XLDocumentTest",
												"testoutput/ReaderTest"};

    
    public static void main(String[] args) {
    	// create testoutput directories
    	for (int i = 0; i < T_O_DIRS.length; i++) {
        	File output = new File(T_O_DIRS[i]);
        	if (!output.exists()) {
        		output.mkdirs();
        	}
		}
    	// run suite
        TestRunner.run(suite());
    }
    
    public AllTests(String arg0) {
		super(arg0);
	}
    
    public static TestSuite suite() {
        TestSuite suite = new TestSuite("AllTests");
        
        suite.addTestSuite(AddressTest.class);
        suite.addTestSuite(AreaTest.class);
        suite.addTestSuite(CellPointerTest.class);       
        suite.addTestSuite(XLUtilTest.class);
        suite.addTestSuite(SSCellTest.class);
        suite.addTestSuite(SSCommentTest.class);
        suite.addTestSuite(SSColumnTest.class);
        suite.addTestSuite(SSRowTest.class);
        suite.addTestSuite(SSTableTest.class);
        suite.addTestSuite(XWorksheetOptionsTest.class);
        suite.addTestSuite(XPaneTest.class);
        suite.addTestSuite(SSWorksheetTest.class);
        
        suite.addTestSuite(XFactoryTest.class);
        suite.addTestSuite(XLWorkbookTest.class);
        suite.addTestSuite(CreateDocumentTest.class);
        
        suite.addTestSuite(XLDocumentTest.class);
        
        suite.addTestSuite(ExcelReaderTest.class);
        suite.addTestSuite(DirectorTest.class);
        suite.addTestSuite(ExcelReaderListenerTest.class);
        suite.addTestSuite(WorkbookListenerTest.class);
        
        return suite;
    }

}
