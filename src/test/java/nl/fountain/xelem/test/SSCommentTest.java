/*
 * Created on 31-mrt-2005
 *
 */
package nl.fountain.xelem.test;

import nl.fountain.xelem.GIO;
import nl.fountain.xelem.excel.Comment;
import nl.fountain.xelem.excel.ss.SSComment;


/**
 *
 */
public class SSCommentTest extends XLElementTest {

    public static void main(String[] args) {
        junit.textui.TestRunner.run(SSCommentTest.class);
    }
       
    public void testGetDataClean() {
        Comment comment = new SSComment();
        char lf = 10;
        comment.setData("the great author:" + lf + "dit is commentaar.");
        comment.setAuthor("the great author");
        assertEquals("dit is commentaar.", comment.getDataClean());
    }
    
    public void testAssemble() {
        Comment comment = new SSComment();
        comment.setData("commentaartekst");
        GIO gio = new GIO();
        String xml = xmlToString(comment, gio);
        assertTrue(xml.indexOf("<ss:Comment>") > 0);
        assertTrue(xml.indexOf("<ss:Data>commentaartekst</ss:Data>") > 0);
        assertTrue(xml.indexOf("</ss:Comment>") > 0);
        
        //System.out.println(xml);
    }
    

}
