/**
 * Created by Thilina on 5/31/2016.
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;


public class ReadText {

    public static void main(String[] args) throws IOException {

        String filename = "ppts\\testin.pptx";

        //open ppt file
        FileInputStream input = new FileInputStream(filename);

        //create XMLSlideShow object
        XMLSlideShow ppt = new XMLSlideShow(input);

        //get all slides
        List<XSLFSlide> slides = ppt.getSlides();

        //access slide 1 (count from 0)
        XSLFSlide slide2 = slides.get(1);

        //get place holder of the text as a TEXT shape
        XSLFTextShape shape1 = slide2.getPlaceholder(1);

        //get text from the place holder
        String text3 = shape1.getText();
        System.out.println(text3);

        //add new text paragraph
        XSLFTextParagraph shape2 = shape1.addNewTextParagraph();
        //add text to paragraph
        XSLFTextRun newTextRun = shape2.addNewTextRun();
        //replace tag
        String text4 = text3.replaceAll("<name>", "Thilina");
        newTextRun.setText(text4);

        //save file
        File file = new File("ppts\\testout.pptx");
        FileOutputStream out = new FileOutputStream(file);

        ppt.write(out);
        out.close();
        System.out.println("File save successful");

        }


}
