import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.OpcPackage;
import org.docx4j.openpackaging.packages.PresentationMLPackage;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

/**
 * Created by Thilina on 6/1/2016.
 */
public class Docx4JChartEx {

    public static void main(String arg[]) throws Docx4JException, FileNotFoundException {

        String filename = "ppts\\testin.pptx";

        //open ppt file
        FileInputStream input = new FileInputStream(filename);

       // PresentationMLPackage presentationMLPackage = PresentationMLPackage.load(input);
        //PresentationMLPackage presentationMLPackage= (PresentationMLPackage) OpcPackage.load(in);

    }

}
