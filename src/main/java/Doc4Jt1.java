import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.OpcPackage;
import org.docx4j.openpackaging.packages.PresentationMLPackage;
import org.docx4j.openpackaging.parts.PresentationML.MainPresentationPart;
import org.docx4j.openpackaging.parts.PresentationML.SlideLayoutPart;
import org.docx4j.openpackaging.parts.PresentationML.SlidePart;
import org.pptx4j.Pptx4jException;

/**
 * Created by Instructor - ICT on 6/1/2016.
 */
public class Doc4Jt1 {

    public static void main(String arg[]) throws Docx4JException, Pptx4jException {

        String inputfilepath =  "ppts\\testin.pptx";

        PresentationMLPackage presentationMLPackage =
                (PresentationMLPackage)OpcPackage.load(new java.io.File(inputfilepath));

        MainPresentationPart mainPresentationPart = presentationMLPackage.getMainPresentationPart();


    }
}

