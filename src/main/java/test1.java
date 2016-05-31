/**
 * Created by Thilina on 5/31/2016.
 */
import com.sun.scenario.effect.impl.sw.sse.SSEBlend_SRC_OUTPeer;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFFreeformShape;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.File;
import java.io.IOException;


public class test1 {

    public static void main(String arg[]) throws IOException {
        //creating a new empty slide show
        XMLSlideShow ppt = new XMLSlideShow();

        //creating an FileOutputStream object
        File file =new File("example1.pptx");
        FileOutputStream out = new FileOutputStream(file);

        //saving the changes to a file
        ppt.write(out);
        System.out.println("Presentation created successfully");
        out.close();
    }


}
