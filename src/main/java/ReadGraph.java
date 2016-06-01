import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;



import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

/**
 * Created by Thilina on 6/1/2016. (Not working for now)
 */
public class ReadGraph {

    public static void main(String arg[]) throws IOException {

        String filename = "ppts\\testin.pptx";

        //open ppt file
        FileInputStream input = new FileInputStream(filename);

        //create XMLSlideShow object
        XMLSlideShow ppt = new XMLSlideShow(input);

        //get all slides
        List<XSLFSlide> slides = ppt.getSlides();

        //get slide with graph
        XSLFSlide slide = slides.get(2);

        //list all document part relations
        List<POIXMLDocumentPart> relations = slide.getRelations();

        //variable to store chart
        XSLFChart chart = null;

        //find chart
        for (POIXMLDocumentPart chr : relations) {
            if (chr instanceof XSLFChart) {
                chart = (XSLFChart) chr;
            }
        }

        CTChart ctChart = chart.getCTChart();

        CTPlotArea plotArea = ctChart.getPlotArea();


    }
}
