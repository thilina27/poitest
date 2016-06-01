import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFGraphicFrame;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSheet;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class PPTCharts {

    public static void main(String args[]) throws InvalidFormatException,
            IOException{

        XMLSlideShow ppt;

        // Read pptx template
        ppt = new XMLSlideShow(new FileInputStream("ppts\\chart.pptx"));

        // Get all slides
        List<XSLFSlide> slide = ppt.getSlides();

        // Get working slide that is slide=0
        XSLFSlide slide0 = slide.get(0);
        List<XSLFShape> shapes = slide0.getShapes();

        // Add all shapes into a Map
        Map <String, XSLFShape> shapesMap = new HashMap<String, XSLFShape>();
        for(XSLFShape shape : shapes)
        {
            shapesMap.put(shape.getShapeName(), shape);
            System.out.println("Shape names " + shape.getShapeName() + "Is this ");
            System.out.println(shape.getShapeName() + "  " + shape.getShapeId() +" "+ shape);

        }

        // Read the bar chart
        XSLFGraphicFrame chart = (XSLFGraphicFrame) shapesMap.get("Chart 8");

        // Get the chart sheet
        XSLFSheet sheet =  chart.getSheet();

        for(int i=0; i<sheet.getRelations().size(); i++)
        {
            System.out.println("Partname =" +
                    sheet.getRelations().get(i).getPackagePart().getPartName());




            if(sheet.getRelations().get(i).getPackagePart().getPartName().toString().contains(".xls"))
            {

                System.out.println("Found the bar chart excel");

                // BarChart Excel package part
                PackagePart barChartExcel  =
                        sheet.getRelations().get(i).getPackagePart();

                // Reference the excel in workbook
                HSSFWorkbook wb = new HSSFWorkbook(barChartExcel.getInputStream());

                // Read sheet where Barchart data is available
                HSSFSheet mysheet =  wb.getSheetAt(1);

                // Read first
                HSSFRow row = mysheet.getRow(1);


                //Print first cell value for debugging
                System.out.println("Updating cell value from - " + row.getCell(1));

                // New value
                double insertValue = 7777777.0;


                wb.getSheetAt(1).getRow(1).getCell(1).setCellValue(insertValue);

                // Set first BarChart as active sheet
                HSSFSheet mysheet0 =  wb.getSheetAt(0);
                mysheet0.setActive(true);

                // Write the updated excel back to workbook
                OutputStream excelOut = barChartExcel.getOutputStream();
                excelOut.flush();
                wb.write(excelOut);
                excelOut.close();

                // Write workbook to file
                FileOutputStream o = new FileOutputStream("MyPresentation.pptx");
                ppt.write(o);
                o.close();
                System.out.println("new ppt is created....");

                break; // Exit
            }

        }
    }
}