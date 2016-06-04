import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

/**
 * Created by Thilina on 6/4/2016.
 */
public class ReadExcell {

    public static void main(String arg[]) throws IOException {

        //open xlsx file
        String excelFilePath = "excel//test1.xlsx";
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));

        //create new work book for the opened file
        Workbook workbook = new XSSFWorkbook(inputStream);
        //get the 1st work sheet
        Sheet firstSheet = workbook.getSheetAt(0);
        //create a iterator to go through rows

        // for row has a next
        for(Row nextRow : firstSheet) {

            //go through each cell (column)
            Iterator<Cell> cellIterator = nextRow.cellIterator();

            //iterate through columns
            while(cellIterator.hasNext()) {
                // print each cell value
                Cell cell = cellIterator.next();

                int cellType = cell.getCellType();

                switch(cellType) {
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getStringCellValue() + " s ");
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        System.out.print(cell.getBooleanCellValue() + " b ");
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(cell.getNumericCellValue() + " n ");
                        break;
                }

                if(cellType == Cell.CELL_TYPE_BLANK) {
                    System.out.print(" same ");
                }

                else{
                    System.out.print(" - ");
                }

            }
            System.out.println();
        }

        workbook.close();
        inputStream.close();


    }
}
