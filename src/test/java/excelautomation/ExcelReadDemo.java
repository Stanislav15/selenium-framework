package excelautomation;

import org.apache.poi.ss.usermodel.*;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;

public class ExcelReadDemo {
    @Test
     public void readXLFile() throws Exception {
        String path = "/Users/mrs/Desktop/Countries.xls";

        //open file and convert to stream of data
        FileInputStream inputStream = new FileInputStream(path);

        //WorkBook>Worksheet>rows>cell -
        //open the workBook.Any type

        Workbook workbook = WorkbookFactory.create(inputStream);

        //go to the First WorkSheet . Index 0
        Sheet worksheet = workbook.getSheetAt(0);

        //Go to the first row.
        Row row = worksheet.getRow(0);

        //go to first Cell
        Cell cell = row.getCell(0);
        Cell cell2 = row.getCell(1);

        //print cell values
        System.out.println(cell.toString());
        System.out.println(cell2.toString());


        //read  cell value using method chaining
        String country1 = worksheet.getRow(1).getCell(0).toString();

        //second way using dif workbooks
        String capital1 = workbook.getSheetAt(0).getRow(1).getCell(1).toString();

        System.out.println("Country1: " + country1);
        System.out.println("Capital1: "+ capital1);


        int rowsCount = worksheet.getLastRowNum();
        System.out.println("Number of rows: "+ rowsCount);


        for (int i=1; i<=rowsCount;i++){
            System.out.println("country #: "+i+ ": "+worksheet.getRow(i).getCell(0).toString()+
                    " ==> "+worksheet.getRow(i).getCell(1).toString());

        }

        Map<String ,String > countriesmap= new HashMap<>();

        int countryCol=0;
        int capitalCol=1;
        for (int i =1; i<=rowsCount;i++){
            countriesmap.put(worksheet.getRow(i).
                    getCell(countryCol).
                    toString(),
                    worksheet.getRow(i).
                            getCell(capitalCol).
                            toString());
        }

        System.out.println(countriesmap);

        //close workbook and stream
        workbook.close();
        inputStream.close();
    }
}
