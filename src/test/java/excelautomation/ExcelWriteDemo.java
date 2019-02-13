package excelautomation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ExcelWriteDemo {


    @Test
    public void writeExcel()throws Exception{

        //getting xls directory
        String filePath = "./src/test/resources/Countries.xls";

        FileInputStream in = new FileInputStream(filePath);

        Workbook workbook = WorkbookFactory.create(in);

        Sheet workSheet = workbook.getSheetAt(0);

        //Write colum name

        Cell colum = workSheet.getRow(0).createCell(2);
        if (colum==null){
            colum = workSheet.getRow(0).createCell(2);
        }
        colum.setCellValue("Continent");

        Cell cont1 = workSheet.getRow(1).createCell(2);
        if (cont1==null){
            workSheet.getRow(1).createCell(2);
        }
        cont1.setCellValue("North America");

        //Save changes111
        //Open the file to Write into it

        FileOutputStream out = new FileOutputStream(filePath);

        //Write and save the changes
        workbook.write(out);

        out.close();
        workbook.close();
        in.close();


    }
}
