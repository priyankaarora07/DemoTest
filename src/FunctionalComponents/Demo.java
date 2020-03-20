package FunctionalComponents;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo {

	public static void main(String[] args) throws IOException {

		System.out.println("Welcome");
		
		FileInputStream fis = new FileInputStream("F:\\Work\\Workspace\\QBCS_Gulf\\src\\supportTools\\Billguru.xlsx");
		String path = "F:\\Work\\Workspace\\QBCS_Gulf\\src\\supportTools\\Billguru.xlsx"; 
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
	    XSSFSheet sheet = workbook.getSheet("Results");
	    XSSFRow row = sheet.getRow(0);
	    int colNum = row.getLastCellNum();
	    System.out.println("Total Number of Columns in the excel is : "+colNum);
	    int rowNum = sheet.getLastRowNum()+1;
	    System.out.println("Total Number of Rows in the excel is : "+rowNum);
	    
	    int passcount = 0;
	    int failcount = 0;
	    int vmcount = 0;
	    int i=1;
	    while(i<rowNum)
	    {
	    	String finalStatus = readFromExcel(path, i, 1);
	    	
	    	if(finalStatus.equalsIgnoreCase("Pass"))
	    	{
	    		passcount++;
	    	}
	    	else if(finalStatus.equalsIgnoreCase("Fail"))
	    	{
	    		failcount++;
	    	}
	    	else if(finalStatus.equalsIgnoreCase("VM"))
	    	{
	    		vmcount++;
	    	}
	    	
	    	i++;

	    }
	    
	    System.out.println("Pass Count : "+passcount+"\nFail Count : "+failcount+"\nVM Count : "+vmcount);
	    	
	}
	
	public static String readFromExcel(String file, int rownum, int colnum) throws IOException {
		
        XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet myExcelSheet = myExcelBook.getSheet("Results");
        XSSFRow row = myExcelSheet.getRow(rownum);
        
        String status = "";
        
        if(row.getCell(colnum).getCellType() == XSSFCell.CELL_TYPE_STRING){
            status = row.getCell(colnum).getStringCellValue();
//          System.out.println("Status : " + status);
        }
        
        myExcelBook.close();
		return status;
        
    }
	
}
