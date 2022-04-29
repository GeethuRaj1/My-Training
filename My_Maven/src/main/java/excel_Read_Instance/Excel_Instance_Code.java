package excel_Read_Instance;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_Instance_Code {
	
	
	static XSSFSheet sh;
	
	public Excel_Instance_Code() throws IOException
	{
		FileInputStream file=new FileInputStream("C:\\Users\\User\\Desktop\\Book1.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(file);
		sh=wb.getSheet("Data");

	}
	
	public String readSheet(int row,int column)throws Exception
	{
		System.out.println("New");
		Row r=sh.getRow(row);
		Cell c=r.getCell(column);
		int type=c.getCellType();
		
		switch(type)
		{
		
		case Cell.CELL_TYPE_NUMERIC:
		{
			int v=(int) c.getNumericCellValue(); 
			return String.valueOf(v);
		}
		
		case Cell.CELL_TYPE_STRING:
		{
			return c.getStringCellValue();
		}
		
		default:
			System.out.println("Data Not Found");
			break;
		}
		
		return null;
		
	}
	

}
