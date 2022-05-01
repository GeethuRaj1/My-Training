package excel_Read;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_Code {
	
	static FileInputStream file; // FileInputStream - class
	static XSSFWorkbook w;
	static XSSFSheet s;

	public static String readStringData(int row,int column)throws IOException
	{
		file=new FileInputStream("\\C:\\Users\\User\\Desktop\\Book1.xlsx\\"); // exception FileNotFound
		w=new XSSFWorkbook(file); // exception Input Output
		s=w.getSheet("Data"); // sheet name
		System.out.println("TEST");
		//data sheet;;
	
	// find row and column ... Row, Cell are interface
	Row r=s.getRow(row);
	Cell c=r.getCell(column);
	return c.getStringCellValue();
	
}
	
	public static String readIntegerData(int row,int column) throws IOException
	{
		file=new FileInputStream("\\C:\\Users\\User\\Desktop\\Book1.xlsx\\"); // exception FileNotFound
		w=new XSSFWorkbook(file); // exception Input Output
		s=w.getSheet("Data"); // sheet name
		
	
	// find row and column ... Row, Cell are interface
	Row r=s.getRow(row);
	Cell c=r.getCell(column);
	
	int v=(int) c.getNumericCellValue();
	System.out.println("Test");
	return String.valueOf(v);
	}
}