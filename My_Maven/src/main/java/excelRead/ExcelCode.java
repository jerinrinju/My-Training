package excelRead;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCode
{
	XSSFSheet sh;
	ExcelCode() throws IOException
	{
		FileInputStream f=new FileInputStream("C:\\Users\\rinju003gmail.com\\Desktop\\Excel.xlsx");
		XSSFWorkbook w= new XSSFWorkbook(f); 
		sh=w.getSheet("Sheet1");
		
	}
	public String read(int i,int j)
	{
		Row r=sh.getRow(i);
		Cell c=r.getCell(j);
		int cellType=c.getCellType();
		switch(cellType)
		{
		case Cell.CELL_TYPE_NUMERIC://static final variable
		{
			double d=c.getNumericCellValue();
			return String.valueOf(d);
		}
		case Cell.CELL_TYPE_STRING:
		{
			return c.getStringCellValue();
		}
	}
		
		return null;
		
		
		
	}
	

}
