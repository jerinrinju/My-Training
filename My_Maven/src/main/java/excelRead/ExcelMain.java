package excelRead;

import java.io.IOException;

public class ExcelMain {

	public static void main(String[] args) throws IOException {
		
		ExcelCode e=new ExcelCode();//aggregation
		String s=e.read(0,0);
		System.out.println(s);
		String s1=e.read(0,1);
		System.out.println(s1);
		
		
		// TODO Auto-generated method stub

	}

}
