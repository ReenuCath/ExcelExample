package excel.example;

import java.io.IOException;

public class ReadExcel {

	public static void main(String[] args) throws IOException 
	{
		String data;
		ExcelEg1 obj=new ExcelEg1();
		obj.readFile();
		data=obj.readData(3,1);
		System.out.println("Cell Value is: "+data);
		
	}

}
