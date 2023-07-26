package excel.example;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelEg1
{
	XSSFSheet sh;
	public ExcelEg1() throws IOException
	{
		FileInputStream f=new FileInputStream("C:\\Users\\Reenu\\eclipse-workspace\\MavenProject\\src\\main\\resources\\Excel1.xlsx");
		XSSFWorkbook w=new XSSFWorkbook(f);
		sh=w.getSheet("Sheet1");
	}
	
	public void readFile()
	{
		for(Row r:sh)
		{
			for(Cell c:r)
			{
				System.out.print(c+" ");
			}
			System.out.println();
		}
	}
	
	
	@SuppressWarnings("deprecation")
	public String readData(int i,int j)
	{
		Row r=sh.getRow(i);
		Cell c=r.getCell(j);
		//Cell c=sh.getRow(i).getCell(j);
		int cellType=c.getCellType();
		switch(cellType)
		{
		case Cell.CELL_TYPE_STRING:
			String data1=c.getStringCellValue();
			return data1;
		case Cell.CELL_TYPE_NUMERIC:
			double num=c.getNumericCellValue();
			String data2=String.valueOf(num);
			return data2;
		default:
			return null;
			
		}
}
}
