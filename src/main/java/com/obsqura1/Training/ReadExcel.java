package com.obsqura1.Training;

import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	
	static FileInputStream fi;
	static XSSFWorkbook wb;
	static XSSFSheet sh;
	
	public static String readstring(int i,int j) throws Exception{
		
		fi =new FileInputStream("C:\\Users\\nayan\\OneDrive\\Documents\\testexcel.xlsx");
		wb=new XSSFWorkbook(fi);
		sh=wb.getSheet("test");
		XSSFRow row=sh.getRow(i);
		XSSFCell cell=row.getCell(j);
		return cell.getStringCellValue();//give the data type that has to be returned
	}
	
 public static double readnumber(int i,int j) throws Exception{
		
		fi =new FileInputStream("C:\\Users\\nayan\\OneDrive\\Documents\\testexcel.xlsx");
		wb=new XSSFWorkbook(fi);
		sh=wb.getSheet("test");
		XSSFRow row=sh.getRow(i);
		XSSFCell cell=row.getCell(j);
		return cell.getNumericCellValue();//give the data type that has to be returned
	}
	

	public static void main(String[] args) throws Exception{
		
		String value= ReadExcel.readstring(1, 0);
		System.out.println(value);
		double salary= ReadExcel.readnumber(1, 1);
		System.out.println(salary);

	}

}
