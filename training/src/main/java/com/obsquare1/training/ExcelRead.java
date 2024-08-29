package com.obsquare1.training;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead{

	static FileInputStream fis;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	public static String readStringData(int row,int column) throws Exception {
	fis =new FileInputStream("C:\\Users\\ATHIRA\\OneDrive\\Documents\\data.xlsx");
	workbook= new XSSFWorkbook(fis);
	sheet =workbook.getSheet("Sheet1");
	XSSFRow r=sheet.getRow(row);
	XSSFCell c= r.getCell(column);
	return c.getStringCellValue();
	}
	public static double readNumericData(int row,int column) throws Exception {
		fis=new FileInputStream("C:\\Users\\ATHIRA\\OneDrive\\Documents\\data.xlsx");
		workbook = new XSSFWorkbook(fis);
		sheet=workbook.getSheet("Sheet1");
		XSSFRow r= sheet.getRow(row);
		XSSFCell c= r.getCell(column);
		return c.getNumericCellValue();
	}
	public static void main(String[] args) throws Exception  {
	System.out.println(ExcelRead.readStringData(0, 0));
	double d=ExcelRead.readNumericData(1, 1);
	System.out.println(d);
	}
		


	}



