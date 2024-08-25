package com.obsquare1.training;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.common.usermodel.*;

public class ExcelRead{
	public static void main(String[] args) {
		File file=new File("C:\\Users\\ATHIRA\\OneDrive\\Documents\\data.xlsx");
		FileInputStream fis=new FileInputStream(file);
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		XSSFSheet sheet=workbook.getsheetAt(0);
		String CellValue=sheet.getRow(0).getCell(0).getStringCellValue();
;
		System.out.println(CellValue);
		workbook.close();
		fis.close();
	}
	
	
		


	}



