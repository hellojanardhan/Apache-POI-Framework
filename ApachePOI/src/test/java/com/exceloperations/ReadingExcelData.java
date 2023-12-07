package com.exceloperations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcelData {
	
	public static void main(String[] args) throws IOException {
		//location of excel sheet
		String excelFileFath=".//DataFiles//Employee Sample Data.xlsx";
		
		//reading the file
		FileInputStream inputStream=new FileInputStream(excelFileFath);
		
		//get the workbook from above file
		XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
		
		//get the sheet from above workbook
		XSSFSheet sheet=workbook.getSheet("Data");
		
		//count the number of rows in sheet
		int rows=sheet.getLastRowNum();
		
		//count the number of columns in sheet
		int cols=sheet.getRow(1).getLastCellNum();
		
		for(int r=1;r<=rows;r++) {
			XSSFRow row=sheet.getRow(r);
			for(int c=0;c<cols;c++) {
				XSSFCell cell=row.getCell(c);
				
				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell);
					break;
				case NUMERIC:
					System.out.print(cell);
				default:
					break;
				}
				System.out.print("      ");
			}
			System.out.println(" ");
		}
		workbook.close();
	}
}
