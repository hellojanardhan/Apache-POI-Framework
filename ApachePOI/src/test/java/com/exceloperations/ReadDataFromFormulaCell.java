package com.exceloperations;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataFromFormulaCell {

	public static void main(String[] args) throws IOException {
		
		//reading data from excel sheet
		FileInputStream inputStream=new FileInputStream(new File(".\\DataFiles\\read formula cells.xlsx"));
		
		//creating workbook
		XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
		
		//getting sheet from excel
		XSSFSheet sheet=workbook.getSheet("read formula cells");
		
		//counting rows and columns
		int rows=sheet.getLastRowNum();
		int cols=sheet.getRow(0).getLastCellNum();
		
		
		for (int r = 0; r <rows; r++) {
			XSSFRow row=sheet.getRow(r);
			for (int c = 0; c <cols; c++) {
				XSSFCell cell=row.getCell(c);
				
				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				case FORMULA:
					System.out.print(cell.getNumericCellValue());
					break;
				default:
					break;
				}
				System.out.print("   |   ");
			}
			System.out.println();
		}
		inputStream.close();
		workbook.close();
	}

}
