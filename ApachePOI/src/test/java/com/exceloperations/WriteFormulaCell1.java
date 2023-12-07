package com.exceloperations;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFormulaCell1 {

	public static void main(String[] args) throws IOException {
		
		String excelFilePath=".//DataFiles//Write formula cell.xlsx";
		
		//reading data form Write formula cell.xlsx
		FileInputStream inputStream=new FileInputStream(excelFilePath);
		
		//get the worbook
		XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
		
		//get the sheet from workbook
		XSSFSheet sheet=workbook.getSheet("Data");
		
		
		sheet.getRow(11).getCell(4).setCellFormula("SUM(D2:D10)");
		
		inputStream.close();
		
		//writing the new data into file
		FileOutputStream outputStream=new FileOutputStream(excelFilePath);
		
		workbook.write(outputStream);

		workbook.close();
		
		outputStream.close();
		
		System.out.println("Write formula cell.xlsx updated successfully");
		
	}

}
