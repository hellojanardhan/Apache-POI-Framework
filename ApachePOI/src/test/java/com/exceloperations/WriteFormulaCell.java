package com.exceloperations;



import java.io.FileOutputStream;
import java.io.IOException;


import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFormulaCell {

	public static void main(String[] args) throws IOException {
		
		String excelFilePath=".//DataFiles//Write formula.xlsx";
		
		
		//creating new workbook
		XSSFWorkbook workbook=new XSSFWorkbook();
		
		//create new sheet in workbook
		XSSFSheet sheet=workbook.createSheet("Data");
		
		//create new row in above sheet
		XSSFRow row=sheet.createRow(0);
		
		//writing data into excel sheet
		FileOutputStream outputStream=new FileOutputStream(excelFilePath);
		
		//create cells
		row.createCell(0).setCellValue(100);
		row.createCell(1).setCellValue(100);
		row.createCell(2).setCellValue(100);
		row.createCell(3).setCellFormula("A1*B1*C1");
		
		workbook.write(outputStream);
		
		outputStream.close();
		
		System.out.println("Write formula.xlsx file created successfully........");
		
		
		
	}

}
