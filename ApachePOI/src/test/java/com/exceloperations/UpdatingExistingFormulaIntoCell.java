package com.exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UpdatingExistingFormulaIntoCell {

	public static void main(String[] args) throws IOException {
		
		//location of empty excel file
		String excelFilePath=".//DataFiles//Employee info.xlsx";
		
		//Data written info excel sheet
		Object employeeData[][]= {
				{"EmpID","EmpName","EmpSalary","EmpAddress"},
				{101,"Suresh",100000,"Hyderabad"},
				{102,"Ramesh",200000,"Bangalore"},
				{103,"Vignesh",300000,"Chennai"}
		};
		
		
		
		//creating workbook
		XSSFWorkbook workbook=new XSSFWorkbook();
		
		//creating sheet from workbook
		XSSFSheet sheet=workbook.createSheet("Employee Info");
		
		//writing data into sheet
		FileOutputStream outputStream=new FileOutputStream(excelFilePath);
		
	
		
		for (int r= 0; r <employeeData.length; r++) {
			//creating row
			XSSFRow row=sheet.createRow(r);

			for (int c = 0; c < employeeData[r].length; c++) {
				//creating cell
				XSSFCell cell=row.createCell(c);
				
				Object values=employeeData[r][c];
				
				if (values instanceof String) {
					cell.setCellValue((String)values);
				}
				if(values instanceof Integer) {
				cell.setCellValue((Integer)values);
				}
			
			}
		}
		
		workbook.write(outputStream);
		System.out.println("employeeData written successfully into employee info.xlsx file");
		
		workbook.close();
		outputStream.close();
		
	}

}
