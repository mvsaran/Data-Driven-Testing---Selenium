package com.datadriven;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataintoExcel {

	public static void main(String[] args) throws IOException {
		
		FileOutputStream file  = new FileOutputStream(System.getProperty("user.dir")+"\\testdata\\writeData.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		XSSFSheet sheet = workbook.createSheet("Sheet1");
		
		// Create rows in the sheet
		
		XSSFRow row1 = sheet.createRow(0);
		row1.createCell(0).setCellValue("Name");
		row1.createCell(1).setCellValue("Age");
		row1.createCell(2).setCellValue("City");
		
		XSSFRow row2 = sheet.createRow(1);
		row2.createCell(0).setCellValue("John");
		row2.createCell(1).setCellValue("30");
		row2.createCell(2).setCellValue("Hyderabad");
		
		XSSFRow row3 = sheet.createRow(2);
		row3.createCell(0).setCellValue("Smith");
		row3.createCell(1).setCellValue("25");
		row3.createCell(2).setCellValue("Bangalore");
		
		workbook.write(file);
		workbook.close();
		file.close();
		
		System.out.println("Data written successfully to Excel file.");
		
      
	}

}
