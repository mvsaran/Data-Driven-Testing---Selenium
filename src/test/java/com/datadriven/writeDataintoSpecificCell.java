package com.datadriven;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writeDataintoSpecificCell {

	public static void main(String[] args) throws IOException {
		
		FileOutputStream file  = new FileOutputStream(System.getProperty("user.dir")+"\\testdata\\writeDataSpecificCell.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		XSSFSheet sheet = workbook.createSheet("Sheet1");
		
		XSSFRow row = sheet.createRow(0);
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("DataDrivenTesting");
		
		workbook.write(file);
		workbook.close();
		file.close();
		System.out.println("Data written successfully to Excel file.");


	}

}
