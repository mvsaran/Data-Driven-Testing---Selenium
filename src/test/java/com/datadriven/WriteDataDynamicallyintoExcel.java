package com.datadriven;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataDynamicallyintoExcel {

	public static void main(String[] args) throws IOException {
		
		FileOutputStream file  = new FileOutputStream(System.getProperty("user.dir")+"\\testdata\\writeDataDynamically.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		XSSFSheet sheet = workbook.createSheet("Sheet1");
		
		// Create rows in the sheet dynamically
		
		
		Scanner scanner = new Scanner(System.in);
		System.out.println("Enter the number of rows you want to create:");
		int noofRows = scanner.nextInt();
		System.out.println("Enter the number of columns you want to create:");
		int noofColumns = scanner.nextInt();
		
		for(int r=0;r<noofRows;r++) 
		{
			XSSFRow currentRow = sheet.createRow(r);
			
			for(int c=0;c<noofColumns;c++) 
			{
				XSSFCell cell = currentRow.createCell(c);
				cell.setCellValue(scanner.next());
			}
				
			
		}
		
		workbook.write(file);
		workbook.close();
		file.close();
		scanner.close();
		System.out.println("Data written successfully to Excel file.");

	}

}
