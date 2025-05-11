package com.datadriven;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDrivenFramework {

	// Excel File --->Workbook ---> Sheet ---> Row ---> Cell
	// FileInputStream ---> WorkbookFactory ---> XSSFWorkbook---> Read Data
	// FileOutputStream ---> WorkbookFactory ---> XSSFWorkbook---> Write Data
	
	public static void main(String[] args) throws IOException {
	
		// Read the Excel file
		
		FileInputStream file = new FileInputStream(System.getProperty("user.dir")+"\\testdata\\data.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		
		XSSFSheet sheet =workbook.getSheet("Sheet1");
		
		int TotalRows = sheet.getLastRowNum();
		int TotalCells = sheet.getRow(0).getLastCellNum();

		System.out.println("Total Rows: " + TotalRows);
		System.out.println("Total Cells: " + TotalCells);
		
		for(int r =0;r<=TotalRows;r++)
		{
			XSSFRow currentRow = sheet.getRow(r);
			
			for(int c=0;c<TotalCells;c++)
			{
				XSSFCell cell = currentRow.getCell(c);
				System.out.print(cell.toString()+"\t");
			}
				System.out.println();
		}
		
		workbook.close();
		file.close();
	}

}
