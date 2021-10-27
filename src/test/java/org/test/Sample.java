package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Sample {

	public static void main(String[] args) throws IOException {
		// path/location of the file
		
File f = new File("C:\\Users\\Vinoth kumar\\eclipse-workspace\\ExcelProject\\Excel\\sample_Excel.xlsx");
	// to get into the file
FileInputStream fin = new FileInputStream(f);
	//to get into the workbook
Workbook w = new XSSFWorkbook(fin);
//to get into the sheet
 Sheet sheet = w.getSheet("Sheet1");
//to get into the row
Row row = sheet.getRow(1);
//to get into the cell
Cell cell = row.getCell(1);
System.out.println(cell);
   // int physicalNumberOfRows=sheet.getPhysicalNumberOfRows();
    // System.out.println("number of rows:" +physicalNumberOfRows);

   // int physicalNumberOfCells = row.getPhysicalNumberOfCells();
   // System.out.println("number of cells:"+physicalNumberOfCells);

   //  to print all the value of a particular cell in all the rows

for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
	Row row2 = sheet.getRow(i);
	Cell cell2 =row2.getCell(2);
	System.out.println(cell2);
}
// to print all the values in a row
for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
	Cell cell2 = row.getCell(i);
	System.out.println(cell2);
	
}
// to print all the values in a sheet
for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
	Row row2 = sheet.getRow(i);
for (int j = 0; j < row2.getPhysicalNumberOfCells(); j++) {
	Cell cell2 = row2.getCell(j);
	System.out.println(cell2);
	}
	
}


	}

}
