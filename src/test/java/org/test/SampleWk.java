package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class SampleWk  {

	public static void main(String[] args) throws IOException {
		
File f = new File("C:\\Users\\Vinoth kumar\\eclipse-workspace\\ExcelProject\\Excel\\sample_Excel.xlsx");
FileInputStream fin = new FileInputStream(f);
Workbook w = new XSSFWorkbook (fin);
Sheet sheet = w.getSheet("New");
Row row = sheet.getRow(1);
Cell cell = row.getCell(1);
String s = cell.getStringCellValue();
if (s.equals("jhone")) {
	cell.setCellValue("chennai");
}

FileOutputStream fout = new FileOutputStream(f);
w.write(fout);
		System.out.println("done...");
	}

	}


