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

public class CreateCell {

	public static void main(String[] args) throws  IOException {
		
File f = new File("C:\\Users\\Vinoth kumar\\eclipse-workspace\\ExcelProject\\Excel\\sample_Excel.xlsx");
FileInputStream fin = new FileInputStream(f);
Workbook w = new XSSFWorkbook(fin);
// to create new sheet
Sheet createSheet = w.createSheet("New");

//to create new row
Row createRow = createSheet.createRow(1);

//to create new cell
Cell createCell = createRow.createCell(1);
// to set the value of the cell
createCell.setCellValue("jhone");
FileOutputStream fout = new FileOutputStream(f);

//to write the value
w.write(fout);
System.out.println("Done..");
	}

}
