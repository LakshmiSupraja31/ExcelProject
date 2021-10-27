package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DateFormate {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		File f = new File("C:\\Users\\Vinoth kumar\\eclipse-workspace\\ExcelProject\\Excel\\sample_Excel.xlsx");
	FileInputStream fin = new FileInputStream(f);
	Workbook w = new XSSFWorkbook(fin);
	Sheet sheet = w.getSheet("Sheet1");
	Row row = sheet.getRow(1);
	Cell cell = row.getCell(4);
	System.out.println(cell);
	int cellType = cell.getCellType();
	System.out.println(cellType);
	if(cellType==1) {
		String stringcellValue = cell.getStringCellValue();
		System.out.println(stringcellValue);
	}
	else if (cellType==0) {
		if (DateUtil.isCellDateFormatted(cell)) {
	// current cell value in date format
		Date dateCellValue=cell.getDateCellValue();
		SimpleDateFormat sim = new SimpleDateFormat("dd-MM-yyyy");
		String format=sim.format(dateCellValue);
		System.out.println(format);
		}
		else
		{
			// current cell value is in number format
			double numericCellValue = cell.getNumericCellValue();
			long l = (long)numericCellValue;
			String ValueOf = String.valueOf(l);
			System.out.println(ValueOf);
		}
		}
	}

}
