package com.fillo;



import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {

	public void writeToExcelSheet(String fileName, int countA, int countB, String status, int row) {
		// Read the spreadsheet that needs to be updated
		FileInputStream fsIP = null;
		try {
			fsIP = new FileInputStream(new File(fileName));
		} catch (FileNotFoundException e) {
		}
		// Access the workbook
		Workbook wb = null;
		try {
			wb = new XSSFWorkbook(fsIP);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		// Access the worksheet, so that we can update / modify it.
		Sheet worksheet = wb.getSheetAt(0);
		// declare a Cell object
		Cell cell = null;
		// Access the third cell in each row to update the value
		cell = worksheet.getRow(row).getCell(5);
		// Get current cell value value and overwrite the value
		cell.setCellValue(countA);
		cell = worksheet.getRow(row).getCell(6);
		cell.setCellValue(countB);
		cell = worksheet.getRow(row).getCell(7);
		cell.setCellValue(status);
		// Close the InputStream
		try {
			fsIP.close();
			System.out.println("Done");
		} catch (IOException e1) {
		}
		// Open FileOutputStream to write updates
		FileOutputStream bfWriter = null;
		try {
			bfWriter = new FileOutputStream(fileName);
		} catch (FileNotFoundException e) {
		}
		// write changes
		try {
			wb.write(bfWriter);
			bfWriter.close();
		} catch (IOException e) {
		}
	}
}
