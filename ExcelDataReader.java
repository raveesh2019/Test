package com.fillo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataReader {

	public static final String FILE_NAME = "C:\\Users\\09861226\\Desktop\\Selenium\\ERBI\\Count.xlsx";

	public List<String> excelDataList = null;
	public int cellCount = 0;
	public int columnCount = 1;
	public int sheetNumberCount = 0;

	public int getSheetNumberCount() {
		return sheetNumberCount;
	}

	public void setSheetNumberCount(int sheetNumberCount) {
		this.sheetNumberCount = sheetNumberCount;
	}

	public int getColumnCount() {
		return columnCount;
	}

	public void setColumnCount(int columnCount) {
		this.columnCount = columnCount;
	}

	public int getCellCount() {
		return cellCount;
	}

	public void setCellCount(int cellCount) {
		this.cellCount = cellCount;
	}

	public List<String> getExcelDataList() {
		return excelDataList;
	}

	public void setExcelDataList(List<String> excelDataList) {
		this.excelDataList = excelDataList;
	}

	@SuppressWarnings("deprecation")
	public void excelSheetData() {
		excelDataList = new ArrayList<>();
		try {
			FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
			@SuppressWarnings("resource")
			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet datatypeSheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = datatypeSheet.iterator();
			int i = 0;
			while (iterator.hasNext()) {
				Row currentRow = iterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();
				while (cellIterator.hasNext()) {
					Cell currentCell = cellIterator.next();
					if (currentCell.getCellTypeEnum() == CellType.STRING) {
						if (i >= 10 ) {
								excelDataList.add(currentCell.getStringCellValue());
						}
                } 
						cellCount = currentCell.getRowIndex();
						columnCount = currentCell.getColumnIndex() + 1;
//					}
						i++;
				}
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public static void main(String[] args) {
		ExcelDataReader excelDataReader = new ExcelDataReader();
		ExcelWrite excelWrite = new ExcelWrite();
		excelDataReader.excelSheetData();
        List<String> list = excelDataReader.getExcelDataList();
        System.out.println("list: "+(list.size()+1));
        System.out.println("Process started");
       // int totalRows = (list.size()+1)/5;
        int row = 0;
        int j = 0;
        int sqlCountA = 0;
        int sqlCountB = 0;
        Connection conn = DbConnection.getConnection();
        Statement st = null;
        ResultSet rs = null;
        Statement st1 = null;
        ResultSet rs1 = null;
        String status = null;
		for (String string : list) {
			if (j < list.size()) {
				if (list.get(j).equals("Y")) {
					row++;
					try {
						st = conn.createStatement();
						rs = st.executeQuery(list.get(j+1));
						while (rs.next()) {
							sqlCountA = rs.getInt("TestCnt");
							System.out.println("sqlCountA: "+sqlCountA);
						}
						st1 = conn.createStatement();
						rs1 = st.executeQuery(list.get(j+2));
						
						while (rs1.next()) {
							sqlCountB = rs1.getInt("TestCnt");
							System.out.println("sqlCountB: "+sqlCountB);
						}
						
						if ( sqlCountA == sqlCountB) {
							status = "Pass";
							System.out.println("in pass: "+status);
						}else{
							status = "Fail";
							System.out.println("in fail: "+status);
						}
						excelWrite.writeToExcelSheet(FILE_NAME, sqlCountA, sqlCountB, status, row);
					} catch (SQLException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					finally {
						try {
//							st.close();
//							st1.close();
							if (conn != null && list.get(j+3) == null) {
								System.out.println("closing connection");
								conn.close();
							}
						} catch (SQLException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						
					}
					
				}else{
					row++;
				}
				j = j+5;
			}
		}
	}
}
