package api.utilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLUtility {

	public FileInputStream fi;
	public FileOutputStream fo;
	public XSSFWorkbook workbook;
	public XSSFSheet sheet;
	public XSSFRow row;
	public XSSFCell cell;
	public XSSFCellStyle style;
	String path;

	public XLUtility(String path) {
		this.path = path;
	}

	public int getRowCount(String sheetName) throws IOException {

		fi = new FileInputStream(path);
		workbook = new XSSFWorkbook(fi);
		sheet = workbook.getSheet(sheetName);
		int lastRowNum = sheet.getLastRowNum();
		workbook.close();
		fi.close();
		return lastRowNum;
	}

	public int getCellCount(String sheetName, int rowNum) throws IOException {

		fi = new FileInputStream(path);
		workbook = new XSSFWorkbook(fi);
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(rowNum);
		int cellCount = row.getLastCellNum();
		workbook.close();
		fi.close();
		return cellCount;
		
	}

	public String getCellData(String sheetName, int rowNum , int cellNum) throws IOException {

		fi = new FileInputStream(path);
		workbook = new XSSFWorkbook(fi);
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(rowNum);
		cell = row.getCell(cellNum);
		DataFormatter formatter = new DataFormatter();
		String data = formatter.formatCellValue(cell);
		workbook.close();
		fi.close();
		return data;
	}
	
	public void setCellData(String sheetName, int rowNum , int cellNum, String data) throws IOException {
		
		File xlFile = new File(path);
		if (!xlFile.exists()) {
			workbook = new XSSFWorkbook(fi);
			fo = new FileOutputStream(path);
			workbook.write(fo);
		}
		
		fi = new FileInputStream(path);
		workbook = new XSSFWorkbook(fi);
		
		if(workbook.getSheetIndex(sheet)==-1) 
			workbook.createSheet(sheetName);
			sheet = workbook.getSheet(sheetName);
		
		if (sheet.getRow(rowNum)==null)
			sheet.createRow(rowNum);
			row = sheet.getRow(rowNum);
		
		cell = row.createCell(cellNum);
		cell.setCellValue(data);
		workbook.write(fo);
		workbook.close();
		fi.close();
		fo.close();
	}
	
}
