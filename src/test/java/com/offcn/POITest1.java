package com.offcn;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class POITest1 {

	@SuppressWarnings("resource")
	@Test
	public void POItest1() throws IOException{
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("工作表1");
		HSSFRow row = sheet.createRow(4);
		HSSFCell cell = row.createCell(2);
		cell.setCellValue("Hello JAVA !");
		 File file = new File("d:\\hello.xls");
		 workbook.write(file);
		 System.out.println("Creating a successful");
	}
	
	@SuppressWarnings("resource")
	@Test
	public void getValue() throws IOException{
		FileInputStream in = new FileInputStream("d:\\hello.xls");
		HSSFWorkbook workbook = new HSSFWorkbook(in);
		HSSFSheet sheet = workbook.getSheet("工作表1");
		HSSFRow row = sheet.getRow(4);
		HSSFCell cell = row.getCell(2);
		System.out.println(cell.getStringCellValue());
	}
	
	@SuppressWarnings("resource")
	@Test
	public void poiByXSSF() throws IOException{
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("工作表2");
		XSSFRow row = sheet.createRow(2);
		XSSFRow row1 = sheet.createRow(0);
		XSSFCell cell = row.createCell(0);
		XSSFCell cell1 = row.createCell(1);
		XSSFCell cell2 = row.createCell(2);
		XSSFCell cell3 = row.createCell(3);
		cell.setCellValue("one");
		cell1.setCellValue("two");
		cell2.setCellValue("three");
		cell3.setCellValue("four");
		FileOutputStream os = new FileOutputStream("d:\\workBook.xlsx");
		workbook.write(os);
		System.out.println("创建成功");
	}
	
	@Test
	public void getValueByWorkbookFactory() throws EncryptedDocumentException, InvalidFormatException, IOException{
		FileInputStream in = new FileInputStream("d:\\workBook.xlsx");
		Workbook workbook = WorkbookFactory.create(in);
		Sheet sheet = workbook.getSheet("工作表2");
		Row row = sheet.getRow(2);
		Cell cell = row.getCell(0);
		Cell cell2 = row.getCell(1);
		System.out.println(cell.getStringCellValue()+"-----"+cell2.getStringCellValue());
	}
	
	@SuppressWarnings("resource")
	@Test
	public void ExcelReadByIteart() throws EncryptedDocumentException, InvalidFormatException, IOException{
		FileInputStream in = new FileInputStream("d:\\workBook.xlsx");
		Workbook workbook = WorkbookFactory.create(in);
		int sheetNum = workbook.getNumberOfSheets();
		for(int i=0; i<sheetNum; i++){
			Sheet sheet = workbook.getSheetAt(i);
			int rowNum = sheet.getPhysicalNumberOfRows();
			for(int j=0; j<rowNum; j++){
				Row row = sheet.getRow(j);
				int cellNum = row.getPhysicalNumberOfCells();
				for(int k=0; k<cellNum; k++){
					Cell cell = row.getCell(k);
					if(cell.getCellType() == HSSFCell.CELL_TYPE_STRING){
						System.out.println(cell.getStringCellValue()+"\t");
					}else if(cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC){
						System.out.println(cell.getNumericCellValue()+"\t");
					}else if(cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN){
						System.out.println(cell.getBooleanCellValue()+"\t");
					}else if(cell.getCellType() == HSSFCell.CELL_TYPE_BLANK){
						System.out.println("NULL"+"\t");
					}else{
						System.out.println(cell.getDateCellValue()+"\t");
					}
				}
			}
		}
	}
}
