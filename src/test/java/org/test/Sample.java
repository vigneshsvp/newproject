package org.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.base.CaseFormat;

public class Sample {
	public static void main(String[] args) throws IOException {
		File file = new File("C:\\Users\\USER\\eclipse-workspace\\Framework1\\Excel\\Book12.xlsx");
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet("Sheet1");
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
	
				CellType type= cell.getCellType();
				switch (type) {
				case STRING:
					String Value = cell.getStringCellValue();
					System.out.println(Value);
					break;
				case NUMERIC:
					double numericCellValue = cell.getNumericCellValue();
					
					BigDecimal b = BigDecimal.valueOf(numericCellValue);
					b.toString();
					System.out.println(b);
					break;

				default:
					break;
				}
		
				}
			}
		}		
		}
		


