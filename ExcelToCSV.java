package com.howtodoinjava.demo.poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

// Null Point Exception
// File to be written as CSV
// Formula Evaluation
public class ExcelToCSV {

	public static void echoAsCSV(Sheet sheet) {
		Row row = null;
		for (int i = 0; i <= sheet.getLastRowNum(); i++) {
			row = sheet.getRow(i);
			for (int j = 0; j < row.getLastCellNum(); j++) {
				System.out.print("\"" + row.getCell(j) + "\";");
			}
			System.out.println();
		}
	}

	/**
	 * @param args the command line arguments
	 */
	public static void main(String[] args) {
		InputStream inp = null;
		try {
			//try with resource
			inp = new FileInputStream("test_xlsx.xlsx");
			Workbook wb = WorkbookFactory.create(inp);

			for (int i = 0; i < wb.getNumberOfSheets(); i++) {
				System.out.println(wb.getSheetAt(i).getSheetName());
				echoAsCSV(wb.getSheetAt(i));
			}
		} catch (InvalidFormatException ex) {
			// Logger.getLogger(ExcelReading.class.getName()).log(Level.SEVERE, null, ex);
		} catch (FileNotFoundException ex) {
			// Logger.getLogger(ExcelReading.class.getName()).log(Level.SEVERE, null, ex);
		} catch (IOException ex) {
			// Logger.getLogger(ExcelReading.class.getName()).log(Level.SEVERE, null, ex);
		} finally {
			try {
				inp.close();
			} catch (IOException ex) {
				// Logger.getLogger(ExcelReading.class.getName()).log(Level.SEVERE, null, ex);
			}
		}
	}
}
