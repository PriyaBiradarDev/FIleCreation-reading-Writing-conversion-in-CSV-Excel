package com.howtodoinjava.demo.poi;

import java.io.DataInputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class CsvToXls {

	public static void main(String args[]) throws IOException {
		ArrayList arList = null;
		ArrayList al = null;
		String fName = "SampleCSVFile_2kb.csv";
		String thisLine;
		int count = 0;
		FileInputStream fis = new FileInputStream(fName);
		DataInputStream myInput = new DataInputStream(fis);
		int i = 0;
		arList = new ArrayList();
		while ((thisLine = myInput.readLine()) != null) {
			al = new ArrayList();
			String strar[] = thisLine.split(",");
			for (int j = 0; j < strar.length; j++) {
				al.add(strar[j]);
			}
			arList.add(al);
			System.out.println();
			i++;
		}

		try {
			HSSFWorkbook hwb = new HSSFWorkbook();
			HSSFSheet sheet = hwb.createSheet("new sheet");
			for (int k = 0; k < arList.size(); k++) {
				ArrayList ardata = (ArrayList) arList.get(k);
				System.out.println("ardata " + ardata.size());
				HSSFRow row = sheet.createRow((short) 0 + k);
				for (int p = 0; p < ardata.size(); p++) {
					System.out.print(ardata.get(p));
					HSSFCell cell = row.createCell((short) p);
					cell.setCellValue(ardata.get(p).toString());
				}
				System.out.println();
			}
			FileOutputStream fileOut = new FileOutputStream("xhhh.xls");
			hwb.write(fileOut);
			fileOut.close();
			System.out.println("Your excel file has been generated");
		} catch (Exception ex) {
		} // main method ends
	}
}
