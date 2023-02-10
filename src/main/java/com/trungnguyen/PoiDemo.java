package com.trungnguyen;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class PoiDemo {

	public static void main(String[] args) {
		try (InputStream inputStream = new FileInputStream("C:\\Users\\thang\\Desktop\\import_to_chuc.xlsx")) {
			var wb = WorkbookFactory.create(inputStream);
			var sheet = wb.getSheetAt(0);
			var row = sheet.getRow(1);
			var style = wb.createCellStyle();
			var format = wb.createDataFormat();
			style.setDataFormat(format.getFormat("@"));
			var cell = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
			cell.setCellValue("012");
			cell.setCellStyle(style);
			
			
			try (OutputStream fileOut = new FileOutputStream("workbook.xlsx")) {
		        wb.write(fileOut);
		    }
			
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
