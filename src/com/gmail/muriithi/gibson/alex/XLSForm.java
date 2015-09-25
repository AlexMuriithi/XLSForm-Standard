package com.gmail.muriithi.gibson.alex;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * XLSForm Template from Specifications defined at XLSForm org
 * 
 * @see <a href = "http://xlsform.org/"> XLSForm.org </a>
 * 
 * @author Alex Muriithi (alex.gibson.muriithi@gmail.com)
 */

public class XLSForm {

	Workbook workbook = new HSSFWorkbook();

	private static final String INSTRUCTION = "instruction";

	public XLSForm() {
		
		Sheet FormulaeTest = workbook.createSheet(INSTRUCTION);
		
		Cell cell1 = FormulaeTest.createRow(0).createCell(0);
		Cell cell2 = FormulaeTest.createRow(0).createCell(1);
		Cell cell3 = FormulaeTest.createRow(0).createCell(2);
		Cell cell4 = FormulaeTest.createRow(0).createCell(3);
		Cell cell5 = FormulaeTest.createRow(0).createCell(4);
		
		cell1.setCellValue(100);
		cell2.setCellValue("+");
		cell3.setCellValue(200);
		cell4.setCellValue("=");
		cell5.setCellFormula("A1+C1");
		
		Cell cell6 = FormulaeTest.createRow(1).createCell(0);
		Cell cell7 = FormulaeTest.createRow(1).createCell(1);
		Cell cell8 = FormulaeTest.createRow(1).createCell(2);
		Cell cell9 = FormulaeTest.createRow(1).createCell(3);
		Cell cell0 = FormulaeTest.createRow(1).createCell(4);
		
		cell6.setCellValue(100);
		cell7.setCellValue(200);
		cell8.setCellValue(300);
		cell9.setCellValue(400);
		cell0.setCellFormula("SUM(A2:D2)");
		
		
		try {
			FileOutputStream output = new FileOutputStream("XLSTrail.xls");
			workbook.write(output);
			workbook.close();
			System.out.println("=== File Created ===");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public static void main(String[] args) {
		new XLSForm();
	}
}
