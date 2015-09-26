package com.gmail.muriithi.gibson.alex;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
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
		FormulaeTest.setColumnWidth(0, 7000);		
		
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setBorderBottom(CellStyle.BORDER_THICK);
		style.setBorderBottom(IndexedColors.GREEN.getIndex());

		Font font = workbook.createFont();
		font.setColor(IndexedColors.YELLOW.getIndex());
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		font.setItalic(true);
		font.setFontHeightInPoints((short) 16);
		font.setUnderline(Font.U_DOUBLE);
		font.setFontName("Trebuchet MS");

		style.setFont(font);
		
		Cell instructionTitle = FormulaeTest.createRow(0).createCell(0);
		instructionTitle.setCellValue(INSTRUCTION);
		instructionTitle.setCellStyle(style);
		instructionTitle.getRow().setHeightInPoints(45);
				
		Cell cell1 = FormulaeTest.createRow(1).createCell(0);
		Cell cell2 = FormulaeTest.createRow(1).createCell(1);
		Cell cell3 = FormulaeTest.createRow(1).createCell(2);
		Cell cell4 = FormulaeTest.createRow(1).createCell(3);
		Cell cell5 = FormulaeTest.createRow(1).createCell(4);

		cell1.setCellValue(100);
		cell2.setCellValue("+");
		cell3.setCellValue(200);
		cell4.setCellValue("=");
		cell5.setCellFormula("A2+C2");
		cell5.setCellStyle(style);

		Cell cell6 = FormulaeTest.createRow(2).createCell(0);
		Cell cell7 = FormulaeTest.createRow(2).createCell(1);
		Cell cell8 = FormulaeTest.createRow(2).createCell(2);
		Cell cell9 = FormulaeTest.createRow(2).createCell(3);
		Cell cell0 = FormulaeTest.createRow(2).createCell(4);

		cell6.setCellValue(100);
		cell7.setCellValue(200);
		cell8.setCellValue(300);
		cell9.setCellValue(400);
		cell0.setCellFormula("SUM(A3:D3)");
		cell0.setCellStyle(style);

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
