package com.gmail.muriithi.gibson.alex;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderFormatting;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Conditional formatting
 * 
 * Excel enables you to highlight cells with a certain color, depending on the
 * cell's value.
 * 
 * @author Alex Muriithi (alex.gibson.muriithi@gmail.com)
 * 
 */
public class ConditionalFormating {

	@SuppressWarnings("resource")
	public static void main(String[] args) {

		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("ConditionalTest");

		Row row0 = sheet.createRow(0);
		row0.createCell(0).setCellValue(50); // A1
		row0.createCell(1).setCellValue("+"); // B1
		row0.createCell(2).setCellValue(20); // C1
		row0.createCell(3).setCellValue("="); // D1

		// Create conditional formats
		SheetConditionalFormatting conditionalFormatting = sheet.getSheetConditionalFormatting();

		// Create Rules
		ConditionalFormattingRule formattingRule = conditionalFormatting
				.createConditionalFormattingRule(ComparisonOperator.EQUAL, "$A1+$C1");

		// Change background color
		PatternFormatting background = formattingRule.createPatternFormatting();
		background.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.getIndex());

		// Change font
		FontFormatting font = formattingRule.createFontFormatting();
		font.setFontColorIndex(IndexedColors.DARK_GREEN.getIndex());

		// Change border
		BorderFormatting border = formattingRule.createBorderFormatting();
		border.setBorderBottom(BorderFormatting.BORDER_DOUBLE);
		border.setBottomBorderColor(IndexedColors.AQUA.getIndex());

		CellRangeAddress[] range = { CellRangeAddress.valueOf("E1:E1") };
		conditionalFormatting.addConditionalFormatting(range, formattingRule);

		try {
			FileOutputStream output = new FileOutputStream("ConditionalTest.xls");
			workbook.write(output);
			output.close();
			System.out.println("=== FILE CREATED ===");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
