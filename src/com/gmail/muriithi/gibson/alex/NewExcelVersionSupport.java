package com.gmail.muriithi.gibson.alex;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * Newer versions of excel documents
 * 
 * @author AlexMuriithi (alex.gibson.muriithi@gmail.com)
 *
 */
public class NewExcelVersionSupport {

	@SuppressWarnings("resource")
	public static void main(String[] args) {

		// Instead of using the classes from the HSSF package, use the classes
		// from the XSSF package (eg. XSSFWorkbook)
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("NewExcelVersionSupport");

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
			// added the "x" to the suffix of the document ("document.xlsx")
			FileOutputStream output = new FileOutputStream("NewExcelVersionSupport.xlsx");
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
