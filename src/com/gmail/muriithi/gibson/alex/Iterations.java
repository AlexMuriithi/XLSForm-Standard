package com.gmail.muriithi.gibson.alex;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import javax.swing.JFileChooser;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Reading excel data and displaying them on the console
 * 
 * @author AlexMuriithi (alex.gibson.muriithi@gmail.com)
 *
 */
public class Iterations {

	public static void main(String[] args) {
		JFileChooser fileChooser = new JFileChooser();
		int returnValue = fileChooser.showDialog(null, null);

		if (returnValue == JFileChooser.APPROVE_OPTION) {
			try {
				@SuppressWarnings("resource")
				Workbook workbook = new HSSFWorkbook(new FileInputStream(fileChooser.getSelectedFile()));
				Sheet sheet = workbook.getSheetAt(0);

				for (Row row : sheet) {
					for (Cell cell : row) {
						cell.setCellType(Cell.CELL_TYPE_STRING);
						System.out.print(cell.getStringCellValue() + "\t");
					}
					System.out.println();
				}
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

}
