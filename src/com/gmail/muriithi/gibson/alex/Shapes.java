package com.gmail.muriithi.gibson.alex;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Excel Shapes
 * 
 * Lines, Ovals, Rectangular, those are all words... And shapes...
 * 
 * @author AlexMuriithi (alex.gibson.muriithi@gmail.com)
 *
 */
public class Shapes {

	@SuppressWarnings("resource")
	public static void main(String[] args) {

		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("Shapes");

		HSSFPatriarch patriarch = (HSSFPatriarch) sheet.createDrawingPatriarch();
		HSSFClientAnchor anchor = new HSSFClientAnchor();

		// Lines
		anchor.setCol1(1);
		anchor.setRow1(1);
		anchor.setCol2(3);
		anchor.setRow2(5);


		HSSFSimpleShape shape = patriarch.createSimpleShape(anchor);
		shape.setShapeType(HSSFSimpleShape.OBJECT_TYPE_LINE);
		shape.setFillColor(236, 157, 245);
		shape.setLineStyle(HSSFSimpleShape.LINESTYLE_DASHGEL);
		shape.setLineStyleColor(128, 255, 0);
		shape.setLineWidth(HSSFSimpleShape.LINEWIDTH_ONE_PT * 3);
		
		// Oval
		HSSFClientAnchor anchor2 = new HSSFClientAnchor();
		anchor2.setCol1(4);
		anchor2.setRow1(4);
		anchor2.setCol2(7);
		anchor2.setRow2(9);

		HSSFSimpleShape shape2 = patriarch.createSimpleShape(anchor2);
		shape2.setShapeType(HSSFSimpleShape.OBJECT_TYPE_OVAL);
		shape2.setFillColor(236, 157, 245);
		shape2.setLineStyle(HSSFSimpleShape.LINESTYLE_DASHGEL);
		shape2.setLineStyleColor(128, 255, 0);
		shape2.setLineWidth(HSSFSimpleShape.LINEWIDTH_ONE_PT * 3);
		
		// Rectangle
		HSSFClientAnchor anchor3 = new HSSFClientAnchor();
		anchor3.setCol1(8);
		anchor3.setRow1(8);
		anchor3.setCol2(11);
		anchor3.setRow2(13);

		HSSFSimpleShape shape3 = patriarch.createSimpleShape(anchor3);
		shape3.setShapeType(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
		shape3.setFillColor(236, 157, 245);
		shape3.setLineStyle(HSSFSimpleShape.LINESTYLE_DASHGEL);
		shape3.setLineStyleColor(128, 255, 0);
		shape3.setLineWidth(HSSFSimpleShape.LINEWIDTH_ONE_PT * 3);
		
		try {
			FileOutputStream output = new FileOutputStream("Shapes.xls");
			workbook.write(output);
			output.close();
			System.out.println("=== File Created ===");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
