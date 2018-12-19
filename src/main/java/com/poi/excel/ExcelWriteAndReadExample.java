package com.poi.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.poi.constants.AppConstants;
import com.poi.utilities.AppUtilities;

public class ExcelWriteAndReadExample {

	private static HashMap<String, CellStyle> excelCellStyleMap = new HashMap<>();

	public static void writeExcelSheet(String filePath) {

		// generate excel workbook - extension based on defined filePath
		Workbook wb = new XSSFWorkbook();
		if (AppUtilities.getExtension(filePath).equalsIgnoreCase("xls")) {
			wb = new HSSFWorkbook();
		}
		Sheet sheet = wb.createSheet("BatchConfigurations");
		excelCellStyleMap(sheet, wb);
		CellStyle style = excelCellStyleMap.get("style");
		Object[][] datatypes = { { "Datatype", "Type", "Size(in bytes)" }, { "int", "Primitive", 2 },
				{ "float", "Primitive", 4 }, { "double", "Primitive", 8 }, { "char", "Primitive", 1 },
				{ "String", "Non-Primitive", "No fixed size" } };

		int rowNum = 0;
		System.out.println("Creating excel");
		for (Object[] datatype : datatypes) {
			Row row = sheet.createRow(rowNum++);
			int colNum = 0;
			for (Object field : datatype) {
				Cell cell = row.createCell(colNum++);
				if (field instanceof String) {
					cell.setCellValue((String) field);
					cell.setCellStyle(style);
				} else if (field instanceof Integer) {
					cell.setCellValue((Integer) field);
				}
			}
		}

		try {
			File file = new File(filePath);
			if (file.exists()) {
				file.delete();
			}
			FileOutputStream outputStream = new FileOutputStream(filePath);
			wb.write(outputStream);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	//

	public static void main(String args[]) {
		ExcelWriteAndReadExample.writeExcelSheet(AppConstants.APP_TEXT_EXCEL_FILE);
	}

	/**
	 * expandAllExpandCollapseColumn
	 * 
	 * @param sheet
	 * @param firstColumn
	 * @param lastColumn
	 */
	public void expandAllExpandCollapseColumn(Sheet sheet, int firstColumn, int lastColumn) {
		sheet.groupColumn(firstColumn, lastColumn);
		sheet.setColumnGroupCollapsed(firstColumn, true);
	}

	/**
	 * Add Styles
	 * 
	 * @param sheet
	 * @param oWorkbook
	 */
	public static void excelCellStyleMap(Sheet sheet, Workbook oWorkbook) {
		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 8);
		font.setColor(HSSFFont.COLOR_RED);

		CellStyle style;
		style = sheet.getWorkbook().createCellStyle();
		style.setFont(font);
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setWrapText(true);
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		excelCellStyleMap.put("style", style);
	}
}
