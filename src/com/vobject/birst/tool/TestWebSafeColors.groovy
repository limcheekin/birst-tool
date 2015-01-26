package com.vobject.birst.tool

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import org.apache.poi.xssf.usermodel.XSSFWorkbook

Workbook wb = new XSSFWorkbook() //or new HSSFWorkbook()
Sheet sheet = wb.createSheet("Web Safe Colors")

// Create a row and put some cells in it. Rows are 0 based.
Row row = sheet.createRow((short) 1)

WebSafeColors.values().eachWithIndex { o, i ->
	CellStyle style = wb.createCellStyle()
	style.setFillForegroundColor(new XSSFColor(o.value()))
	style.setFillPattern(CellStyle.SOLID_FOREGROUND)
	Cell cell = row.createCell(i)
	cell.setCellValue(new XSSFRichTextString(o.name()))
	cell.setCellStyle(style)
}

int j = 3;
WebSafeColors.values().eachWithIndex { o, i ->
	if (o.brightness > 130) {
		row = sheet.createRow(j++);
		CellStyle style = wb.createCellStyle()
		style.setFillForegroundColor(new XSSFColor(o.value()))
		style.setFillPattern(CellStyle.SOLID_FOREGROUND)
		Cell cell = row.createCell(1)
		cell.setCellValue(new XSSFRichTextString(o.name()))
		cell.setCellStyle(style)
	}
}


// Write the output to a file
FileOutputStream fileOut = new FileOutputStream("fill_colors.xlsx")
wb.write(fileOut)
fileOut.close()

