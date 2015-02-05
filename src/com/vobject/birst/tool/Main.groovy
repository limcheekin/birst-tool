/**
 * 
 */
package com.vobject.birst.tool

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * @author limcheek
 *
 */

def birstRecords = []

new ExcelBuilder("InputFile.xlsx").eachLine([sheet:"Birst", labels:true]) {
	//systemSerialNumber, certificateSerialNumber, PONumber, salesOrderNum, partID, partDescription, originalShipDate, startDate, endDate, contractSAPID, reseller, endUser, endUserStandardName, endUserState, soldTo, billTo, shipTo, type, warrantyType, entitlementId
	birstRecords << new BirstRecord (
	  serialNumber: systemSerialNumber ? systemSerialNumber.trim() : certificateSerialNumber.trim(),
	  purchaseOrderNumber: PONumber.trim(),
	  salesOrderNumber: salesOrderNum as Long,
	  partId: partID.trim(),
	  partDescription: partDescription.trim(),
	  originalShipDate: originalShipDate ? originalShipDate as Date : null,
	  startDate: startDate ? startDate as Date : null,
	  endDate: endDate ? endDate as Date : null,
	  contractSapId: contractSAPID.trim(),
	  reseller: reseller.trim(),
	  endUser: endUser.trim(),
	  endUserStandardName: endUserStandardName.trim(),
	  endUserState: endUserState.trim(),
	  soldTo: soldTo.trim(),
	  billTo: billTo.trim(),
	  shipTo: shipTo.trim(),
	  type: type.trim(),
	  warrantyType: warrantyType.trim(),
	  entitlementId: entitlementId.trim()
	)
	
}

	birstRecords.sort()

	matchSalesOrderNumbers(birstRecords)
	matchSerialNumbers(birstRecords)
	matchPartIDs(birstRecords)
	checkDuplicateSerialNumbers(birstRecords)
	generateOutputFile(birstRecords)

	/*birstRecords.eachWithIndex { o, i ->
	 println "$i) $o"
	}*/

	def matchSalesOrderNumbers (List birstRecords) {
		List SSI_SALES_STAGES = ['Quote Request', 'Not Connected']
		Long salesOrderNumber
		List salesOrderNumbers = []
		
		new ExcelBuilder("InputFile.xlsx").eachLine([sheet:"CRM", labels:true]) {
			//println "$SSISalesStage ${SSI_SALES_STAGES.contains(SSISalesStage)}"
			if (SSI_SALES_STAGES.contains(SSISalesStage)) {
				// REF: http://stackoverflow.com/questions/15572481/extract-numeric-data-from-string-in-groovy
				for (number in opportunityID.findAll( /\d+/ )) {
					if (number.length() == 7) {
						salesOrderNumber = number as Long
						if (!salesOrderNumbers.contains(salesOrderNumber))
							salesOrderNumbers << salesOrderNumber
						break
					}
				}
			}
		}
		
		// println salesOrderNumbers
		
		salesOrderNumbers.each { salesOrderNum ->
			birstRecords.each {
				if (it.salesOrderNumber == salesOrderNum) it.isSalesOrderNumberFound = true
			}
		}
	}
	
	def matchSerialNumbers (List birstRecords) {
		String serialNum
		new ExcelBuilder("InputFile.xlsx").eachLine([sheet:"CRM Booking Package", labels:true]) {
			serialNum = serialNumber.trim()
			birstRecords.each {
				if (it.serialNumber == serialNum) it.isSerialNumberFound = true
			}
		}
	}
	
	def matchPartIDs (List birstRecords) {
		List partIDs = []
		
		new ExcelBuilder("InputFile.xlsx").eachLine([sheet:"Price List", labels:true]) {
			if (arubaCareSKU.trim()) {
				partIDs << arubaCareSKU.trim()
			}
		}
		
		birstRecords.each {
			for (partId in partIDs) {
				if (it.partId == partId) {
					it.isPartIdFound = true
					break
				}
			}
		}
	}
	
	def checkDuplicateSerialNumbers (List birstRecords) {
		WebSafeColors duplicateColor
		Integer size = birstRecords.size()
		int k = 0
		// Random number generation
		// REF: http://groovy-almanac.org/create-random-integers-in-a-specific-range/
		Random rand = new Random()
		int max = Constants.DUPLICATE_SERIAL_NUMBER_COLORS.size()
		String lastSerialNumber
		
		for (int i = 0; i < size; i++) {
			for (int j = i + 1; j < size; j++) {
				if (birstRecords[i].serialNumber && birstRecords[j].serialNumber) { 
					println "${++k} i) $i ${birstRecords[i].serialNumber}, j) $j ${birstRecords[j].serialNumber}"
					if (birstRecords[i].serialNumber == birstRecords[j].serialNumber) {
						println "DUPLICATE i) $i ${birstRecords[i].salesOrderNumber} ${birstRecords[i].serialNumber}, j) $j ${birstRecords[j].salesOrderNumber} ${birstRecords[j].serialNumber}"
						if (lastSerialNumber != birstRecords[i].serialNumber) {
							lastSerialNumber = birstRecords[i].serialNumber
							duplicateColor = Constants.DUPLICATE_SERIAL_NUMBER_COLORS[rand.nextInt(max)]
							birstRecords[i].duplicateSerialNumberColor = duplicateColor
						}		
						birstRecords[j].duplicateSerialNumberColor = duplicateColor
					}
				}
			}
		}
	}
	
	def generateOutputFile (List birstRecords) {
		XSSFCell cell
		XSSFRow row
		XSSFWorkbook wb = new XSSFWorkbook()
		XSSFSheet sheet = wb.createSheet("Birst")
		XSSFRichTextString richText
		
		Font font = wb.createFont()
		font.bold = true
		// Lock header row
		// REF: http://stackoverflow.com/questions/17932575/apache-poi-locking-header-rows
		sheet.createFreezePane(0, 1);
		if (birstRecords) {
			CellStyle salesOrderNumberMatchedStyle = wb.createCellStyle()
			salesOrderNumberMatchedStyle.setFillForegroundColor(new XSSFColor(Constants.SALES_ORDER_NUMBER_MATCHED_COLOR.value()))
			salesOrderNumberMatchedStyle.setFillPattern(CellStyle.SOLID_FOREGROUND)
			CellStyle serialNumberMatchedStyle = wb.createCellStyle()
			serialNumberMatchedStyle.setFillForegroundColor(new XSSFColor(Constants.SERIAL_NUMBER_MATCHED_COLOR.value()))
			serialNumberMatchedStyle.setFillPattern(CellStyle.SOLID_FOREGROUND)
			CellStyle partIdNotExistsStyle = wb.createCellStyle()
			partIdNotExistsStyle.setFillForegroundColor(new XSSFColor(Constants.PART_ID_NOT_EXISTS_COLOR.value()))
			partIdNotExistsStyle.setFillPattern(CellStyle.SOLID_FOREGROUND)
			CellStyle duplicateSerialNumberStyle
			
			// REF: http://stackoverflow.com/questions/5794659/poi-how-do-i-set-cell-value-to-date-and-apply-default-excel-date-format
			XSSFCellStyle dateCellStyle = wb.createCellStyle()
			dateCellStyle.dataFormat = wb.getCreationHelper().createDataFormat().getFormat("m/d/yy")
			
			// Create a row and put some cells in it. Rows are 0 based.
			XSSFRow header = sheet.createRow(0)
			Constants.OUTPUT_COLUMNS.eachWithIndex { name, i ->
				cell = header.createCell(i)
				richText = new XSSFRichTextString(NameUtils.getNaturalName(name))
				richText.applyFont(font)
				cell.setCellValue(richText)
			}
			
			birstRecords.eachWithIndex { birstRecord, i ->
				row  = sheet.createRow(i + 1)
				birstRecord.with {
					cell = row.createCell(0)
					if (duplicateSerialNumberColor) {
						duplicateSerialNumberStyle = wb.createCellStyle()
						duplicateSerialNumberStyle.setFillForegroundColor(new XSSFColor(duplicateSerialNumberColor.value()))
						duplicateSerialNumberStyle.setFillPattern(CellStyle.SOLID_FOREGROUND)
						cell.setCellStyle(duplicateSerialNumberStyle)
					}
					
					cell = row.createCell(1)
					cell.setCellValue(new XSSFRichTextString(serialNumber))
					if (isSerialNumberFound) {
						cell.setCellStyle(serialNumberMatchedStyle)
					}

					cell = row.createCell(2)
					cell.setCellValue(new XSSFRichTextString(purchaseOrderNumber))
					
					cell = row.createCell(3)
					cell.setCellValue(new XSSFRichTextString(salesOrderNumber as String))
					if (isSalesOrderNumberFound) {
						cell.setCellStyle(salesOrderNumberMatchedStyle)
					}
					
					cell = row.createCell(4)
					cell.setCellValue(new XSSFRichTextString(partId))
					if (!isPartIdFound) {
						cell.setCellStyle(partIdNotExistsStyle)
					}

					cell = row.createCell(5)
					cell.setCellValue(new XSSFRichTextString(partDescription))

					cell = row.createCell(6)
					cell.setCellValue(originalShipDate)
					cell.cellStyle = dateCellStyle
					
					cell = row.createCell(7)
					cell.setCellValue(startDate)
					cell.cellStyle = dateCellStyle

					cell = row.createCell(8)
					cell.setCellValue(endDate)
					cell.cellStyle = dateCellStyle
					
					cell = row.createCell(9)
					cell.setCellValue(new XSSFRichTextString(contractSapId))

					cell = row.createCell(10)
					cell.setCellValue(new XSSFRichTextString(reseller))
					
					cell = row.createCell(11)
					cell.setCellValue(new XSSFRichTextString(endUser))
					
					cell = row.createCell(12)
					cell.setCellValue(new XSSFRichTextString(endUserStandardName))

					cell = row.createCell(13)
					cell.setCellValue(new XSSFRichTextString(endUserState))
					
					cell = row.createCell(14)
					cell.setCellValue(new XSSFRichTextString(soldTo))
					
					cell = row.createCell(15)
					cell.setCellValue(new XSSFRichTextString(billTo))
					
					cell = row.createCell(16)
					cell.setCellValue(new XSSFRichTextString(shipTo))
					
					cell = row.createCell(17)
					cell.setCellValue(new XSSFRichTextString(type))
					
					cell = row.createCell(18)
					cell.setCellValue(new XSSFRichTextString(warrantyType))
					
					cell = row.createCell(19)
					cell.setCellValue(new XSSFRichTextString(entitlementId))
				}
				// REF: http://stackoverflow.com/questions/20190317/apache-poi-excel-big-auto-column-width
				for(int colNum = 0; colNum<row.getLastCellNum();colNum++)
					wb.getSheetAt(0).autoSizeColumn(colNum);
			}
		}
		
		FileOutputStream fileOut = new FileOutputStream("Output.xlsx")
		wb.write(fileOut)
		fileOut.close()
	}