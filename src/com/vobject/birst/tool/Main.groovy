/**
 * 
 */
package com.vobject.birst.tool

import static org.apache.poi.ss.usermodel.CellStyle.*

/**
 * @author limcheek
 *
 */

// Used for Excel sheet generation
final WebSafeColors SALES_ORDER_NUMBER_MATCHED_COLOR = WebSafeColors._FFFF00 // Yellow
final WebSafeColors SERIAL_NUMBER_MATCHED_COLOR = WebSafeColors._0099FF // Blue
final WebSafeColors PART_ID_NOT_EXISTS_COLOR = WebSafeColors._FF3300 // Red
final List DUPLICATE_SERIAL_NUMBER_COLORS = WebSafeColors.values()
                                                         .minus(SALES_ORDER_NUMBER_MATCHED_COLOR, 
																SERIAL_NUMBER_MATCHED_COLOR, 
																PART_ID_NOT_EXISTS_COLOR)
														  .collect { it.brightness > 130 }

// Random number generation
// REF: http://groovy-almanac.org/create-random-integers-in-a-specific-range/
Random rand = new Random()

def birstRecords = []

new ExcelBuilder("InputFile.xls").eachLine([sheet:"Birst Excel File", labels:true]) {
	//systemSerialNumber, certificateSerialNumber, PONumber, salesOrderNum, partID, partDescription, originalShipDate, startDate, endDate, contractSAPID, reseller, endUser, endUserStandardName, endUserState, soldTo, billTo, shipTo, type, warrantyType, entitlementId
	birstRecords << new BirstRecord (
	  serialNumber: systemSerialNumber ? systemSerialNumber.trim() : certificateSerialNumber.trim(),
	  purchaseOrderNumber: PONumber.trim(),
	  salesOrderNumber: salesOrderNum as Long,
	  partId: partID.trim(),
	  partDescription: partDescription.trim(),
	  originalShipDate: originalShipDate ? originalShipDate as Date : null,
	  startDate: startDate ? startDate as Date : null,
	  endtDate: endtDate ? endtDate as Date : null,
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

birstRecords.eachWithIndex { o, i -> 
	println "$i) $o"
}
