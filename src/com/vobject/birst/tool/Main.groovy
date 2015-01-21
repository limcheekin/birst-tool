/**
 * 
 */
package com.vobject.birst.tool

import static org.apache.poi.ss.usermodel.CellStyle.*
import static org.apache.poi.ss.usermodel.IndexedColors.*
import org.apache.poi.ss.usermodel.IndexedColors

/**
 * @author limcheek
 *
 */

// Used for Excel sheet generation
IndexedColors SALES_ORDER_NUMBER_MATCHED_COLOR = YELLOW
IndexedColors SERIAL_NUMBER_MATCHED_COLOR = BLUE
IndexedColors PART_ID_NOT_EXISTS_COLOR = RED
def DUPLICATE_SERIAL_NUMBER_COLORS = IndexedColors.values().minus(SALES_ORDER_NUMBER_MATCHED_COLOR, SERIAL_NUMBER_MATCHED_COLOR, PART_ID_NOT_EXISTS_COLOR)

// Random number generation
// URL: http://groovy-almanac.org/create-random-integers-in-a-specific-range/
Random rand = new Random()

def birstRecords = []

new ExcelBuilder("BirstFile.xls").eachLine([labels:true]) {
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

birstRecords.eachWithIndex { o, i -> 
	println "$i) $o"
}
