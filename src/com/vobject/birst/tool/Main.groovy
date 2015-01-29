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

new ExcelBuilder("InputFile.xls").eachLine([sheet:"Birst", labels:true]) {
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


	matchSalesOrderNumbers(birstRecords)

	birstRecords.eachWithIndex { o, i ->
	 println "$i) $o"
	}

	def matchSalesOrderNumbers (List birstRecords) {
		List SSI_SALES_STAGES = ['Quote Request', 'Not Connected']
		Long salesOrderNumber
		List salesOrderNumbers = []
		
		new ExcelBuilder("InputFile.xls").eachLine([sheet:"CRM", labels:true]) {
			//println "$SSISalesStage ${SSI_SALES_STAGES.contains(SSISalesStage)}"
			if (SSI_SALES_STAGES.contains(SSISalesStage)) {
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