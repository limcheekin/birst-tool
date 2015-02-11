package com.vobject.birst.tool

class Constants {
	// Used for Excel sheet generation
	final static WebSafeColors SALES_ORDER_NUMBER_MATCHED_COLOR = WebSafeColors._FFFF00 // Yellow
	final static WebSafeColors SERIAL_NUMBER_MATCHED_COLOR = WebSafeColors._0099FF // Blue
	final static WebSafeColors PART_ID_NOT_EXISTS_COLOR = WebSafeColors._FF3300 // Red
	final static List DUPLICATE_SERIAL_NUMBER_COLORS = WebSafeColors.values()
															 .minus(SALES_ORDER_NUMBER_MATCHED_COLOR,
																	SERIAL_NUMBER_MATCHED_COLOR,
																	PART_ID_NOT_EXISTS_COLOR)
															 // .findAll { it.brightness > 130 }
	final static List OUTPUT_COLUMNS = ['duplicate',
										'serialNumber',
										'PO_Number',
										'salesOrderNumber',
										'partId',
										'partDescription',
										'originalShipDate',
										'startDate',
										'endDate',
										'contractSapId',
										'reseller',
										'endUser',
										'endUserStandardName',
										'endUserState',
										'soldTo',
										'billTo',
										'shipTo',
										'type',
										'warrantyType',
										'entitlementId']			
	
	final static String EMPTY_STRING = ""
	
	// REF: http://stackoverflow.com/questions/18983203/how-to-speed-up-autosizing-columns-in-apache-poi
	final static Integer SINGLE_CHARACTER_SIZE = 1.14388 * 256 as Integer
	final static Integer STANDARD_MAX_SIZE = 20
	final static Integer DATE_MAX_SIZE = 12
}
