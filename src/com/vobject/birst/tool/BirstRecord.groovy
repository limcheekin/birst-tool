/**
 * 
 */
package com.vobject.birst.tool

import groovy.transform.ToString
/**
 * @author limcheek
 *
 */

@ToString
class BirstRecord {
	String serialNumber
	String purchaseOrderNumber
	Long salesOrderNumber
	String partId
	String partDescription
	Date originalShipDate
	Date startDate
	Date endtDate
	String contractSapId
	String reseller
	String endUser
	String endUserStandardName
	String endUserState
	String soldTo
	String billTo
	String shipTo
	String type
	String warrantyType
	String entitlementId
}
