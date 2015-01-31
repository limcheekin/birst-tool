/**
 * 
 */
package com.vobject.birst.tool

import groovy.transform.ToString
import groovy.transform.Sortable
/**
 * @author limcheek
 *
 */

// REF: http://mrhaki.blogspot.com/2014/05/groovy-goodness-use-sortable-annotation.html
@Sortable(includes = ['salesOrderNumber', 'serialNumber'])
@ToString
class BirstRecord {
	String serialNumber
	String purchaseOrderNumber
	Long salesOrderNumber
	String partId
	String partDescription
	Date originalShipDate
	Date startDate
	Date endDate
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
	Boolean isSalesOrderNumberFound
	Boolean isSerialNumberFound
	Boolean isPartIdFound
	WebSafeColors duplicateSerialNumberColor
	
}
