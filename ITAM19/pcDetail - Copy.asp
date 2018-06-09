<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit%>
<%

if session("UID")="" then
    Response.Redirect "Logon.asp"
    Response.End
else if Session("SecAdmin")=0 and Session("SecHP")=0 and Session("SecSLDE")=0 then
    Response.Redirect "Logon.asp"
    Response.End
   end if
end if

%>
<%

'-----------------------------------------------------------------------------
' Declaration of variables
'-----------------------------------------------------------------------------
Dim PageTitle
Dim MenuSize

Dim strInternalTag
Dim strComputerName
Dim strSerialNo, strSN, strModel
Dim strAssetType
Dim strAccount
Dim strTmp
Dim strInTag
Dim strColor
Dim intCount, intID
Dim strSQL, strDate, strRowTitle, strDebug, strStyle
Dim strSource
Dim dtInstall

PageTitle = "Asset details"
MenuSize="Large"%>

<!--#include file="Include\pageHeader.asp"-->
<!--#include file="Include\dbFunction.asp"-->
<!--#include file="Include\pcFunction.asp"-->

<table align="center" cols="4" width="970">
<%
'-----------------------------------------------------------------------------
' This page does display appropriate detail for a pc from the SQL database
' Version History
' 22 September 2003
' 16/11/2004 AMN Added Country of location and modified lay-out to center table
'-----------------------------------------------------------------------------


'Declaring LDAP objects
'Dim objCmd, objConnAD, objRootDSE, objRsAD, objComputer
'Declaring LDAP variables
'Dim strBase, strFilter, strAttributes, strDNSDomain, strQuery, strName, strCN

Call acOpenDB()

strInternalTag = Trim(Request.QueryString("frmInternalTag"))

strSQL = "SELECT A.*, CASE WHEN O.[OpCoID] IS NULL THEN 'Invalid OpCo' ELSE CASE WHEN O.[TACSCode] IS NULL THEN CASE O.[Display] WHEN '9000' " & _
	"THEN 'HP administrative OpCo, can not be invoiced' WHEN '9999' THEN 'US transition OpCo, can not be invoiced' WHEN '9998' " & _
	"THEN 'Canadian transition OpCo, can not be invoiced' ELSE 'OpCo can not be invoiced' END ELSE CASE WHEN " & _
	"O.[InvoiceStatus] & (64 + 128) = (64 + 128) THEN 'OpCo is fully managed' ELSE 'OpCo is not fully managed' END END END AS 'OpCoInfo', " & _
	"CASE WHEN [CategoryName] IN ('Desktop computer', 'Laptop', 'Netbook' , 'Thin client', 'Folio Laptop', 'Tablet') THEN 'Computer' ELSE 'Other' END AS 'AssetType' " & _
  "FROM RadiaRIMProd.dbo.acAssetList AS A LEFT OUTER JOIN RadiaRIMProd.dbo.udHRDataOpCo AS O ON A.[AssetOpCoID] = O.[OpCoID] " & _
  "WHERE [InternalTag] = '" &_ 
  strInternalTag & "'"

objRs.Open strSQL, objConn

If objRs("ComputerName") <> "" Then
	strComputerName = Replace(Replace(Replace(UCase(objRs("ComputerName")), "-RAD", ""), "-DUPL", ""), "-DUP", "")
Else
	strComputerName = "n/a"
End If
strSerialNo = objRs("SerialNo")
If strSerialNo = "" Then strSerialNo = "n/a"
strAssetType = objRs("AssetType")
strAccount = objRs("SupervisorUserLogin")

'----------------------------------------------------------------------------
' Display basic asset details from asset table
'----------------------------------------------------------------------------
Response.Write "<tr><td height=""15"" class=""menutitle"" colspan=4>&nbsp;</td></tr>"
Response.Write "<tr><td width=""30%""/><td width=""25%""/><td/></tr>"

If objRs("ComputerName") = strComputerName Then
	Response.Write "<tr><td colspan=4><b>Details for asset with asset name " & objRs("ComputerName") & " and serial number " & objRs("SerialNo")& "</b></td></tr>"
Else
	If strComputerName = "n/a" Then
		Response.Write "<tr><td colspan=4><b>Details for asset with serial number " & objRs("SerialNo")& "</b></td></tr>"
	Else
		Response.Write "<tr><td colspan=4><b>Details for asset with asset name " & objRs("ComputerName") & "//" & strComputerName & " and serial number " & objRs("SerialNo")& "</b></td></tr>"
	End If
End If

dtInstall = objRs("DTInstall")

Response.Write "<tr><td>Main-user last name</td><td colspan=""3"">" & objRs("SupervisorName")& "</td></tr>"
Response.Write "<tr><td>Main-user first name</td><td colspan=""3"">" & objRs("SupervisorFirstName")& "</td></tr>"
Response.Write "<tr><td>Main-user e-mail</td><td colspan=""3"">" & objRs("SupervisorEMail")& "</td></tr>"

Response.Write "<tr><td colspan=""4""><hr/></td></tr>"

Response.Write "<tr><td>Phone</td><td>" & objRs("SupervisorPhone") & "</td></tr>"
Response.Write "<tr><td>OpCo</td><td colspan=""3"">" & objRs("fv_SLDE_BUL") & " " & objRs("fv_SLDE_OpCo") & " (" & objRs("OpCoInfo") & ")" & "</td></tr>"
Response.write "<tr><td>Country</td><td>" & objRs("LocationCountry") & "</td></tr>"
Response.Write "<tr><td>Location</td><td>" & objRs("LocationName") & "</td></tr>"
If objRs("fv_SLDE_LocDetail") <> "" Then
	Response.Write "<tr><td>Location detail</td><td>" & objRs("fv_SLDE_LocDetail") & "</td></tr>"
End If
Response.Write "<tr><td colspan=""4""><hr/></td></tr>"
Response.Write "<tr><td>Computer name</td><td>" & objRs("ComputerName") & "</td></tr>"
Response.Write "<tr><td>Serial number</td><td>" & objRs("SerialNo") & "</td></tr>"
strColor = ""
If DateDiff("d", objRs("DTLastScan"), Now()) > 21 Then
	strColor = " style=""color:red"""
End If
Response.Write "<tr><td>ITAMWeb last connect date</td><td" & strColor & ">" & CISODateTime(objRs("DTLastScan")) & "</td></tr>"
Response.Write "<tr><td>Acquisition date</td><td>" & CISODate(objRs("DTAcquisition")) & "</td></tr>"
Response.Write "<tr><td>Asset tag</td><td>" & objRs("AssetTag") & "</td></tr>"
Response.Write "<tr><td>Internal tag</td><td>" & objRs("InternalTag") & "</td></tr>"
Response.Write "<tr><td>Category</td><td>" & objRs("CategoryName")& "</td></tr>"
Response.Write "<tr><td>Model</td><td colspan=""3"">" & objRs("ProductModel")& "</td></tr>"
strColor = ""
If InStr("In stock//In contract|In stock//In contract, for refresh|In use//In contract|In stock//Not in contract, to be disposed|Retired (or consumed)//Obsolete|In use//In contract, owned by SL|In use//Not in contract, used by HP", objRs("cf_HP_AssgnRead") & "//" & objRs("fv_SLDE_BillingStatus")) = 0 Then
	strColor = " style=""color:red"""
End If
Response.Write "<tr" & strColor & "><td>Billing status</td><td>" & objRs("fv_SLDE_BillingStatus")& "</td></tr>"
Response.Write "<tr" & strColor & "><td>Asset status</td><td>" & objRs("cf_HP_AssgnRead") & "</td></tr>"
If objRs("Status") <> "" Then
	Response.Write "<tr><td>Extra asset status</td><td>" & objRs("Status") & "</td></tr>"
End If
If objRs("ScannerDesc") <> "" Then
	Response.Write "<tr style=""color:red""><td>SW distr. status</td><td>" & objRs("ScannerDesc") & "</td></tr>"
End If

Response.Write "<tr><td colspan=""4""><hr/></td></tr>"
objRs.Close
'----------------------------------------------------------------------------
'Managed refresh information
'-------------------------------------------------------------CK---------------
If strAssetType = "Computer" And Session("UID") <> "KEES-QUICK" Then
	strSQL = "SELECT W.[ID], W.[Name], '' AS 'StatusCode', WI.[SLAgreement], WI.[SLComment], ISNULL(WI.[HPComment], '') AS 'HPComment' FROM RadiaRIMProd.dbo.mrWaveItem AS WI JOIN RadiaRIMProd.dbo.mrWave AS W "
	strSQL = strSQL & "ON WI.[WaveID] = W.[ID] WHERE WI.[InternalTag] = '" & strInternalTag & "' ORDER BY W.[ID] DESC"

	objRs.Open strSQL, objConn

	If Not objRs.EOF Then
		Response.Write "<tr><td colspan=4><b>Managed refresh information</td></tr>"

		Response.Write "<tr><td><i>Wave</i></td><td><i>Agreement</i></td><td colspan=""2""><i>Comment</i></td></tr>"
		While Not objRs.Eof
			strRowTitle = objRs("HPComment")
			If strRowTitle <> "" Then
				strRowTitle = "title="""& strRowTitle &""""
			End If
			Response.Write "<tr>"
			Response.Write "<td>" & objRs("Name") & "</td>"
			Response.Write "<td>" & objRs("SLAgreement") & "</td>"
			Response.Write "<td colspan=""2"" " & strRowTitle & ">" & objRs("SLComment") & "</td>"
			Response.Write "</tr>"

			objRs.MoveNext
		Wend
		Response.Write "<tr><td colspan=""4""><hr/></td></tr>"
	End If

	objRs.Close
End If

'----------------------------------------------------------------------------
' Multiple computer registrations on this user
'----------------------------------------------------------------------------
If strAssetType = "Computer" And strAccount <> "" And Session("UID") <> "KEES-QUICK" Then
	strSQL = "SELECT * FROM acAssetList AS AL "
	strSQL = strSQL & "WHERE AL.[InternalTag] <> '" & strInternalTag & "' "
	strSQL = strSQL & "AND AL.[SupervisorUserLogin] = '" & Replace(strAccount, "'", "''") & "' "
	'strSQL = strSQL & "AND AL.[fv_SLDE_BillingStatus]='In contract' "
	strSQL = strSQL & "AND AL.[CategoryName] IN ('Desktop computer', 'Laptop', 'Thin client', 'Netbook', 'Folio Laptop', 'Tablet') "
	strSQL = strSQL & "AND AL.[cf_HP_AssgnRead]='In use'"

	objRs.Open strSQL, objConn

	If Not objRs.EOF Then
		Response.Write "<tr><td colspan=4><b>Multiple computer registrations on this user</td></tr>"

		Response.Write "<tr><td><i>Computer name</i></td><td><i>Category</i></td><td colspan=""2""><i>Location</i></td></tr>"
		While Not objRs.EOF
			Response.Write "<tr>"
			Response.Write "<td>" & objRs("ComputerName") & "</td>"
			Response.Write "<td>" & objRs("CategoryName") & "</td>"
			strTmp = objRs("LocationName")
			If objRs("fv_SLDE_LocDetail") <> "" Then
				strTmp = strTmp & " - " & objRs("fv_SLDE_LocDetail")
			End If

			strInTag = ""
			If Len(strTmp) > 75 Then
				strTmp = Left(strTmp, 70) & "..."
				strInTag = "title=""" & objRs("fv_SLDE_LocDetail") & """"
			End If

			Response.Write "<td colspan=""2"" " & strInTag & ">" & strTmp & "</td>"
			Response.Write "</tr>"

			objRs.MoveNext
		Wend
		Response.Write "<tr><td colspan=""4""><hr/></td></tr>"
	End If

	objRs.Close
End If

'----------------------------------------------------------------------------
' Duplicate serial numbers and computer names
'----------------------------------------------------------------------------
If strAssetType = "Computer" And Session("UID") <> "KEES-QUICK" Then
	strSQL = "SELECT * FROM acAssetList AS AL "
	strSQL = strSQL & "WHERE AL.[InternalTag] <> '" & strInternalTag & "' "
	strSQL = strSQL & "AND AL.[cf_HP_AssgnRead] <> 'Retired (or consumed)' AND AL.[fv_SLDE_BillingStatus] <> 'Obsolete' "
	strSQL = strSQL & "AND AL.[CategoryName] IN ('Desktop computer', 'Laptop', 'Thin client', 'Folio Laptop', 'Tablet') "
	strSQL = strSQL & "AND (AL.[SerialNo] = '" & strSerialNo & "' "
	strSQL = strSQL & "OR REPLACE(REPLACE(AL.[ComputerName], '-RAD', ''), '-DUP', '') = '" & strComputerName & "') "

	objRs.Open strSQL, objConn

	If Not objRs.EOF Then
		Response.Write "<tr><td colspan=4><b>Duplicates on serial numbers and computer names</td></tr>"

		Response.Write "<tr>"
		Response.Write "<td><i>Computer name</i></td>"
		Response.Write "<td><i>Category</i></td>"
		Response.Write "<td><i>Duplicate on</i></td>"
		Response.Write "<td><i>Location</i></td>"
		Response.Write "</tr>"
		While Not objRs.EOF
			Response.Write "<tr>"
			Response.Write "<td><a href=""" & "pcdetail.asp?frmInternalTag=" & objRs("InternalTag") & """>" & objRs("ComputerName") & "//" & objRs("SerialNo") & "</a></td>"
			Response.Write "<td>" & objRs("CategoryName") & " (" & CISODateTime(objRs("DTLastScan")) & ")" & "</td>"
			
			If objRs("SerialNo") = strSerialNo Then
				Response.Write "<td>Serial no (" & objRs("fv_SLDE_BillingStatus") & ")</td>"
			Else
				Response.Write "<td>Computer name (" & objRs("fv_SLDE_BillingStatus") & ")</td>"
			End If
			
			strTmp = objRs("LocationName")
			If objRs("fv_SLDE_LocDetail") <> "" Then
				strTmp = strTmp & " - " & objRs("fv_SLDE_LocDetail")
			End If

			strInTag = ""
			If Len(strTmp) > 45 Then
				strTmp = Left(strTmp, 42) & "..."
				strInTag = "title=""" & objRs("fv_SLDE_LocDetail") & """"
			End If

			Response.Write "<td " & strInTag & ">" & strTmp & "</td>"
			Response.Write "</tr>"

			objRs.MoveNext
		Wend
		Response.Write "<tr><td colspan=""4""><hr/></td></tr>"
	End If

	objRs.Close
End If

If strAssetType = "Computer" And Session("UID") <> "KEES-QUICK" Then
  '----------------------------------------------------------------------------
  ' Display details from Active Directory
  '----------------------------------------------------------------------------
  strSQL = "SELECT * FROM ITAMData.dbo.adComputerComplete WHERE [ComputerName] = '" & strComputerName & "'"
  objRs.Open strSQL, objConn

  Response.Write "<tr><td colspan=4><b>Details from Active Directory</a></td></tr>"
  If Not objRs.EOF Then
		intCount = 1
		While Not objRs.EOF
			If intCount > 1 Then
				Response.Write "<tr><td colspan=""4""></td></tr>"
			End If
			strInTag = ""
			If Not IsNull(objRs("DTDeletionAD")) Then
				strInTag = "style=""color:red;"""
			End If

			If objRs("adDescription") <> "" Then
				Response.Write "<tr " & strInTag & "><td>Description</td><td colspan=3>" & objRs("adDescription") & "</td></tr>"
			End If
			Response.Write "<tr " & strInTag & "><td>Date creation</td><td colspan=3>" & CISODateTime(objRs("DTCreationAD")) & "</td></tr>"
			Response.Write "<tr " & strInTag & "><td>Operating system</td><td colspan=3>" & objRs("OSName") & "</td></tr>"
			Response.Write "<tr " & strInTag & "><td>Operating system service pack</td><td colspan=3>" & objRs("OSSP") & "</td></tr>"
			Response.Write "<tr " & strInTag & "><td>OU</td><td colspan=3>" & objRs("OUName") & "</td></tr>"
			strColor = ""
			If DateDiff("d", objRs("DTLastConnect"), Now()) >= 31 Then
				strColor = " style=""color:red"""
			End If
			Response.Write "<tr " & strInTag & "><td>Last registered activity</td><td colspan=""3""" & strColor & ">" & CISODateTime(objRs("DTLastConnect")) & "</td></tr>"
			Response.Write "<tr " & strInTag & "><td>DNS name</td><td colspan=3>" & objRs("adDNSName") & "</td></tr>"
			If Not IsNull(objRs("DTDeletionAD")) Then
				Response.Write "<tr " & strInTag & "><td>Date deletion</td><td colspan=3>" & CISODateTime(objRs("DTDeletionAD")) & "</td></tr>"
			End If
			intCount = intCount + 1
			objRs.MoveNext
		Wend
	Else
		Response.Write "<tr><td colspan=4 style=""color:red"">No data found in Active Directory table</td></tr>"
  End If
  Response.Write "<tr><td colspan=""4""><hr/></td></tr>"
  objRs.Close

If strAssetType = "Computer" And Session("UID") <> "KEES-QUICK" Then
  '----------------------------------------------------------------------------
  ' Display details from BitLocker
  '----------------------------------------------------------------------------
  strColor = ""
  strSQL = "SELECT * FROM ITAMProcess.BitLocker.acAssetListNotObsolete WHERE [AssetName] = '" & strComputerName & "'"
  objRs.Open strSQL, objConn

  Response.Write "<tr><td colspan=4><b>Details from BitLocker server</a></td></tr>"
  If Not objRs.EOF Then
    If objRs("BL_IsCompliant") <> "" Then
      If objRs("BL_DTReporting") < dtInstall Then
        strColor = " style=""color:red"" title=""Computer is later imaged then this reporting date"""
			  Response.Write "<tr " & strInTag & "><td>Compliant</td><td colspan=3 style=""color:red"">Computer is later (re-)imaged then the MBAM reporting date</td></tr>"
			  Response.Write "<tr " & strInTag & "><td>Date last connection</td><td colspan=3 " & strColor & ">" & CISODateTime(objRs("BL_DTReporting")) & "</td></tr>"
			ElseIf DateDiff("d", objRs("BL_DTReporting"), Now()) >= 60 Then
				strColor = " style=""color:red"" title=""Computer hasn't reported for 60 days or more"""
			  Response.Write "<tr " & strInTag & "><td>Compliant</td><td colspan=3 style=""color:red"">Computer hasn't reported for 60 days or more to the MBAM server</td></tr>"
			  Response.Write "<tr " & strInTag & "><td>Date last connection</td><td colspan=3 " & strColor & ">" & CISODateTime(objRs("BL_DTReporting")) & "</td></tr>"
			  Response.Write "<tr " & strInTag & "><td>Reported users</td><td colspan=3>" & objRs("BL_UserAccounts") & "</td></tr>"
      Else
			  Response.Write "<tr " & strInTag & "><td>Compliant</td><td colspan=3>" & objRs("BL_IsCompliant") & "</td></tr>"
			  Response.Write "<tr " & strInTag & "><td>Date last connection</td><td colspan=3 " & strColor & ">" & CISODateTime(objRs("BL_DTReporting")) & "</td></tr>"
			  Response.Write "<tr " & strInTag & "><td>Reported users</td><td colspan=3>" & objRs("BL_UserAccounts") & "</td></tr>"
			  Response.Write "<tr " & strInTag & "><td>Error message</td><td colspan=3>" & objRs("BL_ErrorInfoName") & "</td></tr>"
      End If
	  Else
		  Response.Write "<tr><td colspan=4 style=""color:red"">No data found in BitLocker table</td></tr>"
    End If
  End If
  Response.Write "<tr><td colspan=""4""><hr/></td></tr>"
  objRs.Close
End If

  '----------------------------------------------------------------------------
  ' Display details from SAV current situation
  '----------------------------------------------------------------------------
  strSQL = "SELECT * FROM ITAMNetwork.dbo.avComputer WHERE [ComputerName] = '" & strComputerName & "'"
    
	Response.Write "<tr><td colspan=4><b>Details from AntiVirus</td></tr>"
	Response.Write "<tr><td colspan=4><b>Actual status</td></tr>"
	On Error Resume Next
		objRs.Open strSQL, objConn
		Response.Write "<tr><td colspan=4><b></td></tr>"
		If Err.number = 0 Then
	    intID = objRs("ComputerID")
			Response.Write "<tr><td>Product</td><td>" & objRs("Product") & "</td></tr>"
			strColor = ""
			If Left(objRs("Version"), 3) <> "12." Then
				strColor = " style=""color:red"""
			End If
			Response.Write "<tr><td>Version</td><td" & strColor & ">" & objRs("Version") & "</td></tr>"
			Response.Write "<tr><td>Last connect</td><td>" & CISODateTime(objRs("DTLastConnect")) & "</td></tr>"
			strColor = ""
			If DateDiff("d", objRs("DTPattern"), objRs("DTLastConnect")) > 5 Then
				strColor = " style=""color:red"""
			End If
			Response.Write "<tr><td>Definition date</td><td" & strColor & ">" & CISODate(objRs("DTPattern")) & "</td></tr>"
			Response.Write "<tr><td>IP Address</td><td>" & objRs("IPAddress") & "</td></tr>"
			Response.Write "<tr><td>Server</td><td>" & objRs("Server") & "</td></tr>"
			Response.Write "<tr><td>Group</td><td>" & objRs("Group") & "</td></tr>"
			Response.Write "<tr><td>User</td><td>" & objRs("User") & "</td></tr>"
			objRs.MoveNext

		Else
			Response.Write "<tr><td colspan=4 style=""color:red"" title=""" + Err.Msg + """>Error collecting SAV data</td></tr>"
		End If
	  objRs.Close
	On Error Goto 0
	  
  '----------------------------------------------------------------------------
  ' Display details from SAV history
  '----------------------------------------------------------------------------
  strSQL = "SELECT TOP 10 * FROM ITAMNetwork.History.avComputer WHERE [ComputerID] = " & intID & " ORDER BY [DTLastConnect] DESC"

	On Error Resume Next
		objRs.Open strSQL, objConn
		If Err.number = 0 Then
			Response.Write "<tr><td colspan=4><b></td></tr>"
			Response.Write "<tr><td colspan=4><b>History</td></tr>"
			Response.Write "<tr><td><i>Date contact</i></td><td><i>Date virus definition</i></td><td><i>IP Address</i></td><td><i></i></td></tr>"
			While Not objRs.EOF
		    
				Response.Write "<tr>"
				Response.Write "<td>" & CISODateTime(objRs("DTLastConnect")) & "</td>"
				strColor = ""
				If DateDiff("d", objRs("DTPattern"), objRs("DTLastConnect")) > 5 Then
					strColor = " style=""color:red"""
				End If
				Response.Write "<td" & strColor & ">" & CISODate(objRs("DTPattern")) & "</td>"
				Response.Write "<td>" & objRs("IPAddress") & "</td>"
				Response.Write "</tr>"
				objRs.MoveNext
		    
			Wend
		Else
			Response.Write "<tr><td colspan=4 style=""color:red"">Error collecting SAV data</td></tr>"
		End If
	  objRs.Close
	On Error Goto 0
	  
  Response.Write "<tr><td colspan=""4""><hr/></td></tr>"

	strSQL = "SELECT [Device_ID] FROM ITAMSccm.ds.Device WHERE [IsActual] = 1 AND [IsActive] = 1 AND [ComputerName] = '" & strComputerName & "'"
	objRs.Open strSql, objConn

	If objRs.EOF Then
		strSource = "Radia"
	Else
		strSource = objRs("Device_ID")
	End If
	objRs.Close

	If strSource = "Radia" Then
		Response.Write "<tr><td colspan=4><b>Details from Radia Inventory Management</td></tr>"
		'----------------------------------------------------------------------------
		' Display basic computer details from deviceconfig table
		'----------------------------------------------------------------------------
		strSQL = "SELECT * FROM ITAMProcess.Radia.DeviceConfig WHERE [Device_ID] = '" & strComputerName & "'"
		objRs.Open strSQL, objConn

		If Not objRs.EOF Then
			Response.Write "<tr><td>IP address</td><td>" & objRs("ipaddr")& "</td></tr>"
			Response.Write "<tr><td>MAC address</td><td>" & objRs("macaddr")& "</td></tr>"
			Response.Write "<tr><td>Default language</td><td>" & objRs("language")& "</td></tr>"
			Response.Write "<tr><td>Operating system</td><td colspan=2>" & objRs("os")& "</td></tr>"
			Response.Write "<tr><td>OS Level</td><td>" & objRs("os_level")& "</td></tr>"
			Response.Write "<tr><td>Disk size (system)</td><td>" & objRs("sysdrv_total")& "</td></tr>"
			Response.Write "<tr><td>Memory</td><td>" & objRs("memory")& "</td></tr>"
			Response.Write "<tr><td>CPU Speed</td><td>" & objRs("cpu_speed")& "</td></tr>"
			Response.Write "<tr><td>BIOS</td><td>" & objRs("bios")& "</td></tr>"
			Response.Write "<tr><td>Video card</td><td>" & objRs("video")& "</td></tr>"
			Response.Write "<tr><td>Last modification date (in Radia)</td><td>" & CISODateTime(objRs("mtime")) & "</td></tr>"
		End if
		objRs.Close

		'----------------------------------------------------------------------------
		' display model info from Radia
		'----------------------------------------------------------------------------
		strSQL = "SELECT CASE WHEN SE.[wChassisTypes] IN (8, 9, 10, 12) THEN 'Laptop' " &_
			"WHEN SE.[wChassisTypes] IN (3, 4, 5, 6, 7, 15, 16) THEN 'Desktop' ELSE 'Unknown' END AS 'Chassis', CS.[wModel] AS 'Model', CS.[MTime] " &_
			"FROM ITAMProcess.Radia.rWin32_ComputerSystem AS CS LEFT OUTER JOIN ITAMProcess.Radia.rWin32_SystemEnclosure AS SE " &_
			"ON CS.[UserID] = SE.[UserID] WHERE CS.[UserID] = '" & strComputerName & "'"
		objRs.Open strSQL, objConn
  
		If Not objRs.EOF Then
			Response.Write "<tr><td>Scanned chassis</td><td>" & objRs("Chassis")& "</td></tr>"
			Response.Write "<tr><td>Scanned model</td><td>" & objRs("Model") & "</td></tr>"
			Response.Write "<tr><td>Scanned model</td><td>" & CISODateTime(objRs("MTime")) & "</td></tr>"
		End if
		objRs.Close

		'----------------------------------------------------------------------------
		' display BIOS info from Radia
		'----------------------------------------------------------------------------
		strSQL = "SELECT [wSerialNumber], [MTime] FROM ITAMProcess.Radia.rWin32_BIOS WHERE [UserID] = '" & strComputerName & "'"
		objRs.Open strSQL, objConn
  
		If Not objRs.EOF Then
			Response.Write "<tr><td>Serial number in BIOS</td><td>" & objRs("wSerialNumber") & "</td></tr>"
			Response.Write "<tr><td>Scan date/time</td><td>" & CISODateTime(objRs("MTime")) & "</td></tr>"
		End if
		objRs.Close

		'----------------------------------------------------------------------------
		' display OperatingSystem info from Radia
		'----------------------------------------------------------------------------
		strSQL = "SELECT [MTime], [wCurrentTimeZone], [wInstallDate], [wLastBootupTIme], [wVersion], [wWindowsDirectory] FROM ITAMProcess.Radia.rWin32_OperatingSystem WHERE [UserID] = '" & strComputerName & "'"
		objRs.Open strSQL, objConn
  
		If Not objRs.EOF Then
			Response.Write "<tr><td>Windows folder</td><td>" & objRs("wWindowsDirectory") & "</td></tr>"
			Response.Write "<tr><td>Windows version</td><td>" & objRs("wVersion") & "</td></tr>"
			Response.Write "<tr><td>Installation date</td><td>" & CISODateTime(objRs("wInstallDate")) & "</td></tr>"
			Response.Write "<tr><td>Last boot time</td><td>" & CISODateTime(objRs("wLastBootupTime")) & "</td></tr>"
			Response.Write "<tr><td>Last scan</td><td>" & CISODateTime(objRs("MTime")) & "</td></tr>"
			Response.Write "<tr><td>Time zone</td><td>" & objRs("wCurrentTimeZone") & "</td></tr>"
		End if
		objRs.Close
		Response.Write "<tr><td colspan=""4""><hr/></td></tr>"
		'----------------------------------------------------------------------------
		' display user right exceptions
		'----------------------------------------------------------------------------
		strSQL = "SELECT [DTLastScan], [LocalGroup], [GroupMember] FROM ITAMProcess.Process.UserRightExceptions WHERE [ComputerName] = '" & strComputerName & "' " &_
			"ORDER BY [DTLastScan], [LocalGroup], [GroupMember]"
		objRs.Open strSQL, objConn

		Response.Write "<tr><td colspan=""4""><b>User right exceptions</b></td><td/></tr>"
		Response.Write "<tr><td><i>Date monitoring</i></td><td><i>Local group</i></td><td><i>Group member</i></td><td><i></i></td></tr>"
		While Not objRs.EOF
			Response.Write "<tr>"
			Response.Write "<td>" & CISODateTime(objRs("DTLastScan")) & "</td>"
			Response.Write "<td>" & objRs("LocalGroup") & "</td>"
			Response.Write "<td>" & objRs("GroupMember") & "</td>"
			Response.Write "</tr>"
			objRs.MoveNext
		Wend
		objRs.Close
		Response.Write "<tr><td colspan=""4""><hr/></td></tr>"

		'----------------------------------------------------------------------------
		' display Logon History information	
		'----------------------------------------------------------------------------
		strSQL = "SELECT [Account], COUNT(*) AS 'NumberOfLoggedOn', MAX([DTLogon]) AS 'LastLoggedOn', CASE WHEN MAX([ADAccount]) IS NULL THEN 'No information' " &_
			"WHEN MIN([adStatus]) = 1 THEN '' WHEN MIN([adStatus]) & 128 = 128 THEN 'Account deleted' WHEN MIN([adStatus]) & 2 = 2 THEN 'Account disabled' " &_
			"ELSE 'Status ' + CAST(MAX([adStatus]) AS varchar) END AS 'Comment' " &_
			"FROM RadiaRIMProd.dbo.aisImportLogHist LEFT OUTER JOIN ITAMNetwork.dbo.adUser AS U ON [Account] = U.[adAccount] WHERE [Account] NOT LIKE '%\SYSTEM' AND [Device_ID] = '"
		strSQL = strSQL & strComputerName & "' GROUP BY [Account] ORDER BY MAX([DTLogon]) DESC, COUNT(*) DESC"
		objRs.Open strSQL, objConn

		Response.Write "<a name=""LogHist"">"
		Response.Write "<tr><td colspan=""4""><b>Logon history</b></td><td/></tr>"

		If Not objRs.EOF Then
			Response.Write "<tr><td><i>Scanned account</i></td><td><i>Logon count</i></td><td><i>Last logon date</i></td><td><i>Comment</i></tr>"

			While Not objRs.EOF
				strColor = ""
				If objRs("Comment") <> "" Then strColor = "style=""color:red;"""
				Response.Write "<tr " & strColor & "><td>" & objRs("Account") & "</td>"
				Response.Write "<td>" & objRs("NumberOfLoggedOn") & "</td>"
				Response.Write "<td>" & CISODateTime(objRs("LastLoggedOn")) & "</td>"
				Response.Write "<td>" & objRs("Comment")& "</td></tr>"
				objRs.MoveNext
			Wend
		Else
			Response.Write "<tr><td colspan=4 style=""color:red"">Logon History audit has not (yet) been taken place</td></tr>"
		End If

		objRs.Close
		Response.Write "<tr><td colspan=""4""><hr/></td></tr>"
		'----------------------------------------------------------------------------
		' display basic PC details from printer table
		'----------------------------------------------------------------------------
		strSQL="SELECT wDeviceId, wPortName FROM ITAMProcess.Radia.rWin32_Printer"
		strSQL=strSQL &   " WHERE userid="& Chr(39)& strComputerName & chr(39)
		Response.Write "<tr><td colspan=""4""><b>Printers</b></td><td/></tr>"
		objRs.Open strSQL, objConn
		While Not objRs.Eof
			Response.Write "<tr><td>" & objRs("wDeviceID")& "</td><td>" & objRs("wPortName")&"</td><td/></tr>"
			objRs.MoveNExt
		Wend
		objRs.Close
	Else
		'============================================================================
		' SCCM
		'============================================================================
		Response.Write "<tr><td colspan=4><b>Inventory details from SCCM (" & strSource & ")</td></tr>"

		strSQL = "SELECT [SerialNumber00], CAST([TimeKey] AS smalldatetime) AS 'TimeKey' FROM ITAMSccm.sccm.PC_BIOS_DATA WHERE [MachineID] = " & strSource
		objRs.Open strSQL, objConn

		If Not objRs.EOF Then
			Response.Write "<tr><td>Serial number in BIOS</td><td>" & objRs("SerialNumber00") & "</td></tr>"
		End If
		objRs.Close

		strSQL = "SELECT [Manufacturer00], [Model00], CAST([TimeKey] AS smalldatetime) AS 'TimeKey' FROM ITAMSccm.clr.Computer_System_DATA WHERE [MachineID] = " & strSource
		objRs.Open strSQL, objConn

		If Not objRs.EOF Then
			Response.Write "<tr><td>Brand</td><td>" & objRs("Manufacturer00") & "</td></tr>"
			Response.Write "<tr><td>Model</td><td>" & objRs("Model00") & "</td></tr>"
			Response.Write "<tr><td>Scan date/time</td><td>" & CISODateTime(objRs("TimeKey")) & "</td></tr>"
		End If
		objRs.Close

		strSQL = "SELECT CAST([TotalVisibleMemorySize00] AS int) AS 'TotalVisibleMemorySize00', CAST([InstallDate00] AS smalldatetime) AS 'InstallDate00', CAST([LastBootUpTime00] AS smalldatetime) AS 'LastBootUpTime00', CAST([TimeKey] AS smalldatetime) AS 'TimeKey' FROM ITAMSccm.sccm.Operating_System_DATA WHERE [MachineID] = " & strSource
		objRs.Open strSQL, objConn

		If Not objRs.EOF Then
			Response.Write "<tr><td>Memory size</td><td>" & objRs("TotalVisibleMemorySize00") & "</td></tr>"
			Response.Write "<tr><td>Install date</td><td>" & CISODateTime(objRs("InstallDate00")) & "</td></tr>"
			Response.Write "<tr><td>Last recorded boot</td><td>" & CISODateTime(objRs("LastBootUpTime00")) & "</td></tr>"
			Response.Write "<tr><td>Scan date/time</td><td>" & CISODateTime(objRs("TimeKey")) & "</td></tr>"
		End If
		objRs.Close

		strSQL = "SELECT ISNULL([SiteName], 'Unknown') AS 'SiteName', CASE [HasAgent] WHEN 1 THEN 'Has agent installed' WHEN 0 THEN 'Agent has been removed' ELSE 'Agent has not been installed' END AS 'AgentStatus', CASE [IsActual] WHEN 1 THEN 'Actual instance' WHEN 0 THEN 'Non actual instance' ELSE 'Unknown status' END AS 'RecordStatus', ISNULL([DTMutation], [DTCreation]) AS 'DTLastScan' FROM ITAMSccm.ds.Device WHERE [Device_ID] = " & strSource
		objRs.Open strSQL, objConn

		If Not objRs.EOF Then
			Response.Write "<tr><td>Scanned AD Site name</td><td>" & objRs("SiteName") & "</td></tr>"
			Response.Write "<tr><td>Agent status</td><td>" & objRs("AgentStatus") & "</td></tr>"
			Response.Write "<tr><td>Record status</td><td>" & objRs("RecordStatus") & "</td></tr>"
			Response.Write "<tr><td>Scan date/time</td><td>" & CISODateTime(objRs("DTLastScan")) & "</td></tr>"
		End If
		objRs.Close

		Response.Write "<tr><td colspan=""4""><hr/></td></tr>"

		'----------------------------------------------------------------------------
		' display Logon History information	
		'----------------------------------------------------------------------------
		strSQL = "SELECT U.[UserAccount], CASE WHEN AD.[adAccount] IS NULL THEN 'No information' "
		strSQL = strSQL & "WHEN AD.[adStatus] = 1 THEN '' WHEN AD.[adStatus] & 128 = 128 THEN 'Account deleted' "
		strSQL = strSQL & "WHEN AD.[adStatus] & 2 = 2 THEN 'Account disabled' WHEN AD.[adStatus] & 4 = 4 THEN 'Account locked' "
		strSQL = strSQL & "ELSE 'Status ' + CAST(AD.[adStatus] AS varchar) "
		strSQL = strSQL & "END AS 'Comment', ISNULL(S.[Description], '') AS 'Hint', ISNULL(S.[ComputerCount], 0) AS 'ComputerCount', "
		strSQL = strSQL & "U.[TotalTime], U.[LogonCount], U.[DTLastLogon], U.[Device_ID] "
		strSQL = strSQL & "FROM ITAMSccm.ds.UserLogon AS U "
		strSQL = strSQL & "JOIN ITAMSccm.ds.Device AS D ON U.[Device_ID] = D.[Device_ID] "
		strSQL = strSQL & "LEFT OUTER JOIN ITAMData.dbo.adUserCompleter AS AD ON U.[UserAccount] = AD.[adAccount] "
		strSQL = strSQL & "LEFT OUTER JOIN ITAMData.dbo.UserAccountSpecial AS S ON U.[UserAccount] = S.[UserAccount] "
		strSQL = strSQL & "WHERE U.[UserAccount] NOT LIKE '%\Local_Users' "
		strSQL = strSQL & "AND D.[ComputerName] = '" & strComputerName & "' ORDER BY U.[Device_ID] DESC, S.[Description], U.[TotalTime] * CASE WHEN AD.[adStatus] & 2 = 2 OR AD.[adStatus] & 4 = 4  OR AD.[adStatus] & 128 = 128 OR AD.[adStatus] IS NULL THEN 0 ELSE 1 END DESC, U.[LogonCount]"
		objRs.Open strSQL, objConn

		Response.Write "<a name=""LogHist"">"
		Response.Write "<tr><td colspan=""4""><b>Logon history through SCCM</b></td><td/></tr>"

		If Not objRs.EOF Then
			Response.Write "<tr><td><i>Scanned account</i></td><td><i>Logon time / count</i></td><td><i>Last logon date</i></td><td><i>Comment</i></tr>"

			While Not objRs.EOF
				strTmp = ""
				strColor = ""
				If objRs("Comment") <> "" Then strColor = "style=""color:red;"
				If objRs("Device_ID") <> strSource Then
					strTmp = "Before last re-image"
					If strColor = "" Then strColor = "style="""
					strColor = strColor & "font-style: italic;"
				End If
				If strColor = "" And (objRs("Hint") <> "" Or objRs("ComputerCount") >= 25) Then
					strColor = "style=""color:blue;"
				End If
				If strColor <> "" Then strColor = strColor & """"

				Response.Write "<tr " & strColor & " title=""" & strTmp & """><td title=""" & objRs("Hint") & " (" & objRs("ComputerCount") & " computers)" & """>" & objRs("UserAccount") & "</td>"
				Response.Write "<td>" & objRs("TotalTime") & "/" & objRs("LogonCount") & "</td>"
				Response.Write "<td>" & CISODateTime(objRs("DTLastLogon")) & "</td>"
				Response.Write "<td>" & objRs("Comment")& "</td></tr>"
				objRs.MoveNext
			Wend
		Else
			Response.Write "<tr><td colspan=4 style=""color:red"">Logon History audit has not (yet) been taken place</td></tr>"
		End If

		objRs.Close
		Response.Write "<tr><td colspan=""4""><hr/></td></tr>"
	
	End If

  Response.Write "<tr><td colspan=""4""><hr/></td></tr>"
  '----------------------------------------------------------------------------
  ' display basic PC details from software table
  '----------------------------------------------------------------------------
  Response.Write "<tr><td colspan=""2""><b>Software list</b></td><td/></tr>"
  strSQL="SELECT MAX([DTLastScan]) AS 'DTLastScan' "
  strSQL=strSQL &  "FROM ITAMReport.dbo.swrReport "
  strSQL=strSQL &  "WHERE [ComputerName] = " & Chr(39)& strComputerName & chr(39)' & " AND [Certification] <> 'OS' "
  strSQL=strSQL &  "GROUP BY [ComputerName]"
  objRs.Open strSQL, objConn

	If Not objRs.EOF Then
		intCount = DateDiff("w", objRs("DTLastScan"), Now())
		strTmp = objRs("DTLastScan")
		If intCount <= 0 Then
			Response.Write "<tr><td colspan=""4"">The scan has been performed in the last 7 days</td></tr>"
		ElseIf intCount <= 4 Then
			Response.Write "<tr><td colspan=""4"">The scan has been performed " & intCount & " week ago</td></tr>"
		Else
			Response.Write "<tr><td colspan=""4"" style=""color:red"">The scan has been performed " & intCount & " week ago</td></tr>"
		End If
	
		objRs.Close
	
	  Response.Write "<tr><td><i>Software</i></td><td><i>Vendor</i></td><td><i>Version</i></td><td><i>Language</i></td></tr>"
		strSQL="SELECT ISNULL([Application] ,'') AS 'Application', "
		strSQL=strSQL &  "ISNULL([Vendor], '') AS 'Vendor', "
		strSQL=strSQL &  "ISNULL([Version], '') AS 'Version', "
		strSQL=strSQL &  "ISNULL([Language], '') AS 'Language', "
		strSQL = strSQL & "[DTLastScan] "
		strSQL=strSQL &  "FROM ITAMReport.dbo.swrReport "
		strSQL=strSQL &  "WHERE [ComputerName] = " & Chr(39)& strComputerName & chr(39) & " "
		strSQL=strSQL &  "ORDER BY [Application]"
		objRs.Open strSQL, objConn
		While Not objRs.EOF
			strColor = ""
			If DateDiff("d", objRs("DTLastScan"), strTmp) > 1 Then strColor = "" '" style=""color=red"""
			Response.Write "<tr" & strColor & "><td title=""Last scanned on: " & CISODateTime(objRs("DTLastScan")) & """>" & objRs("Application")& "</td><td>" & objRs("Vendor") & "</td><td>" & objRs("Version") & "</td><td>" & objRs("Language") & "</td></tr>"
			objRs.MoveNext
		Wend
	Else
		Response.Write "<tr><td colspan=4 style=""color:red"">Software audit has not (yet) been taken place</td></tr>"
	End If
	objRs.Close
  Response.Write "<tr><td colspan=""4""><hr/></td></tr>"
End If 
%>
</table>

<table align="center" cols="9" width="970">
<%
'----------------------------------------------------------------------------
' Display IMACD details from asset table
'----------------------------------------------------------------------------

strSQL = "SELECT CAST([ActionDate] AS smalldatetime) AS 'Date', [TypeRequest] AS 'Type', " &_
  "ISNULL(NULLIF([NewSerialNr], ''), [OldSerialNr]) AS 'SerialNumber', " &_
  "ISNULL(NULLIF([NewAssetName], ''), [OldAssetName]) AS 'AssetName', [Category] AS 'Category', " &_
  "[Brand] AS 'Brand', [Model] AS 'Model', [CountryOfLocation] AS 'Country', [LocationOfAsset] AS 'Location', " &_
  "[DetailLocation] AS 'LocationDetail', [EWM] AS 'ChangeRef', ISNULL(NULLIF([InvoiceType], ''), 'no status') AS 'BillingStatus', [Status] AS 'AssetStatus', " &_
  "[OpCo] AS 'OpCoName', [OnSiteEng] AS 'Engineer', [UserLName] AS 'LastName', [UserFName] AS 'FirstName', [UserEmail] AS 'EMailAddress', " &_
  "CAST([InstallDate] AS smalldatetime) AS 'InstallationDate', CAST([PurchaseDate] AS smalldatetime) AS 'AcquisitionDate', [ChangeID], " &_
  "ISNULL([Transflag], 'N/A') AS 'Transflag' " &_
  "FROM RadiaRIMProd.dbo.acMACD WHERE [TypeRequest] <> 'Ignore' AND ([OldAssetName] = '" & strComputerName & "' OR [NewAssetName] = '" & strComputerName & "' " &_
  "OR [OldSerialNr] = '" & strSerialNo & "' OR [NewSerialNr] = '" & strSerialNo & "') OR [InternalTag] = '" & strInternalTag & "' " &_
  "ORDER BY [ActionDate], [TypeRequest], ISNULL(NULLIF([NewAssetName], ''), [OldAssetName])"

objRs.Open strSQL, objConn

'Table header
Response.Write "<tr><th colspan=""9"" align=""left"">IMACD web form changes</th></tr>"
Response.Write "<tr><th align=""Left"">#</th><th align=""Left"">Date</th><th align=""Left"">Type</th><th align=""Left"">Serial no</th><th align=""Left"">Asset name</th>"
Response.Write "<th align=""Left"">Model</th><th align=""Left"">Billing status</th><th align=""Left"">OpCo</th><th align=""Left"">Engineer</th>"
Response.Write "<th align=""Left"">Acq. date</th></tr>"

If Not objRs.EOF Then
	intCount = 0
  While Not objRs.EOF
		intCount = intCount + 1
		If objRs("Transflag") <> "Y" Then
			strStyle = "font-style:italic"
			If InStr("ADD|UPDATE|REMOVE", UCase(objRs("Type"))) > 0 Then
				strStyle = strStyle & ";color:red"
			End If
			strStyle = "style=""" & strStyle & """ "
		Else
			If InStr("ADD|UPDATE|REMOVE", UCase(objRs("Type"))) > 0 Then
				strStyle = "style=""" & ";color:blue" & """ "
			Else
				strStyle = ""
			End If
		End If
    strDate = Left(objRs("Date"), InStr(objRs("Date"), " "))
    If strDate = "" Then strDate = objRs("Date")
    
    strSN = objRs("SerialNumber")
    If strSN = "" Then
			strSN = "<i>no serial no</i>"
    ElseIf Len(strSN) > 10 Then 
			strSN = "<i>" & Left(strSN, 10) & "...</i>"
		End If
    
    strRowTitle = "Change date:" & CISODateTime(objRs("Date")) & vbCrLf
    strRowTitle = strRowTitle & "Change type: " & objRs("Type") & vbCrLf
    strRowTitle = strRowTitle & "Serial number: " & objRs("SerialNumber") & vbCrLf
    strRowTitle = strRowTitle & "Asset name: " & objRs("AssetName") & vbCrLf
    strRowTitle = strRowTitle & "Category: " & objRs("Category") & vbCrLf
    strRowTitle = strRowTitle & "Brand: " & objRs("Brand") & vbCrLf
    strRowTitle = strRowTitle & "Model: " & objRs("Model") & vbCrLf
    strRowTitle = strRowTitle & "Country: " & objRs("Country") & vbCrLf
    strRowTitle = strRowTitle & "Location: " & objRs("Location") & vbCrLf
    strRowTitle = strRowTitle & "Location detail: " & objRs("LocationDetail") & vbCrLf
    strRowTitle = strRowTitle & "Change reference: " & objRs("ChangeRef") & vbCrLf
    strRowTitle = strRowTitle & "Billing status: " & objRs("BillingStatus") & vbCrLf
    strRowTitle = strRowTitle & "Asset status:" & objRs("AssetStatus") & vbCrLf
    strRowTitle = strRowTitle & "OpCo name: " & objRs("OpCoName") & vbCrLf
    strRowTitle = strRowTitle & "Engineer: " & objRs("Engineer") & vbCrLf
    strRowTitle = strRowTitle & "Last name: " & objRs("LastName") & vbCrLf
    strRowTitle = strRowTitle & "Frist name: " & objRs("FirstName") & vbCrLf
    strRowTitle = strRowTitle & "E-mail address: " & objRs("EMailAddress") & vbCrLf
    strRowTitle = strRowTitle & "Installation date: " & CISODate(objRs("InstallationDate")) & vbCrLf
    strRowTitle = strRowTitle & "Acquisition date: " & CISODate(objRs("AcquisitionDate"))
    
    Response.Write "<tr " & strStyle & "title=""" & strRowTitle & """>"
    Response.Write "<td><a href=""javascript:showimacd('ChangeID=" & objRs("ChangeID") & "')"">" & intCount & "</a></td>"
    Response.Write "<td>" & Left(CISODateTime(objRs("Date")), 16) & "</td>"
    Response.Write "<td>" & objRs("Type") & "</td>"
    Response.Write "<td>" & strSN & "</td>"
    Response.Write "<td>" & objRs("AssetName") & "</td>"
    Response.Write "<td>" & objRs("Model") & "</td>"
    Response.Write "<td>" & objRs("BillingStatus") & " / " & objRs("AssetStatus") & "</td>"
    Response.Write "<td>" & Left(objRs("OpCoName"), 4) & "</td>"
    Response.Write "<td>" & objRs("Engineer") & "</td>"
    Response.Write "<td>" & CISODate(objRs("AcquisitionDate")) & "</td>"
    Response.Write "</tr>"
    objRs.MoveNext
  Wend
End If
objRs.Close

Call acCloseDB()
Response.Write "<tr><td height=""15"" class=""menutitle"" colspan=10>&nbsp;</td></tr>"
 %>
</table>

<p><%=strDebug %></p>

<!--#include file="Include\pageFooter.asp"-->
