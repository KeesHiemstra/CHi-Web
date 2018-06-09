<%@ LANGUAGE="VBSCRIPT"%>
<% Option Explicit%>
<%Response.buffer = True%>
<%Response.Expires = 0%>

   <!--#include file="Include/globalVars.asp"-->
   <!--#include file="Include/dbFunction.asp"-->
   <!--#include file="Include/strFunction.asp"-->

<%
' *********************
'   Declare variables
' *********************
Dim PageAction
Dim NextAction
Dim strSearch
Dim ArrDupli
Dim strItem
Dim intColumn
Dim colorid
Dim CoreRecord

Const intMaxRows = 7

Dim RecordID(7)
Dim DtLastScan(7)
Dim NewAssetName(7)
Dim NewSerialNr(7)
Dim AssetTag(7)
Dim CountryOfLocation(7)
Dim LocationOfAsset(7)
Dim LocDetail(7)
Dim InstallDate(7)
Dim Opco(7)
Dim CostLoc(7)
Dim PurchaseDate(7)
Dim InvoiceType(7)
Dim Status(7)
Dim Comment(7)
Dim Category(7)
Dim Brand(7)
Dim Model(7)
Dim ProductId(7)
Dim NTLogon(7)
Dim UserLName(7)
Dim UserFName(7)
Dim UserPhone(7)
Dim UserEMail(7)
Dim InternalTag(7)

Dim intRow
Dim strMessage
Dim CurrentDate
Dim strFormStrings
Dim arrFormStrings
Dim strTableStrings
Dim arrTableStrings
Dim strSummaryStrings
Dim arrSummaryStrings
Dim strSQL
Dim intCount
Dim strTmp
Dim strTmp2
Dim iChangeID
Dim strList
Dim strNextPage

' **************************
'   Define variable values
' **************************
PageTitle = "Search duplicate assets"
PageAction = UCase(Request("PageAction"))
strSearch = Request("search")
strItem = Request("Item")
intColumn = Request("intColumn")
CoreRecord = Request("CoreRecord")
If PageAction = "SELECT" Then MenuSize = "Large"

'***********************************************************************************
' START Define functions and subs
'***********************************************************************************

' Open database with extra objects
Public Function acOpenDBExtra()
  If acOpenDB() Then
    Set objRsUser = Server.CreateObject("ADODB.Recordset")
  End If
End Function

' Close database with extra objects
Public Function acCloseDBExtra()
  Set objRsUser = Nothing
  Call acCloseDB()
End Function

' =======================
' CstrSQL
' double single quoutes
' =======================
Function CstrSQL(strLine)
  On Error Resume Next
  If IsNull(strLine) Then
    CstrSQL = ""
  Else
    CstrSQL = Replace(strLine, "'", "''")
  End If
  On Error Goto 0
End Function

' ******************************************************************
'  Check if the user is logged otherwise redirect to the login page
' ******************************************************************
If Session("UID") = "" Or Session("UID") = Null Then
  Session("AfterLoginGoto") = Request.ServerVariables("SCRIPT_NAME")
  Response.Redirect "Logon.asp"
  Response.End
End if

' **********************************
'  Define Subroutines and Functions
' **********************************
Public Sub CreateArray(StrValue, strQuery, StrArray)
  ' Remember that when using this function the first item in the array (0) is empty!
  ' This is because the for-next loop adds a heading | in the variable
  ' So every next for-next loop should be formatted as: for x = 1 to ubound(ArrName)
  Dim objRsTmp
  Dim strTmpFields
  Set objRsTmp = Server.CreateObject("ADODB.Recordset")

  strTmpFields = ""
  objRsTmp.Open strQuery, objConn, 3', 3
  For intCount = 0 To objRsTmp.RecordCount -1
    strTmpFields = strTmpFields & "|" & objRsTmp(strValue)
    If strList <> "" Then strList = strList & "|"
    strList = strList & objRsTmp(strValue)
   objRsTmp.MoveNext
  Next
  objRsTmp.Close
  strArray = Split(strTmpFields, "|")
  strTmpFields = ""
  Set objRsTmp = Nothing
End Sub

Public Sub CreateRow(FriendlyName, arrName, strColor)
  response.write "              <!-- CreateRow -->"  & vbCrLf
  response.write "              <tr class=""" & strColor & """>"  & vbCrLf
  response.write "                <th align=""right"">" & FriendlyName & "</th>" & vbCrLf
  For intCount = 0 To intColumn
    If intCount = CoreRecord Then 
      ColorID="bgcolor1"
    Else 
      ColorID="bgcolor2"
    End If
    response.write "                <td class=""" & ColorID & """>" & vbCrLf
    response.write "                  <input type=""radio"" name=""" & FriendlyName & """ value="""  & Trim(arrName(intCount)) & """ class=""radioinput"""
    
    If intCount = CoreRecord Then
      response.write " checked"
    End If
    
    response.write " />" & vbCrLf
    response.write "                  <input type=""hidden"" name=""" & FriendlyName & intCount & """ value=""" & Trim(arrName(intCount)) & """ />" & vbCrLf
    response.write "                </td>" & vbCrLf
    response.write "                <td align=""left"" class=""" & ColorID & """>"& vbCrLf
    response.write "                  " & ClearText(arrName(intCount)) & vbCrLf
    response.write "                </td>" & vbCrLf
  Next
  response.write "              </tr>" & vbCrLf
End sub

Public Sub CreateDisabledRow(FriendlyName, arrName, strColor)
  response.write "              <!-- CreateDisabledRow -->"  & vbCrLf
  response.write "              <tr class=""" & strColor & """>"  & vbCrLf
  response.write "                <th align=""right"">" & FriendlyName & "</th>" & vbCrLf
  For intCount = 0 to intColumn
    If intCount = CoreRecord Then 
      ColorID="bgcolor1" 
    Else 
      ColorID="bgcolor2"
    End If
    response.write "                <td class=""" & ColorID & """>" & vbCrLf
    response.write "                  <input type=""radio"" name=""" & FriendlyName & """ value="""  & Trim(arrName(intCount)) & """ class=""radioinput"" disabled"

    If intCount = CoreRecord Then 
      response.write " checked"
    End If

    response.write " />" & vbCrLf
    If intCount = CoreRecord Then 
      response.write "                  <input type=""hidden"" name=""" & FriendlyName & """ value=""" & Trim(arrName(intCount)) & """ />" & vbCrLf
    End If
    response.write "                  <input type=""hidden"" name=""" & FriendlyName & intCount & """ value=""" & Trim(arrName(intCount)) & """ />" & vbCrLf
    response.write "                </td>" & vbCrLf
    response.write "                <td align=""left"" class=""" & colorid & """>" & vbCrLf
    response.write "                  " & ClearText(arrName(intCount)) & vbCrLf
    response.write "                </td>" & vbCrLf
  Next
  response.write "              </tr>" & vbCrLf
End sub

Public Sub CreateDisabledRowPlus(FriendlyName, arrName, strColor)
  response.write "              <!-- CreateDisabledRow -->"  & vbCrLf
  response.write "              <tr class=""" & strColor & """>"  & vbCrLf
  response.write "                <th align=""right"">" & FriendlyName & "</th>" & vbCrLf
  For intCount = 0 to intColumn
    If intCount = CoreRecord Then 
      ColorID="bgcolor1" 
    Else 
      ColorID="bgcolor2"
    End If
    response.write "                <td class=""" & ColorID & """>" & vbCrLf
    response.write "                  <input type=""radio"" name=""" & FriendlyName & """ value="""  & Trim(arrName(intCount)) & """ class=""radioinput"" disabled"

    If intCount = CoreRecord Then 
      response.write " checked"
    End If

    response.write " />" & vbCrLf
    response.write "                  <input type=""hidden"" name=""" & FriendlyName & intCount & """ value=""" & Trim(arrName(intCount)) & """ />" & vbCrLf
    response.write "                </td>" & vbCrLf
    response.write "                <td align=""left"" class=""" & colorid & """>" & vbCrLf
    response.write "                  <a target=""_blank"" href=""" & "pcdetail.asp?frmInternalTag=" & ClearText(arrName(intCount)) & """>" & ClearText(arrName(intCount)) & "</a>" & vbCrLf
    response.write "                </td>" & vbCrLf
  Next
  response.write "              </tr>" & vbCrLf
End sub

Public Sub CreateOnClickRow(FriendlyName, arrName)
  response.write "              <!-- CreateOnClickRow -->"  & vbCrLf
  response.write "              <tr>" & vbCrLf
  response.write "                <th align=""right"">" & FriendlyName & "</th>" & vbCrLf
  for intCount = 0 to intColumn
    If intCount = CoreRecord Then
      ColorID="bgcolor1"
    Else 
      ColorID="bgcolor2"
    End If
    response.write "                <td class=""" & ColorID & """>" & vbCrLf
    response.write "                  <input type=""radio"" name=""" & FriendlyName & """ value="""  & Trim(arrName(intCount)) & """ class=""radioinput"""

    If intCount = CoreRecord Then 
      response.write " checked"
    End If

    response.write " onclick=""selectuser('" & intCount & "');"" />" & vbCrLf
    response.write "                  <input type=""hidden"" name=""" & FriendlyName & intCount & """ value=""" & Trim(arrName(intCount)) & """ />" & vbCrLf
    response.write "                </td>" & vbCrLf
    response.write "                <td align=""left"" class=""" & ColorID & """>" &vbCrLf
    response.write "                  " & ClearText(arrName(intCount)) & vbCrLf
    response.write "                </td>" & vbCrLf
  Next
  response.write "              </tr>" & vbCrLf
End Sub

Public Sub CreateSummaryRow(ItemName)
  response.write "<tr>"
  response.write "<td width=""300"" class=""bgcolor1"">" & ItemName & "</td>"
  response.write "<td width=""300"" class=""bgcolor2"">" & request(ItemName) & "</td>"
  response.write "<td><br /><br /></td></tr>"
End sub

' Start opening the database
Call acOpenDB()

Public Function NextItem()
	Dim intPos, strNextItem

	intPos = InStr(UCase(strList), UCase(strItem))
	If intPos + Len(strItem) + 1 < Len(strList) Then
		strNextItem = Mid(strList, intPos + Len(strItem) + 1)
		intPos = InStr(strNextItem, "|")
		If intPos > 0 Then
			strNextItem = Left(strNextItem, intPos - 1)
		End If
		NextItem = strNextItem
	Else
		NextItem = ""
	End If
End Function

' *********************************************
'             PAGEACTION DATA FLOW
'  Controlled by hidden form field in the page
' *********************************************
Select Case UCase(PageAction)
' Depending on the action of the page the following actions needs to be executed.
	Case ""
	' PageAction is empty (first time on the page, not yet posted any)
	' There is nothing to be prepared, show only the box to select Duplicate Assetnames or Duplicate serial numbers.

		NextAction = "Search"
' ******************** SEARCH
Case "SEARCH"
  'User selected one of the search options

	NextAction = "Select"
  If UCase(strSearch) = "SERIALNO" Then
    strSQL = "SELECT AL.[SerialNo] FROM RadiaRIMProd.dbo.acAssetList AS AL "
    strSQL = strSQL + "LEFT OUTER JOIN ITAMProcess.Process.acIMACD AS I "
		strSQL = strSQL + "ON AL.[SerialNo] = I.[SerialNo] OR AL.[SerialNo] = I.[OldSerialNo] "
		strSQL = strSQL + "WHERE AL.[SerialNo] <> '' AND I.[ID] IS NULL "
		strSQL = strSQL + "GROUP BY AL.[SerialNo] HAVING (COUNT(*) > 1 AND COUNT(*) <= " & intMaxRows & ") ORDER BY AL.[SerialNo] "
    '2006-05-03, Kees Hiemstra, Added Active = 1 to prevent duplicates that aren't duplicates at all.
'    If UCase(Session("CountryAccess")) <> "ALL" Then
'      strSQL = strSQL + "(CHARINDEX(udHRDataCountry.[CountryCode], '" & session("CountryAccess") & "') <> 0 OR [LocationCountry] = '' OR udHRDataCountry.[CountryCode] IS NULL) AND "
'    End if
    '2006-05-06, Kees Hiemstra. Prevent changed items to be selected again before the MACD export.
    
'    strSQL = strSQL + "[SerialNo] NOT IN (SELECT ISNULL([NewSerialNr], '') AS 'Name' FROM RadiaRIMProd.dbo.acMACDActual "
'    strSQL = strSQL + "UNION SELECT ISNULL([OldSerialNr], '') AS 'Name' FROM acMACDActual) "

		'(2009-08-10, Kees Hiemstra: Added to exclude duplicate serial number exceptions)
    'strSQL = strSQL + "AND [SerialNo] NOT IN (SELECT V.[ValueStr] FROM dbParameter AS P JOIN dbParameterValue AS V	ON P.[ID] = V.[ParameterID]	AND P.[Name] = 'AssetList' AND P.[Class] = 'Dupl. Exception' AND P.[Type] = 'SerialNo' WHERE V.[DTDeletion] IS NULL)"

    'KHi 2007-01-29: Orginal code where retired assets will not show up with dummies
    'strSQL = strSQL + "AND [cf_HP_AssgnRead] <> 'RETIRED ASSET' AND [" & strSearch & "] <> '' "
'    strSQL = strSQL + "AND [" & strSearch & "] <> '' "
'    strSQL = strSQL + "GROUP BY [" & strSearch & "] HAVING (COUNT(*) > 1 AND COUNT(*) <= " & intMaxRows & ") ORDER BY [" & strSearch & "]"
  Else
    strSQL = "SELECT [ComputerName] FROM RadiaRIMProd.dbo.webDuplicateAssetsOnComputerName WHERE [ComputerName] <> '' "
    If UCase(Session("CountryAccess")) <> "ALL" Then
      strSQL = strSQL + "AND (CHARINDEX([LocationCountryCode], '" & session("CountryAccess") & "') <> 0 OR [LocationCountry] = '' OR [LocationCountryCode] IS NULL) "
    End if
    strSQL = strSQL + "AND [ComputerName] NOT IN (SELECT ISNULL([NewAssetName], '') AS 'Name' FROM RadiaRIMProd.dbo.acMACDActual "
    strSQL = strSQL + "UNION SELECT ISNULL([OldAssetName],'') AS 'Name' FROM RadiaRIMProd.dbo.acMACDActual) "
    strSQL = strSQL + "GROUP BY [ComputerName] "
    strSQL = strSQL + "HAVING (COUNT(*) > 1 AND COUNT(*) <= " & intMaxRows & ") "
    strSQL = strSQL + "ORDER BY [ComputerName]"
  End If

  CreateArray StrSearch, strSQL, ArrDupli
  
' ******************** SELECT
Case "SELECT"
	strList = Request("strList")
	NextAction = "Save"

  ' Check if the selected assetname or serial number exist in the acMACD database. If so, set a variable for showing an alert.
  If StrSearch = "computername" Then
    strSQL = " SELECT NewAssetname, OldAssetname FROM RadiaRIMProd.dbo.acMACDActual WHERE ((NewAssetname = '" & strItem & "') OR (OldAssetname = '" & strItem & "'))"
    strTmp = "computername"
  Else
    strSQL = " SELECT NewSerialNr, OldSerialNr FROM RadiaRIMProd.dbo.acMACDActual WHERE ((NewSerialNr = '" & strItem & "') OR (OldSerialNr = '" & strItem & "'))"
    strTmp = "serial number"
  End If
  
  objRs.Open strSQL, objConn
  
  If Not objRs.EOF Then 
    strMessage = "<strong>The " & strTmp & " you selected (" & strItem & ") exist in a pending change as well.\nWhen you continue this existing record will be ignored.</strong>"
  End If
  
  objRs.Close

  If StrSearch = "computername" Then
    strSQL = " SELECT TOP " & intMaxRows & " * FROM RadiaRIMProd.dbo.webDuplicateAssetsOnComputerName WHERE [" & strSearch & "]='" & strItem & "'"
  Else
    strSQL = " SELECT TOP " & intMaxRows & " * FROM RadiaRIMProd.dbo.acAssetList WHERE [" & strSearch & "]='" & strItem & "'"
  End If
  ' Get duplicate records of selected assetname or serialnumber
  objRs.open strSQL, objConn
  intColumn = -1
  CoreRecord = 0
  strTmp2 = Date
  While Not objRs.EOF
    intColumn = intColumn + 1
    RecordID(intColumn) = " computername='" & objRs("ComputerName") & "' and serialno='" & objRs("SerialNo") & "' "

    strTmp = objRs("DTLastScan")
    If Len(strTmp) > 2 Then 
      strTmp = Year(strTmp) & "-" & Month(strTmp) & "-" & Day(strTmp)
    End If
    DTLastScan(intColumn) = strTmp

    NewAssetName(intColumn) = objRs("ComputerName")
    NewSerialNr(intColumn) = objRs("SerialNo")
    AssetTag(intColumn) = objRs("AssetTag")
    CountryOfLocation(intColumn) = objRs("LocationCountry")
    LocationOfAsset(intColumn) = objRs("LocationName")
    LocDetail(intColumn) = objRs("fv_SLDE_LocDetail")

    strTmp = objRs("DTInstall")
    If Len(strTmp) > 2 Then 
      strTmp = Year(strTmp) & "-" & Month(strTmp) & "-" & Day(strTmp)
    End If
    InstallDate(intColumn) = strTmp
    
    Opco(intColumn) = objRs("fv_SLDE_BUL") & " " & objRs("fv_SLDE_Opco")
    CostLoc(intColumn) = objRs("CostCenterTitle")

    strTmp = objRs("DTAcquisition")
    If Len(strTmp) > 2 Then 
      strTmp = Year(strTmp) & "-" & Month(strTmp) & "-" & Day(strTmp)
    End If
    PurchaseDate(intColumn) = strTmp
    
    Invoicetype(intColumn) = objRs("fv_SLDE_BillingStatus")
    Status(intColumn) = objRs("cf_HP_assgnRead")

    If objRs("ScannerDesc") = "No Radia" Then
      Comment(intColumn) = "No"
    Else
      Comment(intColumn) = "Yes"
    End If
    
    Category(intColumn) = objRs("CategoryName")
    Brand(intColumn) = objRs("Brand")
    Model(intColumn) = objRs("ProductModel")
    ProductID(intColumn) = objRs("ProductBarCode")
    NTLogon(intColumn) = objRs("SupervisorUserLogin")
    UserLName(intColumn) = objRs("SupervisorName")
    UserFName(intColumn) = objRs("SupervisorFirstName")
    UserPhone(intColumn) = objRs("Supervisorphone")
    UserEMail(intColumn) = objRs("SupervisorEmail")
    InternalTag(intColumn) = objRs("InternalTag")
    
    If ISDate(PurchaseDate(intColumn)) And DateDiff("d", PurchaseDate(intColumn), strTmp2) > 0 Then
      strTmp2 = PurchaseDate(intColumn)
      CoreRecord = intColumn
    End If
    objRs.MoveNext
  Wend
  objRs.Close

' ******************** SAVE
Case "SAVE"
	strList = Request("strList")
	NextAction = "Summary"
	
	Randomize
  CurrentDate = Year(Date) & "-" & Month(Date) & "-" & Day(Date)& " " & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time) & "." & Int(1000 * Rnd)
'                   "Category /Brand/Model/Product ID/Serial number/Asset name  /Location       /Location detail/Country          /Network logon/Status/Cost location/Invoice type/Install date/OpCo name                                                 /Purchase date/Contract reference/Radia on/Last name    /First name    /Phone number/E-mail address"
'  StrFormStrings = "Category |Brand|Model|Product ID|Serial Number|AssetName   |Location       |Location Detail|Country          |Network Logon|Status|Cost Location|Invoice Type|Install Date|OpCo name|OldSerialNr|OldAssetName|OldStatus|OldBillStatus|Purchase Date|Contract Reference|Radia On|User Lastname|User Firstname|User's Phone|Supervisor E-mail"
'  StrTableStrings = "Category|Brand|Model|ProductID |NewSerialnr  |NewAssetname|locationOfAsset|DetailLocation |CountryOfLocation|NTLogon      |Status|CostLoc      |InvoiceType |InstallDate |Opco     |OldSerialNr|OldAssetName|OldStatus|OldBillStatus|PurchaseDate |Comment |UserLName    |UserFName     |UserPhone   |UserEmail"

'                   "Category /Brand/Model/Product ID/Serial number/Asset name  /Location       /Location detail/Country          /Network logon/Status/Cost location/Invoice type/Install date/OpCo name                                                 /Purchase date/Contract reference/Radia on/Last name    /First name    /Phone number/E-mail address"
'  StrFormStrings = "Category |Brand|Model|Product ID|Serial number|Asset name  |Location       |Location Detail|Country          |Network Logon|Status|Cost Location|Invoice Type|Install Date|OpCo name|OldSerialNr|OldAssetName|OldStatus|OldBillStatus|Purchase date|Contract reference|Radia on|Lastname     |Firstname     |Phone number|E-mail address"
'  StrTableStrings = "Category|Brand|Model|ProductID |NewSerialnr  |NewAssetname|locationOfAsset|DetailLocation |CountryOfLocation|NTLogon      |Status|CostLoc      |InvoiceType |InstallDate |Opco     |OldSerialNr|OldAssetName|OldStatus|OldBillStatus|PurchaseDate |Comment |UserLName    |UserFName     |UserPhone   |UserEmail"


  strFormStrings = "Category|Brand|Model|Product ID|Serial number|Asset name|Location|Location Detail|Country|Network Logon|Status|Cost Location|Invoice Type|Install Date|OpCo name|OldSerialNr|OldAssetName|OldStatus|OldBillStatus|Purchase date|Radia on|Last name|First name|Phone number|E-mail address"
  strTableStrings = "Category|Brand|Model|ProductID|NewSerialNr|NewAssetname|locationOfAsset|DetailLocation|CountryOfLocation|NTLogon|Status|CostLoc|InvoiceType|InstallDate|Opco|OldSerialNr|OldAssetName|OldStatus|OldBillStatus|PurchaseDate|RadiaOn|UserLName|UserFName|UserPhone|UserEmail"
  strSummaryStrings = "Asset name|Serial number|Asset tag|Country|Location|Location detail|Install date|OpCo name|Cost location|Purchase date|Invoice type|Status|Radia on|Category|Brand|Model|Network Logon|First name|Last name|Phone number|E-mail address"

  arrFormStrings = Split(strFormStrings, "|")
  arrTableStrings = Split(strTableStrings, "|")
  arrSummaryStrings = Split(strSummaryStrings, "|")
  For intCount = 0 to intColumn
		'Build the list of fields to save in IMACD table
    strSQL = "INSERT INTO acMACD ("
    For intRow = 0 to UBound(arrTableStrings)
      strSQL = strSQL & arrTableStrings(intRow) & ", "
    Next
    strSQL = strSQL & " TypeRequest, EWM, OnsiteEng, AssetTag, InternalTag) " & vbCrLf & " VALUES ("
    For intRow = 0 to UBound(arrFormStrings)
      If intCount = int(CoreRecord) Then
        strSQL = strSQL & "'" & CstrSQL(Request(arrFormStrings(intRow))) & "', "
      Else
        strSQL = strSQL & "'" & CstrSQL(Request(arrFormStrings(intRow) & intCount)) & "', "
      End If
    Next

    If intCount = int(CoreRecord) Then
      strSQL = strSQL & " 'UPDATE', 'Saved by " & session("UID") & "@" & CurrentDate & "', '" & session("UID") & "', '" & request("Asset tag") & "', '" & request("Internal tag" & intCount) & "' ) "
    Else
      strSQL = strSQL & " 'REMOVE', 'Removed by " & session("UID") & "@" & CurrentDate & "', '" & session("UID") & "', '" & request("Asset tag") & "-DEL', '" & request("Internal tag" & intCount) & "') "
    End if
	  objConn.Execute strSQL
    
    strSQL = "SELECT @@IDENTITY AS 'ChangeID'"
    objRs.Open strSQL, objConn
    
    If Not objRs.EOF Then
      iChangeID = objRs("ChangeID")
    End If
    objRs.Close
    
    strSQL = "UPDATE RadiaRIMProd.dbo.webAssetList SET [ChangeID]=" & iChangeID &_ 
      " WHERE [InternalTag]='" & request("Internal tag" & intCount) & "'"
	  objConn.Execute strSQL
  Next
Case "SUMMARY"
	strList = Request("strList")
	strItem = Request("Item")
	NextAction = "Select"
Case Else
End Select
%>

<%
' ***************************
'  Start Processing the page
' ***************************

' If a message string has been set show it in a javascript message box
If strMessage <> "" then
  response.write "<script type=""text/javascript"">alert('" & strMessage & "');</script>"
End if

' *****************
'  Javascript code
' *****************
%>

<%
' *************************************
'  Load Header template
' *************************************
%>
   <!--#include file="Include/pageHeader.asp"-->

<script type="text/javascript">
  function selectuser(column) {
    var x = document.form;
    x['Network\ logon'][column].checked = true;
    x['Last\ name'][column].checked = true;
    x['First\ name'][column].checked = true;
    x['Phone\ number'][column].checked = true;
    x['E\-mail\ address'][column].checked = true;
  }
</script>

<%

' *****************************************
'  Create Page depending on the PageAction
' *****************************************
%>

<%
Select Case UCase(PageAction)
Case ""
%>
  <form name="form" method="post" action="">
    <input type="hidden" name="PageAction" value="<%=NextAction %>">
    <table id="maintable" width="970" align="center" border="0" cellpadding="5" cellspacing="0">
      <tr>
        <td height="20" width="165" />
        <td width="400">
          <span id="alert" class="alert">
    
      		  Note: Enter your change requests before 15:30 CET to see them tomorrow.
      		   To be sure, assets cannot be retired/put on stock before 15:30 if they have connected to the network the same day.
    
          </span>
        </td>
        <td />
      </tr>
      <tr>
        <td height="5" colspan="3" />
      </tr>
      <tr>
        <td width="165" />
        <td height="20" class="menutitle" width="400">Select search criteria</td>
        <td />
      </tr>
      <tr>
        <td /><td height="60" width="400">
          <input type="radio" name="search" value="serialno" class="radioinput" onclick="document.form.submit.disabled=false" checked /> Duplicate serial numbers<br />
          <input type="radio" name="search" value="computername" class="radioinput" onclick="document.form.submit.disabled=false" /> Duplicate asset names<br />
        </td>
        <td />
      </tr>
      <tr>
        <td />
        <td height="25">
          <input type="submit" name="submit" value="Search" class="searchbtn" />
        </td>
        <td />
      </tr>
      <tr>
        <td />
      </tr>
    </table>
  </form>
<!-- ================================================================================== -->
<!--                                                                                    -->
<!-- SEARCH                                                                             -->
<!--                                                                                    -->
<!-- ================================================================================== -->
<% Case "SEARCH" %>
  <form name="form" method="post" action="">
    <input type="hidden" name="PageAction" value="<%=NextAction %>">
    <input type="hidden" name="search" value="<%=request("search")%>" />
    <input type="hidden" name="strList" value="<%=strList %>" />
    <input type="hidden" name="SQL" value="<%=strSQL %>" />
    <table id="maintable" width="970" align="center" border="0" cellpadding="5" cellspacing="0">
      <tr>
        <td height="75" colspan="3">
        </td>
      </tr>
      <tr>
        <td width="165" />
        <td height="20" class="menutitle" width="400">
<%
          If StrSearch = "serialno" then
            response.write "Select serial number"
          Else
            response.write "Select asset name"
          End If
%>
        </td>
        <td />
      </tr>
      <tr>
        <td />
        <td height="60" width="400">
          <select name="Item" class="textbox280">
<%
            ' Show all options to select from
            If strSearch = "serialno" then
              response.write "<option value=""NULL"" selected>Select serial number </option>"
            Else
              response.write "<option value=""NULL"" selected>Select asset name </option>"
            End if
            For intCount = 1 To ubound(ArrDupli)
              response.write "<option value=""" & ArrDupli(intCount) & """>" & ArrDupli(intCount) & "</option>" & VbCrLf
            Next
%>
          </select>
          <br />
<%
          If UCase(Session("CountryAccess")) = "ALL" Then 
            Response.Write "Showing all countries" 
          Else 
            Response.Write "Showing only from " & Replace(Session("CountryAccess"), "." , " - ") & " and items without country"
          End If
%>
        </td>
        <td><div id="Helper"></div></td>
      </tr>
      <tr>
        <td />
        <td height="25">
          <input type="submit" name="submit" value="Search" class="searchbtn" onclick="if (document.form.Item.value != 'NULL') {return true} else {return false}">
          <script type="text/javascript">
            document.form.Item.focus();
          </script>
        </td>
        <td />
      </tr>
      <tr>
        <td />
      </tr>
    </table>
    <p><%=strList %></p>
  </form>

<!-- ================================================================================== -->
<!--                                                                                    -->
<!-- SELECT                                                                             -->
<!--                                                                                    -->
<!-- ================================================================================== -->

<% case "SELECT" %>
  <form name="form" method="post" action="acDuplicates.asp">
    <input type="hidden" name="PageAction" value="<%=NextAction %>" />
    <input type="hidden" name="Item" value="<%=request("Item")%>" />
    <input type="hidden" name="search" value="<%=request("search")%>" />
    <input type="hidden" name="strList" value="<%=strList %>" />
    <input type="hidden" name="intColumn" value="<%=intColumn%>" />
    <input type="hidden" name="corerecord" value="<%=CoreRecord%>" />
    <input type="hidden" name="OldAssetName" value="<%=NewAssetName(CoreRecord)%>" />
    <input type="hidden" name="OldSerialNr" value="<%=NewSerialNr(CoreRecord)%>" />
    <input type="hidden" name="OldStatus" value="<%=Status(CoreRecord)%>" />
    <input type="hidden" name="OldBilltype" value="<%=InvoiceType(CoreRecord)%>" />
    <table id="selecttable" width="970" valign="top" align="center" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="15" colspan="2"/>
      </tr>
      <tr>
        <td>
          <% If UCase(StrSearch) = "SERIALNO" Then %>
            Showing all assets where serial number equals "<%=strItem %>"
          <% Else %>
            Showing all assets where asset name equals "<%=strItem %>"
          <% End If %>
        </td>
        <td align="right">
          <input type="submit" onclick="document.form.PageAction.value = ''; {return true};" name="nextbtn" value="Menu" class="SmallBtn" />
					<% 
						strNextPage = NextItem()
						If strNextPage <> "" Then
					%>
	          <input type="submit" onclick="document.form.PageAction.value = 'Select'; document.form.Item.value = '<%=strNextPage %>'; {return true};" name="nextbtn" value="Next" class="SmallBtn" />
					<%
						End If
					%>
          <input type="submit" onclick="if ( confirm('Save changes and delete duplicates?') ) {return true} else {return false}" name="savebtn" value="Save" class="SmallBtn" />
        </td>
      </tr>
      <tr>
        <td height="10" colspan="2">
        </td>
      </tr>
      <tr>
        <td colspan="2" valign="top">
          <div id="scrollbox" class="scrollbox">
<!--
            <script type="text/javascript">
              document.getElementById('scrollbox').style.height=document.body.clientHeight-176
            </script>
-->
            <table class="duplitable" border="1" bordercolor="#ffffff" cellpadding="0" cellspacing="0">
              <!-- header of the table -->
              <%
                response.write "<tr>" & vbCrLf 
                response.write "                <td width=""140"" />" & vbCrLf
                For intCount = 0 to intColumn
                  If intCount = CoreRecord Then 
                    colorid="bgcolor1" 
                  Else 
                    colorid="bgcolor2"
                  End If
                  response.write "                <td width=""50"" class=""" & ColorID & """>&nbsp;</td>" & vbCrLf
                  response.write "                <th align=""left"" class=""" & ColorID & """>"
    
                  If intCount = CoreRecord Then
                    response.write "CORE"
                  Else
                    response.write "DUPLICATE"
                  End If
                  response.write "</th>" & vbCrLf
                Next
                response.write "              </tr>" & vbCrLf 
                '<tr height=""1""><td colspan=" & intColumn * 2 - 1 & "></td></tr>"
                
                If strSearch = "computername" Then
                  CreateDisabledRow "Asset name", NewAssetName, ""
                  CreateRow "Serial number", NewSerialNr, ""
                Else
                  CreateRow "Asset name", NewAssetName, ""
                  CreateDisabledRow "Serial number", NewSerialNr, ""
                End If
                CreateDisabledRow "Date last scan", DtLastScan, "alert"
                CreateRow "Asset tag", AssetTag, ""
                CreateDisabledRowPlus "Internal tag", InternalTag, ""
                CreateRow "Country", CountryOfLocation, ""
                CreateRow "Location", LocationOfAsset, ""
                CreateRow "Location detail", LocDetail, ""
                CreateRow "Install date", InstallDate, ""
                CreateRow "OpCo name", Opco, ""
                CreateRow "Cost location", CostLoc, ""
                CreateRow "Purchase date", PurchaseDate, "alert"
                CreateRow "Invoice type", InvoiceType, ""
                CreateRow "Status", Status, ""
                CreateRow "Radia on", Comment, ""
                CreateRow "Category", Category, ""
                CreateRow "Brand", Brand, ""
                CreateRow "Model", Model, ""
                CreateRow "Product ID", ProductID, ""

                ' Show Network Logon with javascript onclick event for selecting Username/phone/email
                CreateOnclickRow "Network logon", NTLogon
                CreateOnclickRow "First name", UserFName
                CreateOnclickRow "Last name", UserLName
                CreateOnclickRow "Phone number", UserPhone
                CreateOnclickRow "E-mail address", UserEMail
              %>
            </table>
          </div>
        </td>
      </tr>
    </table>
  </form>
  
<!-- ================================================================================== -->
<!--                                                                                    -->
<!-- SAVE                                                                               -->
<!--                                                                                    -->
<!-- ================================================================================== -->

<% Case "SAVE" %>
	<form name="form" method="post" action="acDuplicates.asp">
		<input type="hidden" name="PageAction" value="<%=NextAction %>" />
		<input type="hidden" name="Item" value="<%=Request("Item") %>" />
    <input type="hidden" name="strList" value="<%=strList %>" />
    <input type="hidden" name="search" value="<%=request("search")%>" />
		<table id="savetable" width="970" align="center" border="0" cellpadding="1" cellspacing="1">
			<td hight="15" width="300"> </td>
			<td width="300"> </td>
			<td width="243"> </td>
			<tr>
				<th>
				You saved the following data:
				</th>
				<td align="right">
					<input type="submit" onclick="document.form.PageAction.value = ''; {return true};" name="nextbtn" value="Menu" class="SmallBtn" />
					<% 
						strNextPage = NextItem()
						If strNextPage <> "" Then
					%>
						<input type="submit" onclick="document.form.PageAction.value = 'Select'; document.form.Item.value = '<%=strNextPage %>'; {return true};" name="nextbtn" value="Next" class="SmallBtn" />
					<%
						End If
					%>
				</td>
				<td witdh="273" />
			</tr>
			 <%
				For intCount = 0 to UBound(arrSummaryStrings)
					CreateSummaryRow(arrSummaryStrings(intCount))
				Next
			 %>
			<tr>
			 <td colspan="2">
				<br />This record is saved and is pending to be merged into the main database.
			 </td>
			 <td />
			</tr>
		</table>
	</form>
<% End Select %>
<%
AcCloseDB()
%>
<!--#include file="Include/pageFooter.asp"-->
<%
' ****************************************
'  DONE, Closed page with Footer Template
' ****************************************
%>

</html>