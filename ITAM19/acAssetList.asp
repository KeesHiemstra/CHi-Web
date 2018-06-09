<%@ language="VBSCRIPT" %>
<%
' File: acAssetList
' Version: 2.10 (Kees Hiemstra)
' - Solved the problem with removing when there are more than 2 columns in the page.
%>
<% Option Explicit%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.buffer = True%>
<%Response.Expires = 0%>
<!--#include file="include/globalVars.asp"-->
<!--#include file="include/dbFunction.asp"-->
<!--#include file="include/strFunction.asp"-->
<%
' Declare variables
Dim strSQL
Dim strCount
Dim strQuery
Dim strSearch
Dim strSortOrder
Dim bErrorBeforeAdding
Dim bAlreadyExist
Dim strSortOrderField
Dim strCategory
Dim strTmp
Dim aCountries, strCountry, strEngineers

'Define variable values

' If this page is opened directly (PageAction is empty), then set it's action to "INFO"
If Session("PageAction") = "" Then Session("PageAction") = "INFO"
If Request("PageAction") = "FromVBS" Then Session("PageAction") = "ADD"

Select Case UCase(Session("PageAction"))
  Case "EDIT"
    PageTitle = "Edit Asset attributes"
  Case "DELETE"
    PageTitle = "Cancel Asset modification"
  Case "INFO"
    PageTitle = "View Asset information"
  Case "ADD"
    PageTitle = "Add new asset"
  Case else
    PageTitle = "Unknown page!!!"
End Select
MenuSize = "Large"

' Check if the user is logged otherwise redirect to the login page
If Session("UID") = "" Then
  Session("AfterLoginGoto") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString
  Response.Redirect "Logon.asp"
  Response.End
End if

' PAGE FLOW
' This page has several functions. The function or action is stored in the session variable named PageAction.
' Using a select case statement the PageAction is filtered and all actions before building the page is done within that statement.

Call acOpenDB()

Select Case UCase(Session("PageAction"))
' Depending on the main action of the page the following actions needs to be executed.
  Case "EDIT", "INFO"
    If (Session("ClickedRow") = "Yes") or (Request("ClickedRow") = "Yes") Then
      Session("ChangeID") = ""
      Request("ChangeID") = ""
      Session("ClickedRow") = ""
      Request("ClickedRow") = ""
    End If

    If Request("ChangeID") <> "" Then
      ' The user clicked on of the EDIT or INFO buttons.
      ' Redirect the user directly to the assetdetails page. On that page the required data is loaded again from the database.
      Session("ChangeID") = Request("ChangeID")
      acCloseDB()
      Response.Redirect "acIMACD.asp"
                      
      Response.End
    Else
      strEngineers = "'" & Session("UID") & "'"
      If (Session("SecHP") => 15) And (UCase(Session("CountryAccess")) <> "ALL") Then
        aCountries = Split(UCase(Session("CountryAccess")), ",")
        For Each strCountry In aCountries
          If strCountry <> "" Then strEngineers = strEngineers & ",'HPInstall_" & strCountry & "'"
        Next
      Else
      End If
      ' In all other situations the user refined or changed the query by entering filterdata or changing sort order
      ' Check if the entered query gives some results in the acmacd table first
      
      strCount = "SELECT COUNT(*) AS Frequency "

      strQuery = "SELECT  TOP " & Session("SearchResult") & " udHRDataOpco.[Name] AS 'OpcoName', acMACD.* "

      strSearch = " FROM RadiaRIMProd.dbo.acMACD LEFT JOIN RadiaRIMProd.dbo.udHRDataOpCo ON udHRDataOpco.[Display] = LEFT(acMACD.Opco, 4) WHERE [OnSiteEng] IN (" & strEngineers & ") "
      strSearch = strSearch & " AND NewAssetName LIKE '" & Trim(Request("AssetName")) & "%' "
      strSearch = strSearch & " AND NewSerialNr LIKE '" & Trim(Request("SerialNo")) & "%' "
      strSearch = strSearch & " AND UserFName LIKE '" & Trim(Request("FirstName")) & "%' "
      strSearch = strSearch & " AND UserLName LIKE '" & Trim(Request("LastName")) & "%' "
      strSearch = strSearch & " AND (TransFlag IS NULL OR TransFlag <> 'Y') "
      strSearch = strSearch & " AND TypeRequest <> 'IGNORE' "

      objRs.Open strCount & strSearch, objConn

      If objRs("Frequency") = "0" Then
        ' If the user did not click one of the column titles to sort, then sort by AssetName.
        If request("SortOrder") = "" Then
          strSortOrderField = "[DTLastScan] DESC"
        Else
          strSortOrderField = Request("SortOrder")
        End If

        ' No items found in the acmacd table, continue searching in the acassetlist table and show only the first 100 records.
        objRs.Close

        'Join acassetlist and udHRDatacountry to add the column countrycode. This code is used for filtering the country permissions (countryaccess)
        '2006-05-08, Kees Hiemstra, JOIN with udHRDataCountry lead to duplicates that are no duplicates at all.
        strSQL = "SELECT TOP " & Session("SearchResult") & " AL.[InternalTag], AL.[ComputerName], AL.[SerialNo], AL.[SupervisorFirstName], AL.[SupervisorName], AL.[ProductModel], AL.[fv_SLDE_BUL], AL.[fv_SLDE_Opco], AL.[cf_HP_AssgnRead], AL.[fv_SLDE_BillingStatus] FROM acAssetList AS AL "
        Select Case Trim(UCase(Request("SerialNo")))
          Case "::DISPOSE" 
						'Search function to find assets to be disposed, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (([ComputerName] LIKE '" & Request("AssetName") & "%') "
            strSQL = strSQL & "AND AL.[fv_SLDE_BillingStatus] = 'Not in contract, to be disposed' "
            strSQL = strSQL & "AND [SupervisorName] LIKE '" & Trim(Request("LastName")) & "%' AND [SupervisorFirstName] LIKE '" & Trim(Request("FirstName")) & "%' "
          Case "::DUMMY"
						'Search function to find dummy records, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (([ComputerName] LIKE '" & Trim(Request("AssetName")) & "%') "
            strSQL = strSQL & "AND AL.[Brand] = 'RADIA' AND AL.[ProductModel] LIKE '%dummy%' AND AL.[cf_HP_AssgnRead] = 'In use' "
          Case "::EMPTYEMAIL"
						'Search function to find computers in use without email address, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (AL.[CategoryName] IN ('Desktop computer', 'Laptop') "
            strSQL = strSQL & "AND AL.[fv_SLDE_BillingStatus] = 'In contract' AND [cf_HP_AssgnRead] = 'In use' AND AL.[SupervisorEMail] = '' "
          Case "::EMPTYSN" 
						'Search function to find computers in use without serial number, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (([ComputerName] LIKE '" & Trim(Request("AssetName")) & "%') "
            strSQL = strSQL & "AND [SerialNo] = '' AND AL.[cf_HP_AssgnRead] = 'In use' AND AL.[CategoryName] IN ('Desktop computer', 'Laptop') "
            strSQL = strSQL & "AND [SupervisorName] LIKE '" & Trim(Request("LastName")) & "%' AND [SupervisorFirstName] LIKE '" & Trim(Request("FirstName")) & "%' "
          Case "::EMPTYCOUNTRY"
						'Search function to find computers in use without user, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (AL.[CategoryName] IN ('Desktop computer', 'Laptop') "
            strSQL = strSQL & "AND AL.[fv_SLDE_BillingStatus] = 'In contract' AND AL.[LocationCountry] = '' "
          Case "::EMPTYUSER"
						'Search function to find computers in use without user, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (AL.[CategoryName] IN ('Desktop computer', 'Laptop', 'Folio Laptop') "
            strSQL = strSQL & "AND AL.[fv_SLDE_BillingStatus] IN ('In contract', 'In contract, owned by Sara Lee') AND [cf_HP_AssgnRead] = 'In use' AND AL.[SupervisorUserLogin] = '' "
          Case "::ERRORBILLING" 
						'Search function to find computers and printers that can't be invoiced, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (([ComputerName] LIKE '" & Trim(Request("AssetName")) & "%') "
            strSQL = strSQL & "AND AL.[cf_HP_AssgnRead] = 'In use' AND (AL.[CategoryName] IN ('Desktop computer', 'Laptop', 'Folio Laptop') OR AL.[CategoryName] LIKE '%printer%') "
            strSQL = strSQL & "AND (AL.[fv_SLDE_BillingStatus] NOT IN ('In contract', 'Not in contract, used by HP', 'In contract, owned by Sara Lee', 'Not in contract, used for training', 'In contract, for refresh') OR AL.[AssetOpCoID] IN (SELECT [OpCoID] FROM udHRDataOpCo WHERE [InvoiceStatus] = 0 AND OpCoID <> 9000 AND OpCoID <> 0)) "
            strSQL = strSQL & "AND [SupervisorName] LIKE '" & Trim(Request("LastName")) & "%' AND [SupervisorFirstName] LIKE '" & Trim(Request("FirstName")) & "%' "
          Case "::ERRORDISPOSE" 
						'Search function to find assets to be disposed, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (([ComputerName] LIKE '" & Request("AssetName") & "%') "
            strSQL = strSQL & "AND AL.[fv_SLDE_BillingStatus] = 'Not in contract, to be disposed' AND AL.[cf_HP_AssgnRead] != 'In stock' "
            strSQL = strSQL & "AND [SupervisorName] LIKE '" & Trim(Request("LastName")) & "%' AND [SupervisorFirstName] LIKE '" & Trim(Request("FirstName")) & "%' "
          Case "::ERRORCONTRACT" 
						'Search function to find computers and printers that can't be invoiced, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (([ComputerName] LIKE '" & Trim(Request("AssetName")) & "%') "
            strSQL = strSQL & "AND (AL.[CategoryName] IN ('Desktop computer', 'Laptop', 'Folio Laptop') OR AL.[CategoryName] LIKE '%printer%') "
            strSQL = strSQL & "AND (AL.[fv_SLDE_BillingStatus] IN ('Not in contract')) "
            strSQL = strSQL & " "
          Case "::ERROROPCO" 
						'Search function to find computers and printers that can't be invoiced, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.udHRDataOpCo AS O ON AL.[AssetOpCoID] = O.[OpCoID] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (([ComputerName] LIKE '" & Trim(Request("AssetName")) & "%') "
            strSQL = strSQL & "AND AL.[cf_HP_AssgnRead] != 'Retired (or consumed)' AND (AL.[CategoryName] IN ('Desktop computer', 'Laptop', 'Folio laptop', 'Tablet') OR AL.[CategoryName] LIKE '%printer%') "
            strSQL = strSQL & "AND (AL.[fv_SLDE_BillingStatus] IN ('In contract') AND O.[HPManaged] = 0) AND AL.[AssetOpCoID] != 0 "
          Case "::ERROROPCOCOUNTRY" 
						'Search function to find computers and printers that can't be invoiced, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.udHRDataOpCo AS O ON AL.[AssetOpCoID] = O.[OpCoID] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (([ComputerName] LIKE '" & Trim(Request("AssetName")) & "%') "
            strSQL = strSQL & "AND AL.[cf_HP_AssgnRead] != 'Retired (or consumed)' AND (AL.[CategoryName] IN ('Desktop computer', 'Laptop', 'Folio Laptop', 'Tablet') OR AL.[CategoryName] LIKE '%printer%') "
            strSQL = strSQL & "AND (AL.[fv_SLDE_BillingStatus] IN ('In contract', 'In contract, owned by Sara Lee')) AND AL.[AssetOpCoID] = 2 "
            strSQL = strSQL & "AND (AL.[LocationCountryCode] != O.[CountryCode]) "
          Case "::ERRORIMACD"
						'Search function to find assets where the IMACD is rejected, records with unsend IMACD are not shown
            strSQL = strSQL & "JOIN acImportIMACDError ON AL.[InternalTag] = acImportIMACDError.[InternalTag] "
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (([ComputerName] LIKE '" & Trim(Request("AssetName")) & "%') "
            strSQL = strSQL & "AND [SupervisorName] LIKE '" & Trim(Request("LastName")) & "%' AND [SupervisorFirstName] LIKE '" & Trim(Request("FirstName")) & "%' "
          Case "::ERRORLOCATION"
						'Search function to find assets where the location is the same as the country, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
	          strSQL = strSQL & "JOIN ITAMData.dbo.riCountry AS C ON AL.[LocationCountryCode] = C.[CountryCode] AND C.[Active] = 1 "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (([ComputerName] LIKE '" & Trim(Request("AssetName")) & "%') "
            strSQL = strSQL & "AND [SupervisorName] LIKE '" & Trim(Request("LastName")) & "%' AND [SupervisorFirstName] LIKE '" & Trim(Request("FirstName")) & "%' "
	          strSQL = strSQL & "AND AL.[LocationName] = C.[CountryName] AND AL.[CategoryName] IN ('Desktop computer', 'Laptop', 'Folio Laptop') "
	          strSQL = strSQL & "AND AL.[fv_SLDE_BillingStatus] NOT IN ('In contract, for refresh', 'Obsolete', 'Not in contract, to be disposed') "
			  strSQL = strSQL & "AND AL.[cf_HP_AssgnRead] = 'In use' "
	          strSQL = strSQL & "AND AL.[LocationCountryCode] NOT IN ('HK') "
						strSortOrderField = "AL.[DTMutation], AL.[DTAcquisition] DESC, C.[CountryName], AL.[SerialNo] "
          Case "::ERROROBSOLETE" 
						'Search function to find obsolete assets that are not retired, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (([ComputerName] LIKE '" & Trim(Request("AssetName")) & "%') "
						strSQL = strSQL & "AND AL.[fv_SLDE_BillingStatus] = 'Obsolete' AND AL.[cf_HP_AssgnRead] NOT IN ('Retired (or consumed)', 'Missing') "
            strSQL = strSQL & "AND [SupervisorName] LIKE '" & Trim(Request("LastName")) & "%' AND [SupervisorFirstName] LIKE '" & Trim(Request("FirstName")) & "%' "
          Case "::ERRORUSEDBYHP"
						'Search function to find computers used by HP where the OpCo is not 9000, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (AL.[CategoryName] IN ('Desktop computer', 'Laptop') "
            strSQL = strSQL & "AND AL.[AssetOpCoID] <> 9000 AND AL.[fv_SLDE_BillingStatus] = 'Not in contract, used by HP' "
          Case "::INTTAG"
						'Search function to find specified InternalTag (specified in AssetName)
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE ((AL.[InternalTag] LIKE '" & Trim(Request("AssetName")) & "%') "
          Case "::INSTOCK" 
						'Search function to find assets in stock, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (([ComputerName] LIKE '" & Trim(Request("AssetName")) & "%') "
            strSQL = strSQL & "AND AL.[cf_HP_AssgnRead] = 'In stock' "
          Case "::MISSING" 
						'Search function to find obsolete assets that are not retired, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (([ComputerName] LIKE '" & Trim(Request("AssetName")) & "%') "
						strSQL = strSQL & "AND AL.[cf_HP_AssgnRead] = 'Missing' "
            strSQL = strSQL & "AND [SupervisorName] LIKE '" & Trim(Request("LastName")) & "%' AND [SupervisorFirstName] LIKE '" & Trim(Request("FirstName")) & "%' "
          Case "::NOTONLINE"
						'Search function to find computers that have not been seen by Radia for xxxx months where Radia should be working, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND ([cf_HP_AssgnRead] = 'In use' AND AL.[CategoryName] IN ('Desktop computer', 'Laptop') "
            strSQL = strSQL & "AND AL.[fv_SLDE_BillingStatus] = 'In contract' AND AL.[ScannerDesc] = ''"
						If Request("AssetName") = "" Then
							strSQL = strSQL & "AND AL.[DTLastScan] IS NULL "
						Else
							strSQL = strSQL & "AND DATEDIFF(MONTH, AL.[DTLastScan], GETDATE()) > " & Trim(Request("AssetName")) & " "
						End If
					Case "::OLDCOMPUTER"
						'Search function to find asset that are older than 3 years and 8 months (44 month)
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (([ComputerName] LIKE '" & Trim(Request("AssetName")) & "%') "
						strSQL = strSQL & "AND AL.[CategoryName] IN ('Desktop computer', 'Laptop', 'Thin client') AND AL.[cf_HP_AssgnRead] NOT IN ('Missing', 'Retired (or consumed)') "
						strSQL = strSQL & "AND DATEDIFF(MONTH, AL.[DTAcquisition], GETDATE()) >= 46 "
            strSQL = strSQL & "AND [SupervisorName] LIKE '" & Trim(Request("LastName")) & "%' AND [SupervisorFirstName] LIKE '" & Trim(Request("FirstName")) & "%' "
          Case "::OPCO"
						'Search function to find assets on a specific OpCo
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND ((AL.[fv_SLDE_BUL] LIKE '" & Trim(Request("AssetName")) & "%') "
          Case "::OPCOINUSE"
						'Search function to find computers in use on a specific OpCo
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND ((AL.[fv_SLDE_BUL] LIKE '" & Trim(Request("AssetName")) & "%') "
            strSQL = strSQL & "AND AL.[CategoryName] IN ('Desktop computer', 'Laptop', 'Thin client') AND [cf_HP_AssgnRead] = 'In use' AND AL.[Brand] <> 'MICROSOFT CORPORATION' "
          Case "::TOREFRESH"
						'Search function to find computers that are stated to be refreshed
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
						strSQL = strSQL & "JOIN (SELECT I.[InternalTag] FROM mrWaveItem AS I JOIN (SELECT [InternalTag], MAX(I.[ID]) AS 'ID' "
						strSQL = strSQL & "FROM RadiaRIMProd.dbo.mrWaveItem AS I JOIN RadiaRIMProd.dbo.mrWaveCountry AS C ON I.[CountryCode] = C.[CountryCode] AND I.[WaveID] = C.[WaveID] "
						strSQL = strSQL & "JOIN RadiaRIMProd.dbo.mrWave AS W ON I.[WaveID] = W.[ID] WHERE I.[ToReimage] = 0 AND C.[State] IN ('O', 'P', 'R', 'C') "
						strSQL = strSQL & "AND I.[SLAgreement] = 'Agreed' AND I.[Finished] = 0 GROUP BY [InternalTag]) AS U ON I.[ID] = U.[ID]) AS W "
						strSQL = strSQL & "ON AL.[InternalTag] = W.[InternalTag] "
            strSQL = strSQL & "AND (acMACDActual.[InternalTag] IS NULL "
          Case "::USEDBYHP"
						'Search function to find computers used by HP, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "WHERE (AL.[CategoryName] IN ('Desktop computer', 'Laptop') "
            strSQL = strSQL & "AND AL.[fv_SLDE_BillingStatus] = 'Not in contract, used by HP' AND acMACDActual.[InternalTag] IS NULL "
          Case "::USERINVALID"
						'Search function to find computers in use without a valid user account, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "LEFT OUTER JOIN ITAMNetwork.dbo.adUser ON AL.[SupervisorUserLogin] = adUser.[adAccount] AND adUser.[ADStatus] = 1 "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (AL.[CategoryName] IN ('Desktop computer', 'Laptop', 'Folio Laptop') "
            strSQL = strSQL & "AND AL.[fv_SLDE_BillingStatus] IN ('In contract', 'In contract, owned by Sara Lee') AND [cf_HP_AssgnRead] = 'In use' AND AL.[ComputerName] NOT LIKE 'SL__RPS%' "
            strSQL = strSQL & "AND adUser.[adObjectGUID] IS NULL AND [ComputerName] LIKE '" & Trim(Request("AssetName")) & "%' "
            strSQL = strSQL & "AND [SupervisorName] LIKE '" & Trim(Request("LastName")) & "%' AND [SupervisorFirstName] LIKE '" & Trim(Request("FirstName")) & "%' "
          Case "::USERINVALIDEXCL"
						'Search function to find computers in use without a valid user account, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "LEFT OUTER JOIN ITAMNetwork.dbo.adUser ON AL.[SupervisorUserLogin] = adUser.[adAccount] AND adUser.[ADStatus] = 1 "
            strSQL = strSQL & "JOIN ITAMNetwork.dbo.avComputer AS AV ON AL.[ComputerName] = AV.[ComputerName] AND AV.[DTLastConnect] > GETDATE() - 28 "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (AL.[CategoryName] IN ('Desktop computer', 'Laptop', 'Folio Laptop') "
            strSQL = strSQL & "AND AL.[fv_SLDE_BillingStatus] IN ('In contract', 'In contract, owned by Sara Lee') AND [cf_HP_AssgnRead] = 'In use' AND AL.[ComputerName] NOT LIKE 'SL__RPS%' "
            strSQL = strSQL & "AND adUser.[adObjectGUID] IS NULL AND AL.[ComputerName] LIKE '" & Trim(Request("AssetName")) & "%' "
            strSQL = strSQL & "AND [SupervisorName] LIKE '" & Trim(Request("LastName")) & "%' AND [SupervisorFirstName] LIKE '" & Trim(Request("FirstName")) & "%' AND REPLACE(AL.[SupervisorUserLogin], 'DEMB\', '') != AV.[User] "
          Case "::USERINVALIDNOTEMPTY"
						'Search function to find computers in use without a valid user account, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "LEFT OUTER JOIN ITAMNetwork.dbo.adUser ON AL.[SupervisorUserLogin] = adUser.[adAccount] AND adUser.[ADStatus] = 1 "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (AL.[CategoryName] IN ('Desktop computer', 'Laptop', 'Folio Laptop') "
            strSQL = strSQL & "AND AL.[fv_SLDE_BillingStatus] IN ('In contract', 'In contract, owned by Sara Lee') AND [cf_HP_AssgnRead] = 'In use' AND AL.[ComputerName] NOT LIKE 'SL__RPS%' "
            strSQL = strSQL & "AND adUser.[adObjectGUID] IS NULL AND [ComputerName] LIKE '" & Trim(Request("AssetName")) & "%' AND AL.[SupervisorUserLogin] != '' "
            'strSQL = strSQL & "AND [SupervisorName] LIKE '" & Trim(Request("LastName")) & "%' AND [SupervisorFirstName] LIKE '" & Trim(Request("FirstName")) & "%' "
          Case "::USERINVALIDREFRESH"
						'Search function to find computers in use without a valid user account, records with unsend IMACD are not shown
            strSQL = strSQL & "LEFT OUTER JOIN RadiaRIMProd.dbo.acMACDActual ON AL.[InternalTag] = acMACDActual.[InternalTag] "
            strSQL = strSQL & "LEFT OUTER JOIN ITAMNetwork.dbo.adUser ON AL.[SupervisorUserLogin] = adUser.[adAccount] AND adUser.[ADStatus] = 1 "
            strSQL = strSQL & "WHERE acMACDActual.[InternalTag] IS NULL AND (AL.[CategoryName] IN ('Desktop computer', 'Laptop', 'Folio Laptop') "
            strSQL = strSQL & "AND AL.[fv_SLDE_BillingStatus] IN ('In contract', 'In contract, owned by Sara Lee') AND [cf_HP_AssgnRead] = 'In use' AND AL.[ComputerName] NOT LIKE 'SL__RPS%' "
			strSQL = strSQL & "AND AL.[DTRetirement] <= GETDATE() + 31 "
            strSQL = strSQL & "AND adUser.[adObjectGUID] IS NULL AND [ComputerName] LIKE '" & Trim(Request("AssetName")) & "%' "
            strSQL = strSQL & "AND [SupervisorName] LIKE '" & Trim(Request("LastName")) & "%' AND [SupervisorFirstName] LIKE '" & Trim(Request("FirstName")) & "%' "
          Case Else
            strSQL = strSQL & "WHERE (([ComputerName] LIKE '" & Trim(Request("AssetName")) & "%' "
            strSQL = strSQL & "AND [SupervisorName] LIKE '" & Trim(Request("LastName")) & "%' "
            strSQL = strSQL & "AND [SupervisorFirstName] LIKE '" & Trim(Request("FirstName")) & "%' "
            strSQL = strSQL & "AND [SerialNo] LIKE '" & Trim(Request("SerialNo")) & "%') "
        End Select

        strSQL = strSQL & ") "

        If UCase(Session("CountryAccess")) = "ALL" Then
        ElseIf UCase(Session("CountryAccess")) = "NONE" Then
          strSQL = strSQL & " AND 1=0"
        ElseIf Left(UCase(Session("CountryAccess")), 5) = "OPCO:" Then
          strSQL = strSQL & " AND CHARINDEX(RIGHT('0000' + CAST(ISNULL([AssetOpcoID], 0) AS varchar), 4), '" & Session("CountryAccess") & "') <> 0"
        Else
          strSQL = strSQL & " AND ((AL.[Brand] = 'RADIA' AND AL.[ProductModel] LIKE '%dummy%' AND AL.[cf_HP_AssgnRead] = 'In use') OR CHARINDEX([AssetCountryCode], '" & Session("CountryAccess") & "') <> 0 OR CHARINDEX([LocationCountryCode], '" & Session("CountryAccess") & "') <> 0 OR CHARINDEX(RIGHT('0000' + CAST(ISNULL([AssetOpcoID], 0) AS varchar), 4), '" & Session("CountryAccess") & "') <> 0)"
      	End If
      	
      	If UCase(Session("PageAction")) = "EDIT" And Session("SecHP") <> 255 Then
      		'strSQL = strSQL & "AND AL.[InternalTag] IS NULL "
      	End If

        strSQL = strSQL & " ORDER BY " & strSortOrderField
        objRs.Open strSQL, objConn
        Session("ActiveTable") = "ACASSETLIST"
'				Response.Write(strSQL)
      Else
        ' There are one or more records found in the ACMACD table. Show those records.
        objRs.Close
        ' If the user did not click one of the column titles to sort, then sort by Unique id in descending direction.
        If Request("SortOrder") = "" then
          strSortOrderField = "[ChangeID] DESC"
        Else
          strSortOrderField = Request("SortOrder")
        End If
        strSortOrder = " ORDER BY " & strSortOrderField & " "
        objRs.Open strQuery & strSearch & strSortOrder, objConn
        Session("ActiveTable") = "ACMACD"
        ' Now we have the record set available for viewing, so jump to the creation of the page. :)
      End If
    End If 'Else If Request("ChangeID") <> "" Then
  Case "DELETE"
    Session("ActiveTable") = "ACMACD"
    If Request("ChangeID") <> "" Then
      ' The user clicked one of the deletebuttons buttons.
      strSQL = "UPDATE RadiaRIMProd.dbo.acMACD SET TypeRequest = 'IGNORE', Transflag = 'N' where OnsiteEng = '" & Session("UID") & "' and ChangeID = '" & request("ChangeID") & "'"
      objConn.Execute strSQL
      strSQL = "UPDATE webAssetList SET ChangeID = NULL where ChangeID = '" & request("ChangeID") & "'"
      objConn.Execute strSQL
    End If
    If Request("SortOrder") = "" Then
      strSortOrderField = "NewAssetName"
    else
      strSortOrderField = request("SortOrder")
    End If
    ' When user refined or changed the query by entering filterdata or changing sort order
    strSearch = "select * from RadiaRIMProd.dbo.acMACD where OnsiteEng = '" & Session("UID") & "' "
    strSearch = strSearch & " and NewAssetName like '" & request("AssetName") & "%' "
    strSearch = strSearch & " and NewSerialNr like '" & request("SerialNo") & "%' "
    strSearch = strSearch & " AND UserFName LIKE '" & Request("FirstName") & "%' "
    strSearch = strSearch & " AND UserLName LIKE '" & Request("LastName") & "%' "
    strSearch = strSearch & " and (transflag is null or transflag <> 'Y') "
    strSearch = strSearch & " and Typerequest <> 'IGNORE' "
    strSearch = strSearch & " order by " & strSortOrderField
    objRs.Open strSearch, objConn
  case "ADD"
    bErrorBeforeAdding = False
    if trim(request("AssetName") & request("SerialNo")) <> "" and request("Action") = "Add" then
      ' The user entered an Asset Name and a Serial Number to add to the database
      strError = ValidateText(request("AssetName"), "", "!@#$%^&*()_{}[]|\:;""'<,>?/~`", 4, 15)
      if strError <> "" then
        ' Show popup message if Asset Name is not valid
        response.write "<script>alert('Invalid Asset name:\n" & StrError & "')</script>"
        bErrorBeforeAdding = true
      end if
      StrError = ValidateText(request("SerialNo"), "", "!@#$%^&*()_{}[]|\:;""'<,>.?/~`", 3, 37)
      if StrError <> "" then
        ' Show popup message if Serial Number is not valid
        response.write "<script>alert('Invalid Serial number:\n" & StrError & "')</script>"
        bErrorBeforeAdding = true
      end if
      if not bErrorBeforeAdding then
        ' Check if the Asset Name or the serial Number already exist in the acMACD or acAssetList table.
        strSQL = "select count(*) as frequency from RadiaRIMProd.dbo.acmacdActual where ((NewAssetName = '" & trim(request("AssetName")) & "') or (NewSerialNr = '" & trim(request("SerialNo")) & "')) "
        objRs.Open strSQL, objConn

        if objRs("frequency") <> 0 then bAlreadyExist = true else bAlreadyExist = false ' Check for records in the acMACD table.
        objRs.Close

        if not bAlreadyExist then
          ' If nothing found in the acMACD table, search the acAssetList table.
          strSQL = "select count(*) as frequency from RadiaRIMProd.dbo.acAssetList where ComputerName = '" & trim(request("AssetName")) & "' or serialno = '" & trim(request("SerialNo")) & "'"
          objRs.Open strSQL, objConn
          if objRs("frequency") <> 0 then bAlreadyExist = true ' Check for records in the acAssetList table.
          objRs.Close
        end if

        if bAlreadyExist = true then
          ' Show popup message if Asset Name or Serial Number is found in acMACD
          response.write "<script>alert('Asset name or Serial number already registered in the database.\nPlease verify the Name/Serial or use the edit function.')</script>"
        else
          ' The Asset Name and Serial Number are oke and not in any database.
          ' Create a temporarily record
          session("ActiveTable") = "ACMACD"
          strTmp = year(Date) & "-" & Month(Date) & "-" & day(Date)
          strSQL = "INSERT INTO RadiaRIMProd.dbo.acmacd (Typerequest, NewAssetName, NewSerialNr, EWM, OnsiteEng) values ("
          strSQL = strSQL + "'IGNORE', '" & request("AssetName") & "', '" & request("SerialNo") & "', 'New asset', '" & Session("UID") & "')"
          objConn.Execute strSQL
          objRs.Open "select top 1 ChangeID from acmacd order by ChangeID Desc", objConn
          session("ChangeID") = objRs("ChangeID")

          strSQL = "INSERT INTO RadiaRIMProd.dbo.webAssetList(ComputerName, SerialNo, ChangeID) VALUES ("
          strCategory = request("category")
          if request("Category") = "" then strCategory = "PC"
          strSQL = strSQL & "'" & request("AssetName") & "', '" & request("SerialNo") & "', " & session("ChangeID") & ")"
          objConn.Execute strSQL

          Session("ChangeID") = "ChangeID = " & Session("ChangeID")

          objRs.Close
          acCloseDB()
          response.redirect "acIMACD.asp"
          response.end
        end if
      end if
    end if
  case else
    response.write "<H1>Error in document</H1><BR>Unknown page action (" & session("PageAction") & ")."
    response.end
end select

%>
<!--#include file="include/pageHeader.asp"-->
</script> <span id="alert" class="alert">
  <br />
  <center>
    Note: Enter your change requests before 15.30 CET to see them tomorrow.</center>
  <center>
    To be sure, assets cannot be retired/put on stock before 15:30 if they have connected
    to the network the same day. '<%= Session("ClickedRow") %>'</center>
  <br />
</span>
<table id="maintable" align="center" border="0" cellpadding="0" cellspacing="0" valign="top"
  width="970">
  <tr>
    <td height="15">
    </td>
  </tr>
  <tr>
    <td width="110">
      <form action="" method="get" name="form">
    </td>
    <td width="5">
    </td>
    <td align="right" width="100">
      Asset name:&nbsp;
    </td>
    <td width="150">
      <input name="AssetName" type="text" value="<%=request("AssetName")%>" />
    </td>
    <td align="right" width="100">
      Serial number:&nbsp;
    </td>
    <td width="150">
      <input name="SerialNo" type="text" value="<%=request("SerialNo")%>" />
    </td>
    <td width="100">
      <% if session("PageAction") = "ADD" then %>
        <input name="Action" style="searchbtn" type="submit" value="Add" />
      <% else %>
      <input name="Action" style="searchbtn" type="submit" value="Search" onclick="javascript: document.form.ChangeID.value = '';">
      <% end if %>
    </td>
    <td width="5">
    </td>
    <td align="left" rowspan="2" width="250">
      <table border="0">
        <!-- EXTRA TABLE FOR SHOWING COLORIZED BOXES -->
        <tr>
          <td class="acmacd-bgcolor" width="20">
          </td>
          <td>
            = Your recent changes (cached)</td>
        </tr>
        <tr>
          <td class="acassetlist-bgcolor" width="20">
          </td>
          <td>
            = Actual Asset list</td>
        </tr>
      </table>
    </td>
    <td>
    </td>
  </tr>
  <%
  If session("PageAction") <> "ADD" then
  ' Don't show any list when Page Action is "ADD"
  %>
  <tr>
      <td width="110">
      </td>
      <td width="5">
      </td>
      <td align="right" width="100">
        First name:&nbsp;
      </td>
      <td width="150">
        <input name="FirstName" type="text" value="<%=request("FirstName")%>" />
      </td>
      <td align="right" width="100">
        Last name:&nbsp;
      </td>
      <td width="150">
        <input name="LastName" type="text" value="<%=request("LastName")%>" />
      </td>
      <td width="100">
      </td>
      <td width="5">
      </td>
      <td>
      </td>
    </tr>
    <tr hight="20">
      <td colspan="9">
        &nbsp;</td>
    </tr>
  <tr>
    <td colspan="10" valign="top">
      <p id="scrollbox" class="scrollbox">

        <script>document.getElementById('scrollbox').style.height=document.body.clientHeight-185</script>

        <table border="1" cellpadding="0" cellspacing="0" width="100%">
          <!-- header of the table -->
          <tr align="center" bordercolor="#000000" class="menutitle">
            <td width="40">
              <input name="ChangeID" type="hidden" />
              <input name="ClickedRow" type ="hidden" />
              <input name="SortOrder" type="hidden"`/>
              <%
         select case ucase(session("PageAction"))
          case "EDIT"
              %>
              <img alt="Click below to edit" border="0" src="Image/Edit.png"><%
          case "DELETE"
              %><img alt="Click below to delete" border="0" src="Image/Delete.png"><%
          case "INFO"
              %><img alt="Click below view report" border="0" src="Image/Info.png"><%
         end select
              %>
            </td>
            <% select case session("ActiveTable") %>
            <%  case "ACMACD" %>
            <td>
              <a href="javascript:document.form.SortOrder.value='TypeRequest';document.form.submit();">
                Type Request</a><br>
              <a href="javascript:document.form.SortOrder.value='EWM';document.form.submit();">ChangeRef
                case</a>
            </td>
            <td>
             <a href="javascript:document.form.SortOrder.value='Category';document.form.submit();">
                Category</a><br />
              <a href="javascript:document.form.SortOrder.value='Brand';document.form.submit();">
                Brand</a><br>
              <a href="javascript:document.form.SortOrder.value='Model';document.form.submit();">
                Model</a>
              </td>
              <td>
                <a href="javascript:document.form.SortOrder.value='NewAssetName';document.form.submit();">
                  Computername</a>
              </td>
              <td>
                <a href="javascript:document.form.SortOrder.value='NewSerialNr';document.form.submit();">
                  Serialnumber</a>
              </td>
              <td>
                <a href="javascript:document.form.SortOrder.value='Status';document.form.submit();">
                  Status</a><br>
                <a href="javascript:document.form.SortOrder.value='InvoiceType';document.form.submit();">
                  Billing Status</a>
              </td>
              <td>
                <a href="javascript:document.form.SortOrder.value='UserLName';document.form.submit();">
                  Supervisor</a>
              </td>
              <td>
                <a href="javascript:document.form.SortOrder.value='CountryOfLocation';document.form.submit();">
                  Country</a><br>
                <a href="javascript:document.form.SortOrder.value='Opco';document.form.submit();">
                  Opco</a><br>
              </td>
              <%  case "ACASSETLIST" %>
              <td>
                <a href="javascript:document.form.SortOrder.value='ComputerName';document.form.submit();">
                  Computername</a>
              </td>
              <td>
                <a href="javascript:document.form.SortOrder.value='SerialNo';document.form.submit();">
                  Serialnumber</a>
              </td>
              <td>
                <a href="javascript:document.form.SortOrder.value='SupervisorName';document.form.submit();">
                  Supervisor</a>
              </td>
              <td>
                <a href="javascript:document.form.SortOrder.value='ProductModel';document.form.submit();">
                  Model</a><br>
              </td>
              <td>
                <a href="javascript:document.form.SortOrder.value='fv_SLDE_BUL';document.form.submit();">
                  Opco</a><br>
              </td>
              <td>
                <a href="javascript:document.form.SortOrder.value='cf_HP_AssgnRead';document.form.submit();">
                  Status</a><br>
              </td>
              <td>
                <a href="javascript:document.form.SortOrder.value='fv_SLDE_BillingStatus';document.form.submit();">
                  Billing Status</a><br>
              </td>
              <%  case else %>
              <td>
                ERROR!!</td>
              <% end select %>
          </tr>
          <%
If Not objRs.EOF then
  While not objRs.EOF
    ' Start creating the table rows
    select case session("ActiveTable")
      case "ACMACD"
          %>
          <tr bordercolor="#000000" class="acmacd-color">
            <td align="center">
              <%if ucase(session("PageAction")) = "DELETE" then%>
              <a href="javascript:if (confirm('Are you sure to delete this record?')) { document.form.ChangeID.value='<%=objRs("ChangeID")%>';document.form.submit();}">
                <%else%>
                <a href="javascript:document.form.ChangeID.value=' ChangeID=<%=objRs("ChangeID")%>';document.form.submit();document.form.ClickedRow.Value='Yes';">
                  <%end if%>
                  <%
         select case ucase(session("PageAction"))
          case "EDIT"
                  %>
                  <img alt="Click to edit" border="0" src="Image/Edit.png"><%
          case "DELETE"
                  %><img alt="Click to delete" border="0" src="Image/Delete.png"><%
          case "INFO"
                  %><img alt="Click to view report" border="0" src="Image/Info.png"><%
         end select
                  %>
                </a><a href="javascript:printasset('ChangeID=<%=objRs("ChangeID")%>')">
                  <img alt="Print asset information" border="0" src="Image/Print.gif" />
                </a>
            </td>
            <td>
              <%=ucase(objRs("TypeRequest"))%>
              &nbsp;<br>
              <%=objRs("EWM")%>
              /<%=objRs("ChangeID")%>&nbsp;
            </td>
            <td>
              <%=objRs("Category")%>
              &nbsp;<br>
              <%=objRs("Brand")%>
              &nbsp;<br>
              <%=objRs("Model")%>
              <br>
            </td>
            <td>
              <%=objRs("NewAssetName")%>
              &nbsp;
            </td>
            <td>
              <%=objRs("NewSerialNr")%>
              &nbsp;
            </td>
            <td>
              <%=objRs("Status")%>
              &nbsp;<br>
              <%=objRs("InvoiceType")%>
              &nbsp;
            </td>
            <td>
              <%=objRs("UserFName") & " " & objRs("UserLName")%>
              <br>
            </td>
            <td>
              <%=objRs("CountryOfLocation")%>
              &nbsp;<br>
              <%=objRs("Opco")%>
              &nbsp;
            </td>
          </tr>
          <%  case "ACASSETLIST" %>
          <tr bordercolor="#000000" class="acassetlist-color">
            <td align="center">
              <!--<a href="javascript:document.form.ChangeID.value='<%=" Computername = \'" & objRs("Computername") & "\' and SerialNo = \'" & objRs("SerialNo") & "\' "%>';document.form.submit();"></a>-->
              <a href="javascript:document.form.ChangeID.value='<%=" InternalTag = \'" & objRs("InternalTag")  & "\' "%>';document.form.submit();document.form.ClickedRow.Value='Yes';">
                <%
         select case ucase(session("PageAction"))
          case "EDIT"
                %>
                <img alt="Click to edit" border="0" src="Image/Edit.png"><%
          case "INFO"
                %><img alt="Click to view report" border="0" src="Image/Info.png"><%
          case else
                %><img alt="ERROR IN CODE" border="0" src="Image/Error.png"><%
         end select
                %>
              </a><a href="javascript:printasset('InternalTag=<%=objRs("InternalTag")%>')">
                <img alt="Print asset information" border="0" src="Image/Print.gif" />
              </a>
            </td>
            <td>
              <%=objRs("Computername")%>
              &nbsp;
            </td>
            <td>
              <%=objRs("SerialNo")%>
              &nbsp;
            </td>
            <td>
              <%=objRs("SupervisorFirstName") & " " & objRs("SupervisorName")%>
              &nbsp;
            </td>
            <td>
              <%=objRs("ProductModel")%>
              &nbsp;
            </td>
            <td>
              <%=objRs("fv_SLDE_BUL") & " " & objRs("fv_SLDE_Opco")%>
              &nbsp;
            </td>
            <td>
              <%=objRs("cf_HP_AssgnRead")%>
              &nbsp;
            </td>
            <td>
              <%=objRs("fv_SLDE_BillingStatus")%>
              &nbsp;
            </td>
          </tr>
          <%  case else %>
          <% end select %>
          <%
  ' Jump to next record
  objRs.MoveNext
 Wend
 ' End of table, close Record Set
 objRs.Close
End If
          %>
          </form>
        </table>
      </p>
    </td>
  </tr>
  </table>
<%else%>
<tr>
  <td>
    &nbsp;
  </td>
</tr>
</table>
<%End If%>
<%acCloseDB()%>
<!--#include file="include/pageFooter.asp"-->
