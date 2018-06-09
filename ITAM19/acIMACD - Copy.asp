<%Response.buffer = True%>
<%Response.Expires = 0%>
<%
  ' Check if the user is logged otherwise redirect to the login page
  If Session("UID") = "" Then
    Session("AfterLoginGoto") = Request.ServerVariables("SCRIPT_NAME")
    Response.Redirect "Logon.asp"
    Response.End
  End if
%>
<html lang="uk" xml:lang="uk" xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>IW19 - IMACD</title>

  <script language="javascript" src="Include/popCalendar.js" type="text/javascript"></script>

  <link href="Include/Template.css" rel="stylesheet" type="text/css" />

  <script language="vbscript" type="text/vbscript">
    Option Explicit
    'Read the location of the database (ODBC connection)
    Dim strCustomer
    Dim strCustomerLong

<!--#include file="Include\EnvironmentDefaults.vbs"-->

    'Constants

    Dim iSecReadOnly, iSecAbleToEdit
    Dim iFactorGuest, iFactorSLDE, iFactorHP, iFactorAdmin
    Dim iSecSLDESuperUser, iSecSLDEEngineer, iSecHPSuperUser, iSecHPEngineer, iSecHPSupervisor, iSecHPAdmin
    Dim bCanUpdateOpCo
    
    iSecReadOnly = 16843263
    iSecAbleToEdit = 4278124032
    
    iFactorGuest = 1
    iFactorSLDE  = 256
    iFactorHP    = 65536
    iFactorAdmin = 16777216

    iSecSLDESuperUser = 1023
    iSecSLDEEngineer = 4095
    iSecHPSuperUser = 262143
    iSecHPEngineer = 1048575
    iSecHPSupervisor = 2097151
    iSecHPAdmin = 16777215

    'Declare test variables
    Dim iRight
    Dim iSecGuest
    Dim iSecSLDE
    Dim iSecHP
    Dim iSecAdmin
    Dim sSelectCountry
    Dim bDisplayState
    
    'Declare objects
    Public objConn   'Connection
    Public objRs     'Record set

    'Declare variables
    Dim strConnect
    Dim strSQL
    Dim sDefaultCountry
    Dim strPageAction
    Dim iRecCount
    Dim iSelected
    Dim sSelected
    Dim iCount
    Dim strCategory
    Dim sIMACDErrorSDesc
    Dim strBundle
    Dim strFullLastName
    Dim strTypeRequest
    Dim strRadiaOn
    Dim strNTLogon
    Dim strOpCo
  	Dim strLocation
    Dim strToday
    Dim strPurchaseDate
    Dim strStatus
    Dim strCheckQueue
    Dim strDisplayChecks
    
    'Declare field variables
    Dim sAssetName
    Dim sAssetNameOriginal
    Dim sSerialNumber
    Dim sSerialNumberOriginal
    Dim sInternalTag
    Dim sAssetTag
    Dim sCountry
    Dim sCountryCode
    Dim sState
    Dim sLocation
    Dim sLocationDetail
    Dim sInstallationDate
    Dim sLastScanDate
    Dim sTimeToRetirement
    Dim sChangeReference
    Dim sActionDate
    Dim sCategory
    Dim sBrand
    Dim sModel
    Dim sCatalogueReference
    Dim sBillingTier
    Dim sBundle
    Dim sMRNewSerialNo
    Dim sMRBackupTime
    Dim sMRActionStatus
    Dim sOpCo
    Dim sOpCoName
    Dim sOpCoFull
    Dim sCostLocation
    Dim sContractReference
    Dim sPurchaseDate
    Dim sInvoiceStatus
    Dim sInvoiceStatusOriginal
    Dim sAssetStatus
    Dim sAssetStatusOriginal
    Dim sRadiaStatus
    Dim sNetworkDomain
    Dim sNetworkLogon
    Dim sLastName
    Dim sNetworkAccountOriginal
    Dim sFirstName
    Dim sPhoneNumber
    Dim sEMailAddress
    Dim sDepartment
    Dim iChangeID
    Dim bPendingUpdate
    Dim bDummyRecord
    
    Dim sIMACDErrorEWM
    Dim sIMACDErrorDescr
    
    'Declare field arrays
    Dim aCountry
    Dim aState
    Dim aLocation
    Dim aOpCo
    Dim aInvoiceStatus
    Dim aAssetStatus
    Dim aRadiaStatus
    Dim aCategory
    Dim aBrand
    Dim aModel
    Dim aCatalogueReference
    Dim aNetworkDomain
    Dim aNetworkLogon
    Dim aStatus
    Dim aMRNewSerialNo

    'Declare screen variable
    Dim sUID 'User Identification
    Dim sChangeID
    Dim sPageAction
    Dim sAction
    Dim sAutoIMACD
    Dim sLastURL
    Dim sFormState
    Dim bDataError
    Dim bUserFound
    Dim bEdit
    Dim mNetworkLogon 'Used as temporary variable
    Dim sEngineer
    Dim strTitle
    Dim strInformation
    Dim strSpecialStatus
    Dim strRefreshStatus
    Dim strNewSerialNo
    Dim strBackupTime
    Dim strRefreshResult
    Dim bManagedRefreshAction
    Dim bTestUser
		Dim strChangeID
    Dim intMRRecordOrder
    Dim intWaveItemID

		Dim objErr, strErr
    
    ' =======================================================================================================================
    ' START PROCEDURE
    ' =======================================================================================================================
    Sub StartIMACD
      intMRRecordOrder = 0
      bDisplayState = False
      strCheckQueue = ""
      strDisplayChecks = ""
      'If an error appears in the vbscript syntax, this will not run and the error will be shown
      divOpen.InnerHTML = ""
      divOpen.Style.Visibility = "hidden"
      'There are no errors compiling the source

      sUID = MACD.urlUID.Value
      sPageAction = MACD.urlPageAction.Value
      sChangeID = MACD.urlChangeID.Value
      sAutoIMACD = MACD.urlAuto.Value
      sLastURL = MACD.urlLastURL.Value
      sEngineer = MACD.urlEngineer.Value

      If Left(sUID, 10) = "HPInstall_" Then
        'These variable would normally be collected from the database, but in case of the AutoIMACD this is hardcoded
        iSecGuest = 255 'Full
        iSecSLDE  = 255 'Full
        iSecHP    = 003 'Superuser
        iSecAdmin = 000 'None
        If sEngineer = "" Then
          sChangeReference = "Installation"
        Else
          sChangeReference = "Installation by " & sEngineer & "@demb.com"
        End If
        sDefaultCountry = Right(sUID, 2)
        sSelectCountry = sDefaultCountry
      Else
        Call GetUIDFromDB
      End If

      'Calculate acces rights
      iRight = (iSecGuest * iFactorGuest) + (iSecSLDE * iFactorSLDE) + (iSecHP * iFactorHP) + _
        (iSecAdmin * iFactorAdmin)
      strPageAction = sPageAction
      divData.Style.Visibility = "visible"
      sFormState = "Data"

      Select Case strPageAction
        Case "INFO"
          divPageTitle.InnerHTML = "Display asset information"
          Call GetDataFromDB
          Call DisplayData(False)
        Case "EDIT"
          divPageTitle.InnerHTML = "Modify asset information"
          Call GetDataFromDB
          
          If strSpecialStatus <> "" Then
						MsgBox strSpecialStatus
          End If
          
          Call DisplayData(True)
        Case "ADD"
          divPageTitle.InnerHTML = "Add new asset"
          Call GetDataFromDB
          Call DisplayData(True)
        Case Else
          divPageTitle.InnerHTML = "Unknown page action"
      End Select

      divData.InnerHTML = ""
      divData.Style.Visibility = "hidden"
      divForm.Style.Visibility = "visible"
      sFormState = "Form"
    End Sub

    ' =======================================================================================================================
    ' INITIALIZE DATA
    ' =======================================================================================================================

    ' =======================
    ' ASSET DETAILS
    ' =======================
    Sub InitializeNewAsset
      sAssetName = ""
      sAssetNameOriginal = ""
      sSerialNumber = ""
      sSerialNumberOriginal = ""
      sInternalTag = ""
      sAssetTag = ""
      sCountry = ""
      sCountryCode = ""
      sState = ""
      sLocation = ""
      sLocationDetail = ""
      sInstallationDate = Today("day")
      sLastScanDate = ""
      sTimeToRetirement = ""
      sChangeReference = ""
      sActionDate = ""
      sCategory = ""
      sBrand = ""
      sModel = ""
      sCatalogueReference = ""
      sBillingTier = ""
      sBundle = ""
      sOpCo = "-"
      sOpCoName = ""
      sOpCoFull = ""
      sCostLocation = ""
      sContractReference = ""
      sPurchaseDate = Today("day")
      sInvoiceStatus = ""
      sAssetStatus = ""
      sRadiaStatus = ""
      sNetworkDomain = ""
      sNetworkLogon = ""
      sLastName = ""
      sFirstName = ""
      sPhoneNumber = ""
      sEMailAddress = ""
      sDepartment = ""
      iChangeID = 0
      bPendingUpdate = True
      bDummyRecord = False
      
      sIMACDErrorEWM = ""
      sIMACDErrorDescr = ""

      GetAllListsFromDB
    End Sub

    ' =======================================================================================================================
    ' DATABASE ACTIONS
    ' =======================================================================================================================

    ' =======================
    ' DBOpen
    ' =======================
    Sub DBOpen
      strTitle = "IW19 - IMACD"
      SetTitle("Opening database connection")

      Set objConn = CreateObject("ADODB.Connection") 'Define object for connection
      Set objRs = CreateObject("ADODB.Recordset")    'Define object for record set

      On Error Resume Next
      objConn.Open strConnect
      If err.number <> 0 Then
'msgbox err.number
				spanHelp.style.visibility = "visible"
				'Errornumber: 3716
				spanHelp.href = ""
      End If
      On Error Goto 0
      SetTitle("Connection established")
      Call StartIMACD
    End Sub
    
    ' =======================
    ' DBClose
    ' =======================
    Sub DBClose
'      On Error Resume Next
      objConn.Close
      
      Set objRs = Nothing
      Set ObjConn = Nothing
      On Error Goto 0
      'divLoadingAction.InnerHTML = "Openening database connection..."
    End Sub
    
    ' =======================
    ' WEBUSERS
    ' =======================
    Sub GetUIDFromDB
      strSQL = "SELECT * FROM RadiaRIMProd.dbo.webUsers WHERE [UID] = '" & sUID & "'"
      
      objRs.Open strSQL, objConn

      If Not objRs.EOF Then
        iSecAdmin = objRs("SecAdmin") And 127 'To avoid overflow of the iRights long integer
        iSecHP = objRs("SecHP") And 255
        iSecSLDE = objRs("SecSLDE") And 255
        iSecGuest = objRs("SecGuest") And 255
        sDefaultCountry = objRs("Country")
        sSelectCountry = objRs("CountryAccess")
        bTestUser = objRs("TestUser")
        bCanUpdateOpCo = objRs("CanUpdateOpCo")
      End If
      objRs.Close 
    End Sub

    ' =======================
    ' ASSET DETAILS
    ' =======================
    'This function will collect the data from webAssetList.
    '
    Sub GetDataFromDB
      strSQL = "SELECT * FROM RadiaRIMProd.dbo.webAssetListComplete WHERE " & sChangeID
      objRs.Open strSQL, objConn
     
      If Not objRs.EOF Then
        sAssetName = objRs("ComputerName")
        sAssetNameOriginal = sAssetName
        sSerialNumber = objRs("SerialNo")
        sSerialNumberOriginal = sSerialNumber
        sInternalTag = objRs("InternalTag")
        sAssetTag = objRs("AssetTag")
        sCountry = objRs("LocationCountry")
        sCountryCode = objRs("CountryCode")
        sLocation = objRs("LocationName")
        If (sCountry = "United States") Then
          sState = Right(sLocation, 2)
        End If
        sLocationDetail = objRs("fv_SLDE_LocDetail")
        sInstallationDate = SQLDate(objRs("DTInstall"))
        sLastScanDate = SQLDate(objRs("DTLastScan"))
        sTimeToRetirement = objRs("TimeToRetirement")
        sActionDate = "needs to be a separate field"
        sCategory = objRs("CategoryName")
        sBrand = objRs("Brand")
        sModel = objRs("ProductModel")
        sCatalogueReference = objRs("ProductBarCode")
        sBillingTier = objRs("BillingTier")
        sBundle = objRs("fv_SLDE_Bundle")
        sOpCo = objRs("fv_SLDE_BUL")
        If IsNull(sOpCo) Then
          sOpCo = "-"
        End If
        sOpCoName = objRs("fv_SLDE_OpCo")
        sOpCoFull = objRs("OpCoFull")
        sCostLocation = objRs("CostcenterTitle")
        sPurchaseDate = SQLDate(objRs("DTAcquisition"))
        If sPurchaseDate = "" Then sPurchaseDate = Today("Day")
        sInvoiceStatus = objRs("fv_SLDE_BillingStatus")
        If sInvoiceStatus = "In contract, for refresh" And strPageAction = "EDIT" Then
					strSpecialStatus = "This asset is ordered for the Managed Refresh. Check the status of the replaced asset before changing this one to In contract, to avoid that the " & strCustomerLong & " is invoiced twice!"
					intMRRecordOrder = 2
        End If
        sInvoiceStatusOriginal = sInvoiceStatus
        sAssetStatus = objRs("cf_HP_AssgnRead")
        sAssetStatusOriginal = sAssetStatus
        sRadiaStatus = objRs("ScannerDesc")
        sNetworkDomain = objRs("ADDomain")
        sNetworkLogon = objRs("ADLogon")
        sNetworkAccountOriginal = objRs("SupervisorUserLogin")
        sLastName = objRs("SupervisorName")
        sFirstName = objRs("SupervisorFirstName")
        sPhoneNumber = objRs("SupervisorPhone")
        sEMailAddress = objRs("SupervisorEMail")
        sDepartment = Trim(objRs("SupervisorTitle"))
        iChangeID = objRs("ChangeID")
        bPendingUpdate = objRs("PendingUpdate")
        bDummyRecord = objRs("ProductModel") = "Dummy_Model for RIM Import"
        
        If sAssetStatus = "In use" And sLocation <> "" Then
					If LCase(Left(sLocation, 6)) = "stock " Then
						sLocation = Mid(sLocation, 7)
					End If
        End If

        If VarType(sPurchaseDate) <= 1 Then sPurchaseDate = Today("date")
      Else
        'The record is not found
        Call InitializeNewAsset
      End If

			On Error Resume Next      
				objRs.Close
			On Error Goto 0

      'If sRadiaStatus = "No Radia" Then
			'	strSQL = "SELECT * FROM usAssetListInfo WHERE [AssetName] = '" & sAssetName & "'"
				
			'	objRs.Open strSQL, objConn
			'	If Not objRs.EOF Then
			'		strInformation = objRs("Information")
			'	End If
			'	objRs.Close
      'End If
      
      Call GetAllListsFromDB
      
      Call GetIMACDErrorInfo
      
      Call GetRefreshInfo
    End Sub

    ' =======================
    ' ALL LISTS
    ' =======================
    Sub GetAllListsFromDB
      Call GetCountryFromDB
      Call GetState
      Call GetLocationFromDB
      Call GetCategoryFromDB
      Call GetBrandFromDB
      Call GetModelFromDB
      Call GetCatalogueReferenceFromDB
      Call GetOpCoFromDB
      Call GetNetworkDomainFromDB
      Call ValidateNetworkLogonFromDB
      Call GetInvoiceStatus
      Call GetAssetStatus
      Call GetRadiaStatus
      Call GetBillingTierFromDB
      Call DisplayBillingTier
    End Sub

    ' =======================
    ' COUNTRIES
    ' =======================
    Sub GetCountryFromDB
      Dim sTmpCountry
      strSQL = "SELECT [Country] FROM RadiaRIMProd.dbo.udHRDataManagedCountries"
      If sSelectCountry <> "ALL" Then
        strSQL = strSQL & " WHERE CHARINDEX([CountryCode], '" & sSelectCountry & "') <> 0"
      End If
    
			On Error Resume Next
				objRs.Close
			On Error Goto 0  
      objRs.Open strSQL, objConn

      aCountry = ""
      iRecCount = 0
      iSelected = 0
      If Not objRs.EOF Then
        While Not objRs.EOF
          iRecCount = iRecCount + 1
          If objRs("Country") = sCountry Then
            iSelected = 1
            sSelected = " Selected"
          Else
            sSelected = ""
          End If
          aCountry = aCountry & "<option value=""" & objRs("Country") & """" & sSelected & ">" & objRs("Country") &_
            "</option>" & vbCrLf
          If iRecCount = 1 Then sTmpCountry =  objRs("Country")
          objRs.MoveNext
        Wend
        If iSelected = 0 And iRecCount > 1 Then
          aCountry = "<option value=""-"">-- select --</option>" & vbCrLf & aCountry
        End If
        If iRecCount = 1 Then
          objRs.Close
          sCountry = sTmpCountry
          Call GetCountryCodeFromDB
          Call GetLocationFromDB
          Call DisplayLocation
          SetValidation("Location")
          Call ChangeLocation

          'Country change will have impact on the OpCo
          Call GetOpCoFromDB
          Call DisplayOpCo
          SetValidation("OpCo")
          Call ChangeOpCo
        Else
          objRs.Close
        End If 'If iRecCount
      Else
        objRs.Close
      End If
    End Sub

    ' =======================
    ' LOCATIONS
    ' =======================
    Sub GetLocationFromDB
      'This query can be used when the acSite will be setup for optimalisation.
      strSQL = "SELECT S2.[ID], " &_
        "S2.[acSite], " &_
        "MAX(CASE S1.[acSite] " &_
        "WHEN '" & sLocation & "' THEN 1 " &_
        "ELSE 0 " &_
        "END) AS 'Selected' " &_
        "FROM RadiaRIMProd.dbo.acSite AS S1 " &_
        "JOIN RadiaRIMProd.dbo.acSite AS S2 " &_
        "ON ISNULL(S1.[DestinationID], S1.[ID]) = S2.[ID] " &_
        "WHERE S1.[CountryCode] = '" & sCountryCode & "' "
        If sCountry = "United States" Then
          strSQL = strSQL & "AND (S1.[acSite] LIKE '%, " & sState & "' AND S1.[OpCoID] IS NOT NULL "
          
          If IfEdit(iSecAdmin) Or sLocation = "United States" Then
						strSQL = strSQL & "Or S1.[acSite] = 'United States'"
          End If
          
          strSQL = strSQL & ")"
        End If

        strSQL = strSQL & "GROUP BY S2.[ID], S2.[acSite] " &_
        "ORDER BY S2.[acSite] "

      objRs.Open strSQL, objConn

      aLocation = ""
      iRecCount = 0
      iSelected = 0
      If Not objRs.EOF Then
        While Not objRs.EOF
          iRecCount = iRecCount + 1
          If objRs("Selected") = 1 Then
            iSelected = 1
            sSelected = " Selected"
          Else
            sSelected = ""
          End If
          aLocation = aLocation & "<option value=""" & objRs("acSite") & """" & sSelected & ">" & objRs("acSite") &_
            "</option>" & vbCrLf
          objRs.MoveNext
        Wend
        If iSelected = 0 And iRecCount > 1 Then
          aLocation = "<option value=""-"">-- select --</option>" & vbCrLf & aLocation
        End If
      End If
      objRs.Close 
    End Sub

    ' =======================
    ' COUNTRY CODE
    ' =======================
    Sub GetCountryCodeFromDB
      'This query can be used when the acSite will be setup for optimalisation.
      strSQL = "SELECT [CountryCode] " &_
        "FROM RadiaRIMProd.dbo.udHRDataCountry " &_
        "WHERE [Country] = '" & sCountry & "' " &_
        "AND [Active] = 1 "

      objRs.Open strSQL, objConn

      If Not objRs.EOF Then
        sCountryCode = objRs("CountryCode")
      End If
      objRs.Close 
    End Sub
    
    ' =======================
    ' OPCOS
    ' =======================
    Sub GetOpCoFromDB
      If sCountry <> "United States" Then
				'Non US
        strSQL = "SELECT [OpCo], [OpCoFull] FROM RadiaRIMProd.dbo.webHRDataOpCoFull WHERE [Country] = '" & sCountry & "' AND [HPManaged] = 1 "
      ElseIf IsEmpty(sState) And IsEmpty(sLocation) Then
        strSQL = "SELECT [OpCo], [OpCoFull] FROM RadiaRIMProd.dbo.webHRDataOpCoFull AS O LEFT OUTER JOIN acSite AS S ON O.[OpCoID] = " &_
          "S.[OpCoID] WHERE O.[Country] = '" & sCountry & "' AND [HPManaged] = 1 "
      ElseIf Not IsEmpty(sState) And IsEmpty(sLocation) Then
        strSQL = "SELECT [OpCo], [OpCoFull] FROM RadiaRIMProd.dbo.webHRDataOpCoFull AS O LEFT OUTER JOIN acSite AS S ON O.[OpCoID] = " &_
          "S.[OpCoID] WHERE O.[Country] = '" & sCountry & "' AND [acSite] LIKE '%, " & sState & "' AND [HPManaged] = 1 "
      Else
        strSQL = "SELECT [OpCo], [OpCoFull] FROM RadiaRIMProd.dbo.webHRDataOpCoFull AS O LEFT OUTER JOIN acSite AS S ON O.[OpCoID] = " &_
          "S.[OpCoID] WHERE O.[Country] = '" & sCountry & "' AND [acSite] = '" & sLocation & "' AND [HPManaged] = 1 "
      End If
      'SLiCE is an OpCo without a country, so an engineer should be able to select this option.
      If IfEdit(iSecHPEngineer) Then
        strSQL = strSQL + "OR [OpCo] = '0440'"
      End If
      'Only HP Country Administrations should be able to select OpCo 9000 if the OpCo is not filled.
      If IfEdit(iSecAdmin) Or sOpCo = "6732" Then
        strSQL = strSQL + "OR [OpCo] = '6732' "
      End If
      If IfEdit(iSecHPAdmin) Or sOpCo = "9000" Then
        strSQL = strSQL + "OR [OpCo] = '9000' "
      End If
      If IfEdit(iSecAdmin) Or sOpCo = "9999" Then
        strSQL = strSQL + "OR [OpCo] = '9999' "
      End If
      strSQL = strSQL & "GROUP BY [OpCo], [OpCoFull]"
      objRs.Open strSQL, objConn

      aOpCo = ""
      iRecCount = 0
      iSelected = 0
      If Not objRs.EOF Then
        While Not objRs.EOF
          iRecCount = iRecCount + 1
          If objRs("OpCo") = sOpCo Then
            iSelected = 1
            sSelected = " Selected"
          Else
            sSelected = ""
          End If
          aOpCo = aOpCo & "<option value=""" & objRs("OpCo") & """" & sSelected & ">" &_ 
            DisplayInHTML(objRs("OpCoFull")) & "</option>" & vbCrLf
          objRs.MoveNext
        Wend
        If iSelected = 0 And iRecCount > 1 Then
          aOpCo = "<option value=""-"">-- select --</option>" & vbCrLf & aOpCo
        End If
      End If
      objRs.Close 
    End Sub

    ' =======================
    ' OPCO NAME
    ' =======================
    Sub GetOpCoNameFromDB
'      If sOpCo <> "-" Then
        strSQL = "SELECT [OpCo], [OpCoName], [OpCoFull] FROM RadiaRIMProd.dbo.webHRDataOpCoFull WHERE [OpCo] = '" & sOpCo & "'"
        objRs.Open strSQL, objConn

        If Not objRs.EOF Then
          sOpCoName = objRs("OpCoName")
          sOpCoFull = objRs("OpCoFull")
        End If
        objRs.Close 
'      Else
'        sOpCoName = ""
'        sOpCoFull = ""
'      End If
    End Sub

    ' =======================
    ' CATEGORY
    ' =======================
    Sub GetCategoryFromDB
      strSQL = "SELECT DISTINCT [Category] FROM RadiaRIMProd.dbo.acCatalog WHERE [Category] <> '' AND [Active] = 1 "
      If bDummyRecord or IsComputer(sCategory) Then
        strSQL = strSQL & "AND [Category] IN ('Desktop computer', 'Laptop', 'Thin client', 'Netbook', 'Folio Laptop', 'Tablet'" 
        'Only for administrator or if the category is already selected
				If IfEdit(iSecAdmin) Or LCase(sCategory) = "virtual machine" Then
					strSQL = strSQL & ", 'Virtual machine'"
				End If
        strSQL = strSQL & ") "
      End If
      strSQL = strSQL & "ORDER BY [Category]"
      objRs.Open strSQL, objConn

      aCategory = ""
      iRecCount = 0
      iSelected = 0
      If Not objRs.EOF Then
        While Not objRs.EOF
          iRecCount = iRecCount + 1
          If objRs("Category") = sCategory Then
            iSelected = 1
            sSelected = " Selected"
          Else
            sSelected = ""
          End If
          aCategory = aCategory & "<option value=""" & objRs("Category") & """" & sSelected & ">" & objRs("Category") &_
            "</option>" & vbCrLf
          objRs.MoveNext
        Wend
        If iSelected = 0 And iRecCount > 1 Then
          aCategory = "<option value=""-"">-- select --</option>" & vbCrLf & aCategory
        End If
      End If
      objRs.Close 
    End Sub

    ' =======================
    ' BRAND
    ' =======================
    Sub GetBrandFromDB
      strSQL = "SELECT DISTINCT [Brand] FROM RadiaRIMProd.dbo.acCatalog WHERE [Active] = 1 AND [Brand] <> '' AND [Category] = '" & sCategory &_
        "' AND [Category] <> '' AND [Brand] NOT IN ('Radia') ORDER BY [Brand]"
      objRs.Open strSQL, objConn

      aBrand = ""
      iRecCount = 0
      iSelected = 0
      If Not objRs.EOF Then
        While Not objRs.EOF
          iRecCount = iRecCount + 1
          If LCase(objRs("Brand")) = LCase(sBrand) Then
            iSelected = 1
            sSelected = " Selected"
          Else
            sSelected = ""
          End If
          aBrand = aBrand & "<option value=""" & objRs("Brand") & """" & sSelected & ">" & DisplayInHTML(objRs("Brand")) &_
            "</option>" & vbCrLf
          objRs.MoveNext
        Wend
        If iSelected = 0 And iRecCount > 1 Then
          aBrand = "<option value=""-"">-- select --</option>" & vbCrLf & aBrand
        End If
      End If
      objRs.Close 
    End Sub

    ' =======================
    ' MODEL
    ' =======================
    Sub GetModelFromDB
      strSQL = "SELECT S2.[Model], MAX(CASE S1.[Model] WHEN '" & sModel & "' THEN 1 ELSE 0 END) AS 'Selected'" &_
        "FROM RadiaRIMProd.dbo.acCatalog AS S1 JOIN RadiaRIMProd.dbo.acCatalog AS S2 ON ISNULL(S1.[DestinationID], S1.[ID]) = S2.[ID] " &_
        "WHERE S2.[Model] <> '' AND S2.[Category] = '" & sCategory & "' AND S2.[Brand] = '" & sBrand & "' AND S2.[DTDeletion] IS NULL " &_
        "GROUP BY S2.[ID], S2.[Model] " &_
        "ORDER BY S2.[Model]"
      objRs.Open strSQL, objConn

      aModel = ""
      iRecCount = 0
      iSelected = 0
      If Not objRs.EOF Then
        While Not objRs.EOF
          iRecCount = iRecCount + 1
          If objRs("Selected") = 1 Then
            iSelected = 1
            sSelected = " Selected"
          Else
            sSelected = ""
          End If

          aModel = aModel & "<option value=""" & objRs("Model") & """" & sSelected & ">" & DisplayInHTML(objRs("Model")) &_
            "</option>" & vbCrLf
          objRs.MoveNext
        Wend

        If iSelected = 0 And iRecCount > 1 Then
          aModel = "<option value=""-"">-- select --</option>" & vbCrLf & aModel
        End If
      End If
      objRs.Close 
    End Sub

    ' =======================
    ' CATALOGUE REFERENCE
    ' =======================
    Sub GetCatalogueReferenceFromDB
      strSQL = "SELECT DISTINCT [CatalogRef] FROM RadiaRIMProd.dbo.acCatalog WHERE [CatalogRef] <> '' AND [Category] = '" & sCategory &_
        "' AND [Brand] = '" & sBrand & "' AND [Model] = '" & sModel & "' ORDER BY [CatalogRef]"
      objRs.Open strSQL, objConn

      aCatalogueReference = ""
      iRecCount = 0
      iSelected = 0
      If Not objRs.EOF Then
        While Not objRs.EOF
          iRecCount = iRecCount + 1
          If objRs("CatalogRef") = sCatalogueReference Then
            iSelected = 1
            sSelected = " Selected"
          Else
            sSelected = ""
          End If
          aCatalogueReference = aCatalogueReference & "<option value=""" & objRs("CatalogRef") & """" & sSelected & ">" &_
            DisplayInHTML(objRs("CatalogRef")) & "</option>" & vbCrLf
          objRs.MoveNext
        Wend

        If iSelected = 0 And iRecCount > 1 Then
          aCatalogueReference = "<option value=""-"">-- select --</option>" & vbCrLf & aCatalogueReference
        End If
      Else
        aCatalogueReference = "<option value=""-"">&lt;Unknown&gt;</option>" & vbCrLf
      End If     
      objRs.Close 
    End Sub

    ' =======================
    ' BILLING TIER
    ' =======================
    Sub GetBillingTierFromDB
      strSQL = "SELECT DISTINCT [BillingTier] FROM RadiaRIMProd.dbo.acCatalog WHERE [BillingTier] <> '' AND [Category] = '" & sCategory &_
        "' AND [Brand] = '" & sBrand & "' AND [Model] = '" & sModel & "' AND [CatalogRef] = '" & sCatalogueReference & "'"
      objRs.Open strSQL, objConn

      sBillingTier = ""
      If Not objRs.EOF Then
        sBillingTier = objRs("BillingTier")
      End If
      objRs.Close 
    End Sub

    ' =======================
    ' NETWORK DOMAIN
    ' =======================
    Sub GetNetworkDomainFromDB
      strSQL = "SELECT DISTINCT [ADDomain] FROM RadiaRIMProd.dbo.webHRDataUser WHERE [IsDeleted] = 0 ORDER BY [ADDomain]"
      objRs.Open strSQL, objConn

      aNetworkDomain = ""
      iRecCount = 0
      iSelected = 0
      If Not objRs.EOF Then
        While Not objRs.EOF
          iRecCount = iRecCount + 1
          If objRs("ADDomain") = sNetworkDomain Then
            iSelected = 1
            sSelected = " Selected"
          Else
            sSelected = ""
          End If
          aNetworkDomain = aNetworkDomain & "<option value=""" & objRs("ADDomain") & """" & sSelected & ">" &_
            DisplayInHTML(objRs("ADDomain")) & "</option>" & vbCrLf
          objRs.MoveNext
        Wend

        If iSelected = 0 And iRecCount > 1 Then
          aNetworkDomain = "<option value=""-"">-- select --</option>" & vbCrLf & aNetworkDomain
        End If
      Else
        aNetworkDomain = "<option value=""-"">&lt;Unknown&gt;</option>" & vbCrLf
      End If     
      objRs.Close
    End Sub

    ' =======================
    ' VALIDATE NETWORK LOGON
    ' =======================
    Sub ValidateNetworkLogonFromDB
			On Error Resume Next
				objRs.Close
			On Error Goto 0
			
      If IsEmpty(mNetworkLogon) Then
        mNetworkLogon = sNetworkLogon
      End If

      aNetworkLogon = ""
      If Not IsEmpty(mNetworkLogon) Then
        strSQL = "SELECT DISTINCT [ADLogon], [OpCo], [ADDomain], [FirstName], [LastName], [Phone], " &_
          "[SMTPAddress], [adLocation] FROM RadiaRIMProd.dbo.webHRDataUser " &_
          "WHERE [ADDomain] = '" & sNetworkDomain & "' AND [ADLogon] = '" & CStrSQL(mNetworkLogon) & "' ORDER BY [ADLogon]"
        objRs.Open strSQL, objConn

        If objRs.EOF Then
          fUserFound.InnerHTML = "<font color=""tomato"">User not found in database</font>"
          If mNetworkLogon <> "" Then
            sNetworkLogon = mNetworkLogon
          ElseIf mNetworkLogon <> sNetworkLogon Then
            If sNetworkLogon = "" Then
              sNetworkLogon = mNetworkLogon
            End If
            sLastName = ""
            sFirstName = ""
            sPhoneNumber = ""
            sEMailAddress = ""
          End If
          Call DisplaySupervisor
          On Error Resume Next
          If MACD.eLastName.Style.Visibility <> "hidden" Then
            MACD.eLastName.Focus
            MACD.eLastName.Select
          End If
          On Error Goto 0
          bUserFound = False
        Else
          strOpCo = ""
    		  strLocation = objRs("adLocation")
          If objRs("OpCo") <> "" Then 
            strOpCo = " OpCo " & objRs("OpCo")
          End If
    		  If strLocation <> "" Then
		      	if strOpCo <> "" Then
				      strOpCo = strOpCo & " "
			      End If
			      strOpCo = strOpCo & "(" & strLocation & ")"
		      End If
		  
          fUserFound.InnerHTML = "<font color=""green"">User found" & strOpCo & "</font>"

          sNetworkDomain = objRs("ADDomain")
          sNetworkLogon = objRs("ADLogon")
          sLastName = objRs("LastName")
          sFirstName = objRs("FirstName")
          sPhoneNumber = objRs("Phone")
          sEMailAddress = objRs("SMTPAddress")
'          sDepartment = ""
          bUserFound = True
          Call DisplaySupervisor
        End If

        objRs.Close
      End If
    End Sub

    ' =======================
    ' NETWORK LOGON
    ' =======================
    Sub GetNetworkLogonFromDB
      If mNetworkLogon = "" Then
        mNetworkLogon = sNetworkLogon
      End If
      
      strSQL = "SELECT DISTINCT [ADLogon] FROM RadiaRIMProd.dbo.webHRDataUser WHERE [ADDomain] = '" & sNetworkDomain &_
        "' AND [ADLogon] LIKE '" & mNetworkLogon & "' ORDER BY [ADLogon]"
      objRs.Open strSQL, objConn

      aNetworkLogon = ""
      iRecCount = 0
      iSelected = 0
      If Not objRs.EOF Then
        Do While Not objRs.EOF
          iRecCount = iRecCount + 1
          If iRecCount > 25 Then
            iRecCount = 0
            aNetworkLogon = "-"
            Exit Do
          End If
          If objRs("ADLogon") = sNetworkLogon Then
            iSelected = 1
            sSelected = " Selected"
          Else
            sSelected = ""
          End If
          aNetworkLogon = aNetworkLogon & "<option value=""" & objRs("ADLogon") & """" & sSelected & ">" &_
            DisplayInHTML(objRs("ADLogon")) & "</option>" & vbCrLf
          objRs.MoveNext
        Loop
        If iSelected = 0 And iRecCount > 1 Then
          aNetworkLogon = "<option value=""-"">-- select --</option>" & vbCrLf & aNetworkLogon
        End If
      End If     
      objRs.Close
    End Sub

    ' =======================
    ' IMACD ERROR INFO
    ' =======================
    Sub GetIMACDErrorInfo
      'sIMACDErrorEWM = ""
      'sIMACDErrorSDesc = ""
      'strSQL = "SELECT [EWM], [OTSComment] FROM RadiaRIMProd.dbo.acImportIMACDError WHERE [InternalTag] = '" & sInternalTag & "'"
      'objRs.Open strSQL, objConn

      'If Not objRs.EOF Then
      '  sIMACDErrorEWM = objRs("EWM")
      '  sIMACDErrorDescr = objRs("OTSComment")
      'End If

      'objRs.Close 
    End Sub

    ' =======================
    ' REFRESH INFO
    ' =======================
		Sub GetRefreshInfo
			Dim bGetSerials, intCategory
			strRefreshResult = ""
			
			bGetSerials = False
			sMRNewSerialNo = ""
			intCategory = 0
			
			If sInternalTag <> "" Then
				strSQL = "SELECT TOP 1 I.[ID] AS 'WaveItemID', 'MR//' + W.[Name] + ' #' + CAST(I.[ID] AS varchar(6)) AS 'ChangeID', C.[State] AS 'State', " & _
					"I.[SLAgreement] AS 'SLAgreement', ISNULL(I.[NewSerialNo], '') AS 'NewSerialNo', ISNULL(I.[BackupTime], -1) AS 'BackupTime', ISNULL(I.[Action], '-') AS 'Action', W.[IsDesktop], W.[IsHandheld] "
				strSQL = strSQL & "FROM RadiaRIMProd.dbo.mrWaveItem AS I JOIN RadiaRIMProd.dbo.mrWaveCountry AS C ON I.[CountryCode] = C.[CountryCode] AND I.[WaveID] = C.[WaveID] JOIN mrWave AS W ON I.[WaveID] = W.[ID] "
				strSQL = strSQL & "WHERE I.[InternalTag] = '" & sInternalTag & "' AND I.[ToReimage] = 0 AND I.[WaveID] != 29"
				strSQL = strSQL & "ORDER BY I.[WaveID] DESC"

				objRs.Open strSQL, objConn
				
				If objRs.EOF Then
					strRefreshStatus = ""
					strChangeID = ""
				Else
					intMRRecordOrder = 1 'The current record is the asset to be refreshed
					intWaveItemID = objRs("WaveItemID")
					If objRs("IsHandheld") Then intCategory = 2
					If objRs("IsDesktop") Then intCategory = 1
					If objRs("State") = "Z" Then
						strRefreshStatus = "The refresh has been cancelled"
					ElseIf InStr("CDTX", objRs("State")) > 0 Then
						Select Case LCase(objRs("SLAgreement"))
							Case "agreed"
								strRefreshStatus = "This asset should already have been refreshed"
							Case "To be replaced by stock"
								strRefreshStatus = "This asset should already have been refreshed by stock"
							Case "postpone refresh"
								strRefreshStatus = "This assets has been postponed for this refresh"
							Case Else
								strRefreshStatus = "This asset is not validated to be refreshed"
						End Select 'Agreement
					Else
						Select Case LCase(objRs("SLAgreement"))
							Case ""
								If InStr("FISV", objRs("State")) > 0 Then
									strRefreshStatus = "Negotiation about refreshing this asset is ongoing"
								Else
									strRefreshStatus = "This asset is not validated to be refreshed"
								End If
							Case "agreed", "to be replaced by stock"
								strNewSerialNo = objRs("NewSerialNo")
								If objRs("BackupTime") < 0 Then
									strBackupTime = "User data backup not applicable"
								ElseIf objRs("BackupTime") = 0 Then
									strBackupTime = "User data backup has not been done"
								Else
									strBackupTime = "User data backup costed " & objRs("BackupTime") & " minutes"
								End If
								If strNewSerialNo = "" Then
									Select Case LCase(objRs("SLAgreement"))
										Case "agreed"
											strRefreshStatus = "This asset is bound to be refreshed"
										Case "to be replaced by stock"
											strRefreshStatus = "This asset is bound to be refreshed by stock"
									End Select
									bGetSerials = True
								Else
									strRefreshStatus = "This asset has already been refreshed"
								End If
								Select Case objRs("Action")
									Case "B"
										strRefreshResult = "User kept old equipment"
									Case "R"
										strRefreshResult = "Old equipment has been returned"
									Case Else
										strRefreshResult = "?"
								End Select
							Case "postpone refresh"
								strRefreshStatus = "This assets has been postponed for this refresh"
							Case Else
								strRefreshStatus = "This asset is not validated to be refreshed"
						End Select 'Agreement
					End If
					strChangeID = objRs("ChangeID")
				End If 'EOF
				objRs.Close
				
				If bGetSerials Then
					SetCheck("Replaced with")
					strSQL = "SELECT [SerialNo], MAX([ProductModel]) AS 'Model' FROM RadiaRIMProd.dbo.webAssetListComplete WHERE [fv_SLDE_BillingStatus] = 'In contract, for refresh' "
					Select Case intCategory
						Case 1
							strSQL = strSQL & "AND [CategoryName] IN ('Desktop computer', 'Laptop', 'Thin client', 'Netbook', 'Folio Laptop', 'Tablet') "
						Case 2
							strSQL = strSQL & "AND [CategoryName] IN ('Handheld') "
					End Select
					strSQL = strSQL & "AND [CountryCode] = '" & sCountryCode & "' GROUP BY [SerialNo] HAVING COUNT([SerialNo]) = 1"
				
					objRs.Open strSQL, objConn

					aMRNewSerialNo = "<option value=""-"">-- select --</option>" & vbCrLf
					aMRNewSerialNo = aMRNewSerialNo & "<option value=""-0-"">-- not refreshed yet --</option>" & vbCrLf
					If Not objRs.EOF Then
						While Not objRs.EOF

							aMRNewSerialNo = aMRNewSerialNo & "<option value=""" & objRs("SerialNo") & """>" & DisplayInHTML(objRs("SerialNo") & " (" &_ 
								objRs("Model") & ")") & "</option>" & vbCrLf
							
							objRs.MoveNext
						Wend
					End If 'Not Eof
					
					objRs.Close
				End If 'bGetSerials
			End If 'InternalTag <> ""
		End Sub

    ' =======================
    ' SAVE DATA ALL
    ' =======================
    Sub SaveDataAll
      Call CleanupSaveData
      
			'-------------------------
      'Saving the current record
      '-------------------------
      If sMRActionStatus <> "" Then
				'There has been a selection of a new serial number, otherwise this event couldn't have happened.
				Select Case sMRActionStatus
					Case "Back"
						'Old equipment has been taken back
						sAssetStatus = "In stock"
						sInvoiceStatus = "Not in contract, to be disposed"
						
						sChangeReference = strChangeID & "//Old back"
					Case "Both"
						'Both asset are in use by the end-user, the old equipment is NOT taken back (yet)
						sAssetStatus = "In use"
						sInvoiceStatus = "In contract"

						sChangeReference = strChangeID & "//Old both"
				End Select
      End If 'If there is an action status

			'Always save the current record
      If Trim(sChangeReference) <> "#####" Then
				Call SaveDataMACD
			End If
      Call SaveDataAssetList
      
      '---------------------
      'Saving the new record
      '---------------------
      If sMRActionStatus <> "" Then
				'Save the new record only when there is a Managed Refresh action

				iChangeID = 0
				Select Case intMRRecordOrder
					Case 1
						Call SaveManagedRefreshData

						'Get data for second record
						strSQL = "SELECT * FROM RadiaRIMProd.dbo.acAssetList WHERE [SerialNo] = '" & sMRNewSerialNo & "'"
						
						objRs.Open strSQL, objConn
						
						If objRs.EOF Then
							objRs.Close

							MsgBox "The serial number " & sMRNewSerialNo & " cannot be found in the database for refresh, please perform the action manually"
							Exit Sub
						End If

						Select Case sMRActionStatus
							Case "Back"
								sChangeReference = strChangeID & "//New back"
							Case "Both"
								sChangeReference = strChangeID & "//New both"
						End Select

						sCategory = objRs("CategoryName")
						sBrand = objRs("Brand")
						sModel = objRs("ProductModel")
						sCatalogueReference = objRs("ProductBarcode")
						sBundle = objRs("fv_SLDE_Bundle")
						sSerialNumber = objRs("SerialNo")
						sAssetName = objRs("ComputerName")
						sAssetStatus = "In use"
						sInvoiceStatus = "In contract"
						sInstallationDate = Today("date")
						sSerialNumberOriginal = objRs("SerialNo")
						sAssetNameOriginal = objRs("ComputerName")
						sAssetStatusOriginal = objRs("cf_HP_AssgnRead")
						sInvoiceStatusOriginal = objRs("fv_SLDE_BillingStatus")
						sNetworkAccountOriginal = objRs("SupervisorUserLogin")
						sInternalTag = objRs("InternalTag")
						sAssetTag = objRs("AssetTag")
						sPurchaseDate = SQLDate(objRs("DTAcquisition"))
						If sPurchaseDate = "" Then sPurchaseDate = Today("date")
						sChangeID = "[InternalTag] = '" & sInternalTag & "'"

						objRs.Close
				End Select

				If Trim(sChangeReference) <> "#####" Then
					Call SaveDataMACD
				End If
				Call SaveDataAssetList
      End If 'If there is an action status
    End Sub

    ' =======================
    ' CLEANUP SAVE DATA
    ' =======================
    Sub CleanupSaveData
      sAssetTag = CleanupStr(sAssetTag)
      sLocation = CleanupStr(sLocation)
      sLocationDetail = CleanupStr(sLocationDetail)
      sCostLocation = CleanupStr(sCostLocation)
      sContractReference = CleanupStr(sContractReference)
      sNetworkLogon = Left(CleanupStr(sNetworkLogon), 50)
      strFullLastName = CleanupStr(strFullLastName)
      sFirstName = CleanupStr(sFirstName)
      sPhoneNumber = CleanupStr(sPhoneNumber)
      sEMailAddress = Left(CleanupStr(sEMailAddress), 75)
      sDepartment = CleanupStr(sDepartment)
      
      If LCase(sAssetStatus) <> "in use" Then
				sNetworkDomain = ""
				sNetworkLogon = ""
				sLastName = ""
				sFirstName = ""
				sPhoneNumber = ""
				sEMailAddress = ""
				sDepartment = ""
      End If
    End Sub

    ' =======================
    ' SAVE DATA MACD
    ' =======================
    Sub SaveDataMACD
      strTypeRequest = ""
      'IMACD is the process to update data within AssetCenter. The IMACD is exported to and processed by AssetCenter.
      'Currently this page can handle only to events: ADD and UPDATE.

      If iChangeID > 0 Then
        'Find the latest TypeRequest and duplicate this one if it is not already exported

        strSQL = "SELECT [TypeRequest] FROM RadiaRIMProd.dbo.acMACD WHERE ISNULL([TransFlag], 'N') <> 'Y' AND [ChangeID] = " & iChangeID
        
        objRs.Open strSQL, objConn
        
        If Not objRs.EOF Then
          strTypeRequest = objRs("TypeRequest")
          If Trim(strTypeRequest) = "IGNORE" Then strTypeRequest = ""
        End If
        objRs.Close
      
        'Cancel the saved change (the last change is not exported to AssetCenter).
        strSQL = "UPDATE RadiaRIMProd.dbo.acMACD SET [TypeRequest] = 'IGNORE' WHERE [ChangeID] = " & iChangeID
        
        If LCase(sUID) <> "kees-nosave" Then
					objConn.Execute strSQL
				End If
      End If
      
      iChangeID = 0 'MACD not saved

      If strTypeRequest = "" Then
        Select Case sPageAction
          Case "ADD"
						If Trim(sInternalTag) = "" Then
							strTypeRequest = "ADD"
            Else
							strTypeRequest = "UPDATE"
						End If
          Case "EDIT"
            strTypeRequest = "UPDATE"
        End Select
      End If

      strFullLastName = sLastName
      strRadiaOn = ""
      If IsComputer(sCategory) Then
        strRadiaOn = "Yes"
        If Left(sRadiaStatus, 8) = "No Radia" Then strRadiaOn = "No"
      End If
      
      strNTLogon = ""
      If (sNetworkDomain <> "") And (sNetworkDomain <> "-") And (sNetworkLogon <> "") Then
        strNTLogon = sNetworkDomain & "\" & sNetworkLogon
      End If

      If Not IsEmpty(strTypeRequest) Then
        strSQL = "INSERT INTO RadiaRIMProd.dbo.acMACD (" &_
            "[TypeRequest]," &_
            "[Category]," &_
            "[Brand]," &_
            "[Model]," &_
            "[ProductID]," &_
            "[NewSerialNr]," &_
            "[NewAssetName]," &_
            "[LocationOfAsset]," &_
            "[DetailLocation]," &_
            "[CountryOfLocation]," &_
            "[NTLogon]," &_
            "[Status]," &_
            "[EWM]," &_
            "[CostLoc]," &_
            "[InvoiceType]," &_
            "[InstallDate]," &_
            "[OpCo]," &_
            "[ActionDate]," &_
            "[OnSiteEng]," &_
            "[OldSerialNr]," &_
            "[OldAssetName]," &_
            "[OldStatus]," &_
            "[OldBillStatus]," &_
            "[OldDomainNTLogin]," &_
            "[InternalTag]," &_
            "[AssetTag]," &_
            "[PurchaseDate]," &_
            "[RadiaOn]," &_
            "[UserLName]," &_
            "[UserFName]," &_
            "[UserPhone]," &_
            "[UserEmail]," &_
            "[UserDept]" &_
          ") VALUES (" &_
            "'" & strTypeRequest & "'," &_
            "'" & sCategory & "'," &_
            "'" & sBrand & "'," &_
            "'" & sModel & "'," &_
            "'" & sCatalogueReference & "'," &_
            "'" & sSerialNumber & "'," &_
            "'" & sAssetName & "'," &_
            "'" & CStrSQL(sLocation) & "'," &_
            "'" & CStrSQL(sLocationDetail) & "'," &_
            "'" & sCountry & "'," &_
            "'" & CStrSQL(strNTLogon) & "'," &_
            "'" & sAssetStatus & "'," &_
            "'" & CStrSQL(sChangeReference) & "'," &_
            "'" & CStrSQL(sCostLocation) & "'," &_
            "'" & sInvoiceStatus & "'," &_
            "'" & sInstallationDate & "'," &_
            "'" & sOpCoFull & "'," &_
            "'" & Today("time") & "'," &_
            "'" & MACD.urlUID.Value & "'," &_
            "'" & sSerialNumberOriginal & "'," &_
            "'" & sAssetNameOriginal & "'," &_
            "'" & sAssetStatusOriginal & "'," &_
            "'" & sInvoiceStatusOriginal & "'," &_
            "'" & CStrSQL(sNetworkAccountOriginal) & "'," &_
            "'" & sInternalTag & "'," &_
            "'" & CStrSQL(sAssetTag) & "'," &_
            "'" & sPurchaseDate & "'," &_
            "'" & strRadiaOn & "'," &_
            "'" & CStrSQL(strFullLastName) & "'," &_
            "'" & CStrSQL(sFirstName) & "'," &_
            "'" & CStrSQL(sPhoneNumber) & "'," &_
            "'" & CStrSQL(sEMailAddress) & "'," &_
            "'" & CStrSQL(sDepartment) & "'" &_
          ")"

        If LCase(sUID) <> "kees-nosave" Then
					On Error Resume Next
					strErr = ""
					objConn.Execute strSQL
					For Each objErr In objConn.Errors
						If strErr <> "" Then strErr = strErr & "; "
						strErr = strErr & objErr.Description
					Next
					On Error Goto 0
					if strErr <> "" Then
						tTxt.innerHTML = "Error: " & strErr
					End If
	      End If

        strSQL = "SELECT @@IDENTITY AS 'ChangeID'"
        objRs.Open strSQL, objConn
        
        If Not objRs.EOF Then
          iChangeID = objRs("ChangeID")
        End If
        objRs.Close
      End If
    End Sub

    ' =======================
    ' SAVE DATA ASSETLIST
    ' =======================
    Sub SaveDataAssetList
      strFullLastName = sLastName
      If VarType(sPurchaseDate) <= 1 Or sPurchaseDate = "" Then
        strPurchaseDate = "NULL"
      Else
        strPurchaseDate = "'" & Year(CDate(sPurchaseDate)) & "-" & Month(CDate(sPurchaseDate)) & "-" &_
          Day(CDate(sPurchaseDate)) & "'"
      End If

      strStatus = ""

      strSQL = "UPDATE RadiaRIMProd.dbo.webAssetList SET " &_
				"[ComputerName]='" & CStrSQL(sAssetName) & "'," &_
				"[SerialNo]='" & CStrSQL(sSerialNumber) & "'," &_
        "[AssetTag]='" & CStrSQL(sAssetTag) & "'," &_
        "[InternalTag]='" & sInternalTag & "'," &_
        "[LocationCountry]='" & sCountry & "'," &_
        "[LocationName]='" & CStrSQL(sLocation) & "'," &_
        "[fv_SLDE_LocDetail]='" & CStrSQL(sLocationDetail) & "', " &_
        "[DTInstall]='" & Year(CDate(sInstallationDate)) & "-" & Month(CDate(sInstallationDate)) & "-" &_
          Day(CDate(sInstallationDate)) & "'," &_
        "[DTAcquisition]=" & strPurchaseDate & ","

			If sChangeReference <> "#####" Then
        strSQL = strSQL & "[DTMutation]='" & Today("time") & "',"
			End If
        
      strSQL = strSQL & "[fv_SLDE_BUL]='" & sOpCo & "'," &_
        "[fv_SLDE_OpCo]='" & sOpCoName & "'," &_
        "[CostcenterTitle]='" & CStrSQL(sCostLocation) & "'," &_
        "[MaintContractRef]='" & CStrSQL(sContractReference) & "'," &_
        "[fv_SLDE_BillingStatus]='" & sInvoiceStatus & "'," &_
        "[cf_HP_AssgnRead]='" & sAssetStatus & "'," &_
        "[Status]='" & strStatus & "'," &_
        "[ScannerDesc]='" & sRadiaStatus & "'," &_
        "[CategoryName]='" & sCategory & "'," &_
        "[Brand]='" & sBrand & "'," &_
        "[ProductModel]='" & sModel & "'," &_
        "[ProductBarcode]='" & sCatalogueReference & "'," &_
        "[BillingTier]='" & sBillingTier & "'," &_
        "[fv_SLDE_Bundle]='" & sBundle & "'," &_
        "[SupervisorUserLogin]='" & sNetworkDomain & "\" & CStrSQL(sNetworkLogon) & "'," &_
        "[SupervisorName]='" & CStrSQL(strFullLastName) & "'," &_
        "[SupervisorFirstName]='" & CStrSQL(sFirstName) & "'," &_
        "[SupervisorPhone]='" & CStrSQL(Left(sPhoneNumber, 20)) & "'," &_
        "[SupervisorEMail]='" & CStrSQL(sEMailAddress) & "'," &_
        "[SupervisorTitle]='" & CStrSQL(sDepartment) & "'"
			If sChangeReference <> "#####" Then
        strSQL = strSQL & ", [ChangeID]=" & iChangeID
			End If
      strSQL = strSQL & " WHERE " & sChangeID
      
      If LCase(sUID) <> "kees-nosave" Then
				objConn.Execute strSQL
			End If
    End Sub
    
    ' =======================
    ' SAVE MANAGED REFRESH
    ' =======================
    Sub SaveManagedRefreshData
    Dim cAction
    
			Select Case sMRActionStatus
				Case "Back"
					cAction = "R"
				Case "Both"
					cAction = "B"
				Case Else
					cAction = "?"
			End Select
    
			strSQL = "UPDATE RadiaRIMProd.dbo.mrWaveItem SET [NewSerialNo] = '" & sMRNewSerialNo & "', [BackupTime] = " & sMRBackupTime & _
				", [Finished] = 1, [Action] = '" & cAction & "', [DTMutation] = GETDATE() WHERE [ID] = " & intWaveItemID

			If LCase(sUID) <> "kees-nosave" Then
				objConn.Execute strSQL
			End If
    End Sub

    ' =======================================================================================================================
    ' SELECTION LISTS
    ' =======================================================================================================================

    ' =======================
    ' STATE
    ' =======================
    Sub GetState
      aState = "<option value=""-"">---</option>"
      aState = aState & "<option value=""AK"">Alaska</option>"
      aState = aState & "<option value=""AL"">Alabama</option>"
      aState = aState & "<option value=""AS"">American Samoa</option>"
      aState = aState & "<option value=""AZ"">Arizona</option>"
      aState = aState & "<option value=""AR"">Arkansas</option>"
      aState = aState & "<option value=""CA"">California</option>"
      aState = aState & "<option value=""CO"">Colorado</option>"
      aState = aState & "<option value=""CT"">Connecticut</option>"
      'aState = aState & "<option value=""DE"">Delaware</option>"
      'aState = aState & "<option value=""DC"">District of Columbia</option>"
      'aState = aState & "<option value=""FM"">Federated States of Micronesia</option>"
      aState = aState & "<option value=""FL"">Florida</option>"
      aState = aState & "<option value=""GA"">Georgia</option>"
      aState = aState & "<option value=""GU"">Guam</option>"
      aState = aState & "<option value=""HI"">Hawaii</option>"
      aState = aState & "<option value=""ID"">Idaho</option>"
      aState = aState & "<option value=""IL"">Illinois</option>"
      aState = aState & "<option value=""IN"">Indiana</option>"
      aState = aState & "<option value=""IA"">Iowa</option>"
      aState = aState & "<option value=""KS"">Kansas</option>"
      aState = aState & "<option value=""KY"">Kentucky</option>"
      aState = aState & "<option value=""LA"">Louisiana</option>"
      aState = aState & "<option value=""ME"">Maine</option>"
      aState = aState & "<option value=""MH"">Marshall Islands</option>"
      aState = aState & "<option value=""MD"">Maryland</option>"
      aState = aState & "<option value=""MA"">Massachusetts</option>"
      aState = aState & "<option value=""MI"">Michigan</option>"
      aState = aState & "<option value=""MN"">Minnesota</option>"
      aState = aState & "<option value=""MS"">Mississippi</option>"
      aState = aState & "<option value=""MO"">Missouri</option>"
      aState = aState & "<option value=""MT"">Montana</option>"
      aState = aState & "<option value=""NE"">Nebraska</option>"
      aState = aState & "<option value=""NV"">Nevada</option>"
      aState = aState & "<option value=""NH"">New Hampshire</option>"
      aState = aState & "<option value=""NJ"">New Jersey</option>"
      aState = aState & "<option value=""NM"">New Mexico</option>"
      aState = aState & "<option value=""NY"">New York</option>"
      aState = aState & "<option value=""NC"">North Carolina</option>"
      aState = aState & "<option value=""ND"">North Dakota</option>"
      aState = aState & "<option value=""MP"">Northern Mariana Islands</option>"
      aState = aState & "<option value=""OH"">Ohio</option>"
      aState = aState & "<option value=""OK"">Oklahoma</option>"
      aState = aState & "<option value=""OR"">Oregon</option>"
      aState = aState & "<option value=""PW"">Palau</option>"
      aState = aState & "<option value=""PA"">Pennsylvania</option>"
      aState = aState & "<option value=""PR"">Puerto Rico</option>"
      aState = aState & "<option value=""RI"">Rhode Island</option>"
      aState = aState & "<option value=""SC"">South Carolina</option>"
      aState = aState & "<option value=""SD"">South Dakota</option>"
      aState = aState & "<option value=""TN"">Tennessee</option>"
      aState = aState & "<option value=""TX"">Texas</option>"
      aState = aState & "<option value=""UT"">Utah</option>"
      aState = aState & "<option value=""VT"">Vermont</option>"
      aState = aState & "<option value=""VI"">Virgin Islands</option>"
      aState = aState & "<option value=""VA"">Virginia</option>"
      aState = aState & "<option value=""WA"">Washington</option>"
      aState = aState & "<option value=""WV"">West Virginia</option>"
      aState = aState & "<option value=""WI"">Wisconsin</option>"
      aState = aState & "<option value=""WY"">Wyoming</option>"
      If iSecAdmin = 127 Or sLocation = "United States" Then
	      aState = aState & "<option value=""es"">-- No state --</option>"
      End If
    End Sub

    ' =======================
    ' INVOICE STATUS
    ' =======================
    Sub GetInvoiceStatus
      If Not IsRole(iSecHPSuperUser) Then
        aStatus = Array("In contract", "In contract, for refresh", "In contract, owned by Sara Lee", "Not in contract", _
          "Not in contract, used by HP", "Not in contract, to be disposed", _
          "Not in contract, Managed Print Services", "Not in contract, used for training", "Obsolete")

        aInvoiceStatus = ""
        iSelected = 0
        For iCount = 0 To UBound(aStatus)
          If sInvoiceStatus = aStatus(iCount) Then
            iSelected = 1
            sSelected = " Selected"
          Else
            sSelected = ""
          End If

          If (aStatus(iCount) <> "Not in contract, Managed Print Services") Then
            'When it is not about MPS, add the item
						If (aStatus(iCount) <> "In contract, for refresh") Then
							  'When it is not about refresh, add the item
							  If (aStatus(iCount) = "Not in contract, used for training") Then
								  'Check for training computer conditions
								  If (sInvoiceStatus = "Not in contract, used for training") Then
									  'It is already set, the engineer doesn't need to change it
									  aInvoiceStatus = aInvoiceStatus & "<option value=""" & aStatus(iCount) & """" & sSelected & ">" &_
										  aStatus(iCount) & "</option>" & vbCrLf
								  ElseIf ((IsEmpty(sTimeToRetirement) Or sTimeToRetirement <= 6) And IfEdit(iSecAdmin)) _
									  Or (Not IsEmpty(sTimeToRetirement) And sTimeToRetirement <= 2) And IsRole(iSecHPSupervisor) Then
									  'An administrator or supervisor may select this option when the asset is old enough?
									  aInvoiceStatus = aInvoiceStatus & "<option value=""" & aStatus(iCount) & """" & sSelected & ">" &_
										  aStatus(iCount) & "</option>" & vbCrLf
								  End If
                ElseIf (InStr("Not in contract|Obsolete", aStatus(iCount)) = 0) or (InStr("Not in contract|Obsolete", sInvoiceStatus) > 0) or IfEdit(iSecHPAdmin) Then
								  'When it is not about training, add the item
								  aInvoiceStatus = aInvoiceStatus & "<option value=""" & aStatus(iCount) & """" & sSelected & ">" &_
									  aStatus(iCount) & "</option>" & vbCrLf
							  End If
						Else
							If IfEdit(iSecHPAdmin) Then
								aInvoiceStatus = aInvoiceStatus & "<option value=""" & aStatus(iCount) & """" & sSelected & ">" &_
									aStatus(iCount) & "</option>" & vbCrLf
							End If
						End If
          Else
            If IsPrinter(sCategory) Then
              'Asset is a printer
              aInvoiceStatus = aInvoiceStatus & "<option value=""" & aStatus(iCount) & """" & sSelected & ">" &_
                aStatus(iCount) & "</option>" & vbCrLf
            End If
          End If
        Next 'iCount

        'Add "-- select --" when nothing is selected yet.
        If iSelected = 0 Then
          aInvoiceStatus = "<option value=""-"">-- select --</option>" & vbCrLf & aInvoiceStatus
        End If

        'Refresh the selection list
        fInvoiceStatus.InnerHTML = "<select class=""textbox280"" name=""eInvoiceStatus"" " &_
          "onchange=""ChangeInvoiceStatus()"">" & aInvoiceStatus & "</select>"
      Else
        If sInvoiceStatus = "Obsolete" And Left(sAssetStatus, 7) = "Retired" Then
          aInvoiceStatus = "<option value=""In contract"">In contract</option>" & vbCrLf
          aInvoiceStatus = aInvoiceStatus & "<option value=""Obsolete"" Selected>Obsolete</option>" & vbCrLf
        Else
          If sInvoiceStatus = "Not in contract, Managed Print Services" Then
            aInvoiceStatus = "<option value=""Not in contract, Managed Print Services"" Selected>Not in contract, Managed Print Services</option>" & vbCrLf &_
              "<option value=""In contract"">In contract</option>" & vbCrLf
          Else
            If sInvoiceStatus = "In contract, owned by Sara Lee" Then
              aInvoiceStatus = "<option value=""In contract, owned by Sara Lee"" Selected>In contract, owned by Sara Lee</option>" & vbCrLf
            Else
              aInvoiceStatus = "<option value=""In contract"" Selected>In contract</option>" & vbCrLf
            End If
          End If
        End If
      End If
    End Sub

    ' =======================
    ' ASSET STATUS
    ' =======================
    Sub GetAssetStatus
      aStatus = Array("In use", "In stock", "Retired (or consumed)")
      
			If sAssetStatus = "Awaiting receipt" Then sAssetStatus = ""

      aAssetStatus = ""
      iSelected = 0
      For iCount = 0 To UBound(aStatus)
        If (aStatus(iCount) <> "Retired (or consumed)" or sAssetStatus = "Retired (or consumed)" or IfEdit(iSecHPAdmin)) Then
          If sAssetStatus = aStatus(iCount) Then
            iSelected = 1
            sSelected = " Selected"
          Else
            sSelected = ""
          End If

          aAssetStatus = aAssetStatus & "<option value=""" & aStatus(iCount) & """" & sSelected & ">" & aStatus(iCount) & "</option>" & vbCrLf
        End If
      Next 'iCount

      If sAssetStatus = "Missing" Then
				iSelected = 1
        aAssetStatus = "<option value=""Missing"" Selected>Missing</option>" & aAssetStatus & vbCrLf
      End If

      'Add "-- select --" when nothing is selected yet
      If iSelected = 0 Then
        aAssetStatus = "<option value=""-"">-- select --</option>" & vbCrLf & aAssetStatus
      End If

      'Refresh the selection list
      fAssetStatus.InnerHTML = "<select class=""textbox280"" name=""eAssetStatus"" " &_
        "onchange=""ChangeAssetStatus()"">" & aAssetStatus & "</select>"
    End Sub

    ' =======================
    ' RADIA STATUS
    ' =======================
    Sub GetRadiaStatus
      aStatus = Array("No Radia")

      aRadiaStatus = ""
      iSelected = 0
      If IsComputer(sCategory) Then
        For iCount = 0 To UBound(aStatus)
          If sRadiaStatus = aStatus(iCount) Then
            iSelected = 1
            sSelected = " Selected"
          Else
            sSelected = ""
          End If

          aRadiaStatus = aRadiaStatus & "<option value=""" & aStatus(iCount) & """" & sSelected & ">" & aStatus(iCount) &_
              "</option>" & vbCrLf
        Next 'iCount
        aRadiaStatus = "<option value="""">Radia installed and functioning</option>" &_
          vbCrLf & aRadiaStatus
      Else
        If iSelected = 0 Then
          aRadiaStatus = "<option value=""N/A"">Not applicable</option>" &_
            vbCrLf & aRadiaStatus
        End If
      End If

      'Refresh the selection list
      fRadiaStatus.InnerHTML = "<select class=""textbox280"" name=""eRadiaStatus"" onchange=""ChangeRadiaStatus()"">" &_
        aRadiaStatus & "</select>"
    End Sub

    ' =======================================================================================================================
    ' DISPLAY DATA ON SCREEN
    ' =======================================================================================================================

    ' =======================
    ' FULL DISPLAY
    ' =======================
    Sub DisplayData(bEditParameter)
      bEdit = bEditParameter And Not ((iSecAdmin = 1) Or (iSecHP) = 1 Or (iSecSLDE = 1) Or _
        (iSecAdmin + iSecHP + iSecSLDE = 0))

      If Not bEdit Then
        MACD.buttonSubmit.value = "Okay"
      End If

			If sInternalTag <> "" Then
				fDetailLink.InnerHTML = "<a target=""_blank"" accesskey=""q"" href=""" & "pcdetail.asp?frmInternalTag=" & sInternalTag & """>Info</a>"
			End If

      If bEdit And (IfEdit(iSecHPEngineer)) Then
        fAssetName.InnerHTML = "<input class=""textbox280"" maxlength=""15"" name=""eAssetName"" type=""text"" value=""" &_
          sAssetName & """ onchange=""ChangeAssetName()"" />"
      Else
        fAssetName.InnerHTML = sAssetName
      End If

      If bEdit And (IfEdit(iSecHPEngineer)) Then
        fSerialNumber.InnerHTML = "<input class=""textbox280"" maxlength=""50"" name=""eSerialNumber"" type=""text"" "&_ 
          "value=""" & sSerialNumber & """ onchange=""ChangeSerialNumber()"" />"
      Else
        fSerialNumber.InnerHTML = sSerialNumber
      End If

      If bEdit And (IfEdit(iSecHPSuperUser)) Then
        fAssetTag.InnerHTML = "<input class=""textbox280"" maxlength=""41"" name=""eAssetTag"" type=""text"" value=""" &_ 
          sAssetTag & """ />"
      Else
        fAssetTag.InnerHTML = sAssetTag
      End If

      If bEdit And (IfEdit(iSecHPSuperUser)) Then
        fCountry.InnerHTML = "<select class=""textbox280"" name=""eCountry"" id=""objCountry"" "&_
          "onchange=""ChangeCountry()"">" & aCountry & "</select>"
      Else
        fCountry.InnerHTML = sCountry
      End If

      Call DisplayState

      Call DisplayLocation

      If bEdit And (IfEdit(iSecSLDESuperUser)) Then
        fLocationDetail.InnerHTML = "<input class=""textbox280"" maxlength=""101"" name=""eLocationDetail"" " &_
          "type=""text"" value=""" & sLocationDetail & """ />"
      Else
        fLocationDetail.InnerHTML = sLocationDetail
      End If

      If bEdit And (IfEdit(iSecHPSuperUser)) Then
        fInstallationDate.InnerHTML = "<input class=""textboxDate"" name=""eInstallationDate"" " &_
          "title=""Click on the calendar icon to select a date"" type=""text"" value=""" &_
          sInstallationDate & """ readonly />&nbsp;<a href=""#"" name=""jsInstallationDate"" " &_
          "onclick=""popUpCalendar(this, MACD.eInstallationDate, 'yyyy-mm-dd', '" & MACD.Style.BackgroundColor & "');"">" &_
          "<img src=""image/DateIcon.gif"" border=""0""></a>"
      Else
        fInstallationDate.InnerHTML = sInstallationDate
      End If
      
      fLastScanDate.InnerHTML = sLastScanDate

      If sTimeToRetirement = "" Then
        fTimeToRetirement.InnerHTML = ""
      ElseIf sTimeToRetirement < -1 Then
        fTimeToRetirement.InnerHTML = Abs(sTimeToRetirement) & " months overdue"
      ElseIf sTimeToRetirement = -1 Then
        fTimeToRetirement.InnerHTML = "1 month overdue"
      ElseIf sTimeToRetirement = 0 Then
        fTimeToRetirement.InnerHTML = "Now"
      ElseIf sTimeToRetirement = 1 Then
        fTimeToRetirement.InnerHTML = "1 month"
      ElseIf sTimeToRetirement > 2 Then
        fTimeToRetirement.InnerHTML = sTimeToRetirement & " months"
      Else
        fTimeToRetirement.InnerHTML = "Unknown"
      End If

      If bEdit And (IfEdit(iSecSLDESuperUser)) Then
        fChangeReference.InnerHTML = "<input class=""textbox280"" maxlength=""128"" name=""eChangeReference"" " &_
          "type=""text"" value=""" & sChangeReference & """ onchange=""ChangeChangeReference()"" />"
      Else
        fChangeReference.InnerHTML = sChangeReference
      End If

      ' =====================
      ' HARDWARE
      ' =====================
      Call DisplayCategory

      Call DisplayBrand

      Call DisplayModel

      Call DisplayCatalogueReference

      If bEdit And (IfEdit(iSecHPSuperUser)) Then
        fBundle.InnerHTML = "<input class=""checkbox"" name=""eBundle"" type=""checkbox"" name=""eBundle"" " &_
          "onchange=""ChangeBundle()"""& strBundle &" />"
        MACD.eBundle.Checked = LCase(sBundle) = "yes"
      Else
        fBundle.InnerHTML = sBundle
      End If
      
			If sInternalTag <> "" And strRefreshStatus <> "" Then
				Call DisplayManagedRefresh
			End If

      ' =====================
      ' SESSION INFO
      ' =====================
      If bDummyRecord Then
        fDummyState.InnerHTML = "Dummy record"
      End If

			If strSpecialStatus <> "" Then
				fSpecialStatus.InnerHTML = "<span style=""color:red"">" & strSpecialStatus & "</span>"
			End If

      If strInformation <> "" Then
        fInformation.InnerHTML = strInformation
      End If

			If bTestUser Then
				fTestArea.InnerHTML = "InternalTag: " & sInternalTag & "<br />"
			End If

      If iChangeID > 0 Then
        fEditState.InnerHTML = "This change is not yet exported"
      ElseIf bPendingUpdate Then
        fEditState.InnerHTML = "The last change is not yet processed"
      End If
      
      ' =====================
      ' IMACD Error info
      ' =====================
      If sIMACDErrorEWM & sIMACDErrorDescr <> "" Then
        fIMACDError.InnerHTML = "<b>IMACD error report:</b><br />" & sIMACDErrorDescr & "<br />" &_
          "<b>EWM:</b> " & sIMACDErrorEWM
      End If

      ' =====================
      ' COSTING
      ' =====================
      Call DisplayOpCo

      If bEdit And (IfEdit(iSecSLDESuperUser)) Then
        fCostLocation.InnerHTML = "<input class=""textbox280"" maxlength=""21"" name=""eCostLocation"" type=""text"" " &_
          "value=""" & sCostLocation & """ />"
      Else
        fCostLocation.InnerHTML = sCostLocation
      End If

      If bEdit And (IfEdit(iSecHPSuperUser)) Then
        fPurchaseDate.InnerHTML = "<input class=""textboxDate"" name=""ePurchaseDate"" type=""text"" value=""" &_
          sPurchaseDate & """ readonly title=""Click on the calendar icon to select a date"" />&nbsp;" &_
          "<a href=""#"" title=""Click to select a date"" name=""jsPurchaseDate"" " &_
          "onclick=""popUpCalendar(this, MACD.ePurchaseDate, 'yyyy-mm-dd', '" & MACD.Style.BackgroundColor & "')"">" &_
          "<img src=""image/DateIcon.gif"" border=""0""></a>"
      Else
        fPurchaseDate.InnerHTML = sPurchaseDate
      End If

      If bEdit And (IfEdit(iSecHPSuperUser) Or bDummyRecord) Then
        fInvoiceStatus.InnerHTML = "<select class=""textbox280"" name=""eInvoiceStatus"" " &_
          "onchange=""ChangeInvoiceStatus()"">" & aInvoiceStatus & "</select>"
      Else
        fInvoiceStatus.InnerHTML = sInvoiceStatus
      End If

      If bEdit And (IfEdit(iSecSLDEEngineer)) Then
        fAssetStatus.InnerHTML = "<select class=""textbox280"" name=""eAssetStatus"" onchange=""ChangeAssetStatus()"">" &_
          aAssetStatus & "</select>"
      Else
        fAssetStatus.InnerHTML = sAssetStatus
      End If

      If bEdit And (IfEdit(iSecSLDEEngineer)) Then
        fRadiaStatus.InnerHTML = "<select class=""textbox280"" name=""eRadiaStatus"" onchange=""ChangeRadiaStatus()"">" &_
          aRadiaStatus & "</select>"
      Else
        fRadiaStatus.InnerHTML = sRadiaStatus
      End If

      Call DisplaySupervisor

      'Validate displayed data
      Call ValidateAllFields
    End Sub

    ' =======================
    ' DISPLAY STATE
    ' =======================
    Sub DisplayState
      If bEdit And (IfEdit(iSecHPSuperUser)) And (sCountry = "United States") Then
        fState.InnerHTML = "<select class=""textbox280"" name=""eState"" id=""objState"" "&_
          "onchange=""ChangeState()"">" & aState & "</select>"
        MACD.eState.Value = sState
        bDisplayState = True
      Else
        If sCountry <> "Unites States" Then sState = ""
        fState.InnerHTML = sState
        bDisplayState = False
      End If
    End Sub
    
    ' =======================
    ' DISPLAY LOCATION
    ' =======================
    Sub DisplayLocation
      If bEdit And (IfEdit(iSecSLDEEngineer)) Then
        fLocation.InnerHTML = "<select class=""textbox280"" name=""eLocation"" onchange=""ChangeLocation()"">" & aLocation &_
          "</select>"
      Else
        fLocation.InnerHTML = sLocation
      End If
    End Sub

    ' =======================
    ' DISPLAY OPCO
    ' =======================
    Sub DisplayOpCo
      'An engineer should be able to change the OpCo if it is empty, 0000 or 9000
      If bEdit And (IfEdit(iSecHPAdmin) Or sOpCo = "-" Or sOpCo = "0000" Or sOpCo = "9000" Or IsEmpty(sOpCo) Or bCanUpdateOpCo) Then
        fOpCo.InnerHTML = "<select class=""textbox280"" name=""eOpCo"" title=""Original OpCo: " & sOpCoFull &_ 
          """ onchange=""ChangeOpCo()"">" & aOpCo & "</select>"
      Else
        fOpCo.InnerHTML = sOpCoFull
      End If
    End Sub

    ' =======================
    ' DISPLAY CATEGORY
    ' =======================
    Sub DisplayCategory
      If bEdit And (IfEdit(iSecHPSuperUser) Or bDummyRecord) Then
        fCategory.InnerHTML = "<select class=""textbox280"" name=""eCategory"" onchange=""ChangeCategory()"">" & aCategory &_
          "</select>"
      Else
        fCategory.InnerHTML = sCategory
      End If
    End Sub

    ' =======================
    ' DISPLAY BRAND
    ' =======================
    Sub DisplayBrand
      If bEdit And (IfEdit(iSecHPSuperUser) Or bDummyRecord) Then
        fBrand.InnerHTML = "<select class=""textbox280"" name=""eBrand"" onchange=""ChangeBrand()"">" & aBrand & "</select>"
      Else
        fBrand.InnerHTML = sBrand
      End If
    End Sub

    ' =======================
    ' DISPLAY MODEL
    ' =======================
    Sub DisplayModel
      If bEdit And (IfEdit(iSecHPSuperUser) Or bDummyRecord) Then
        fModel.InnerHTML = "<select class=""textbox280"" name=""eModel"" onchange=""ChangeModel()"">" & aModel & "</select>"
      Else
        fModel.InnerHTML = sModel
      End If
    End Sub

    ' =======================
    ' DISPLAY CATALOGUE REF.
    ' =======================
    Sub DisplayCatalogueReference
      If bEdit And (IfEdit(iSecHPSuperUser) Or bDummyRecord) Then
        fCatalogueReference.InnerHTML = "<select class=""textbox280"" name=""eCatalogueReference"">" & aCatalogueReference &_
          "</select>"
      Else
        fCatalogueReference.InnerHTML = sCatalogueReference
      End If
    End Sub

    ' =======================
    ' DISPLAY BILLING TIER
    ' =======================
    Sub DisplayBillingTier
      fBillingTier.InnerHTML = sBillingTier
    End Sub
    
    ' =====================
    ' MANAGED REFRESH
    ' =====================
		Sub DisplayManagedRefresh
			'Title
			rMRTitle.style.visibility = "visible"
			Select Case intMRRecordOrder
				Case 1
					lMRTitle.href = ""
				Case Else
					lMRTitle.style.visibility = "hidden"
			End Select

			'Refresh status
			rMRRefreshStatus.style.visibility = "visible"
			fMRRefreshStatus.InnerHTML = strRefreshStatus

			If strNewSerialNo <> "" Then
				rMRReplacedWith.style.visibility = "visible"
				fMRReplacedWith.InnerHTML = strNewSerialNo

				If strBackupTime <> "" Then
					rMRBackupTime.style.visibility = "visible"
					fMRBackupTime.InnerHTML = strBackupTime
				End If
				
				rMRActionStatus.style.visibility = "visible"
				If strRefreshResult <> "" Then
					fMRActionStatus.InnerHTML = strRefreshResult
				Else
					fMRActionStatus.InnerHTML = "-"
				End If
			ElseIf bEdit And (IfEdit(iSecSLDEEngineer)) And aMRNewSerialNo <> "" Then
				Call DisplayReplacedWith
			End If

			If bManagedRefreshAction Then
				rMRActionStatus.style.visibility = "visible"
			End If
		End Sub

    ' =====================
    ' REPLACED WITH
    ' =====================
		Sub DisplayReplacedWith
			If bEdit And (IfEdit(iSecSLDEEngineer)) And aMRNewSerialNo <> "" Then
				rMRReplacedWith.style.visibility = "visible"
				fMRReplacedWith.InnerHTML = "<select class=""textbox280"" name=""eReplacedWith"" id=""objReplacedWith"" " &_
					"onchange=""ChangeNewSerialNo()"">" & aMRNewSerialNo & "</select>"
			End If
		End Sub

    ' =====================
    ' BACKUP TIME
    ' =====================
		Sub DisplayBackupTime
			Dim aBackupTime
			
			If bEdit And (IfEdit(iSecSLDEEngineer)) And Left(sMRNewSerialNo, 1) <> "-" Then
				aBackupTime = "<option value=""-"">-- select --</option>" & vbCrLf
				aBackupTime = aBackupTime & "<option value=""0"">No backup performed</option>" & vbCrLf
				aBackupTime = aBackupTime & "<option value=""15"">Less than 15 minutes</option>" & vbCrLf
				aBackupTime = aBackupTime & "<option value=""30"">Between 15 minutes and 30 minutes</option>" & vbCrLf
				aBackupTime = aBackupTime & "<option value=""45"">Between 30 minutes and 45 minutes</option>" & vbCrLf
				aBackupTime = aBackupTime & "<option value=""60"">Between 45 minutes and an hour</option>" & vbCrLf
				aBackupTime = aBackupTime & "<option value=""75"">Between an hour and 75 minutes</option>" & vbCrLf
				aBackupTime = aBackupTime & "<option value=""90"">Between 75 minutes and 90 minutes</option>" & vbCrLf
				aBackupTime = aBackupTime & "<option value=""120"">Between 90 minutes and 2 hours</option>" & vbCrLf
				aBackupTime = aBackupTime & "<option value=""180"">Between 2 hours and 3 hours</option>" & vbCrLf

				rMRBackupTime.style.visibility = "visible"
				fMRBackupTime.InnerHTML = "<select class=""textbox280"" name=""eBackupTime"" id=""objBackupTime"" " &_
					"onchange=""ChangeBackupTime()"">" & aBackupTime & "</select>"

				If sMRBackupTime = "" Then
					MACD.eBackupTime.Value = "-"
				Else
					MACD.eBackupTime.Value = sMRBackupTime
				End If
			Else
				rMRBackupTime.style.visibility = "hidden"
			End If
		End Sub

    ' =====================
    ' ACTION STATUS
    ' =====================
		Sub DisplayActionStatus
			Dim aMRActionStatus
			
			If bEdit And (IfEdit(iSecSLDEEngineer)) And (sMRBackupTime <> "" And Left(sMRNewSerialNo, 1) <> "-") Then
				aMRActionStatus = "<option value=""-"">-- select --</option>" & vbCrLf
				aMRActionStatus = aMRActionStatus & "<option value=""Both"">User kept the old equipment</option>" & vbCrLf
				aMRActionStatus = aMRActionStatus & "<option value=""Back"">Old equipment has been taken back</option>" & vbCrLf
				
				rMRActionStatus.style.visibility = "visible"
				fMRActionStatus.InnerHTML = "<select class=""textbox280"" name=""eActionStatus"" id=""objActionStatus"" " &_
					"onchange=""ChangeActionStatus()"">" & aMRActionStatus & "</select>"
				
				If sMRActionStatus = "" Then
					MACD.eActionStatus.Value = "-"
				Else
					MACD.eActionStatus.Value = sMRActionStatus
				End If	
			Else
				rMRActionStatus.style.visibility = "hidden"
			End If
		End Sub

    ' =======================
    ' DISPLAY SUPERVISOR
    ' =======================
    Sub DisplaySupervisor
      Call DisplayNetworkLogon

      If IsUserMandatory And bEdit And (IfEdit(iSecSLDEEngineer)) Then
        fLastName.InnerHTML = "<input class=""textbox280"" maxlength=""101"" name=""eLastName"" type=""text"" value=""" &_
          sLastName & """ onchange=""ChangeLastName()"" />" 
      Else
        fLastName.InnerHTML = sLastName
      End If

      If IsUserMandatory And bEdit And (IfEdit(iSecSLDEEngineer)) Then
        fFirstName.InnerHTML = "<input class=""textbox280"" maxlength=""31"" name=""eFirstName"" type=""text"" value=""" &_
          sFirstName & """ onchange=""ChangeFirstName()"" />"
      Else
        fFirstName.InnerHTML = sFirstName
      End If

      If IsUserMandatory And bEdit And (IfEdit(iSecSLDEEngineer)) Then
        fPhoneNumber.InnerHTML = "<input class=""textbox280"" maxlength=""20"" name=""ePhoneNumber"" type=""text"" " &_
          "value=""" & sPhoneNumber & """ onchange=""ChangePhoneNumber()"" />"
      Else
        fPhoneNumber.InnerHTML = sPhoneNumber
      End If

      If IsUserMandatory And bEdit And (IfEdit(iSecSLDEEngineer)) Then
        fEMailAddress.InnerHTML = "<input class=""textbox280"" maxlength=""41"" name=""eEMailAddress"" type=""text"" " &_
          "value=""" & sEMailAddress & """ onchange=""ChangeEMailAddress()"" />"
      Else
        fEMailAddress.InnerHTML = sEMailAddress
      End If

      If IsUserMandatory And bEdit And (IfEdit(iSecSLDEEngineer)) Then
        fDepartment.InnerHTML = "<input class=""textbox280"" maxlength=""43"" name=""eDepartment"" type=""text"" value=""" &_
          sDepartment & """ onchange=""ChangeDepartment()"" />"
      Else
        fDepartment.InnerHTML = sDepartment
      End If
    End Sub

    ' =======================
    ' DISPLAY NETWORK LOGON
    ' =======================
    Sub DisplayNetworkLogon
      If IsUserMandatory And bEdit And (IfEdit(iSecSLDEEngineer)) Then
        fNetworkDomain.InnerHTML = "<select class=""textbox128"" name=""eNetworkDomain"" " &_
          "onchange=""ChangeNetworkDomain()"">" & aNetworkDomain & "</select>"
        If IsEmpty(sNetworkDomain) And IsUserMandatory Then
          Call SetCheck("Network domain")
          MACD.eNetworkDomain.Style.BackgroundColor = "yellow"
          bDataError = True
        Else
					MACD.eNetworkDomain.Value = sNetworkDomain
        End If
      Else
				If sNetworkLogon <> "" Then
					fNetworkDomain.InnerHTML = sNetworkDomain & "\" & sNetworkLogon
				Else
					fNetworkDomain.InnerHTML = ""
				End If
      End If

      If IsUserMandatory And bEdit And (IfEdit(iSecSLDEEngineer)) Then
        If aNetworkLogon = "-" Then
          fNetworkLogon.InnerHTML = "<input class=""textbox128"" maxlength=""43"" name=""eNetworkLogon"" type=""text"" " &_
            "value=""" & mNetworkLogon & """ onchange=""ChangeNetworkLogon()"" />"
          tTxt.InnerHTML = "There are more than 25 items found, please narrow the search."
          If IsComputer(sCategory) Then
            Call SetCheck("Network logon")
            MACD.eNetworkLogon.Style.BackgroundColor = "yellow"
          End If
        ElseIf aNetworkLogon = "" Then
          fNetworkLogon.InnerHTML = "<input class=""textbox128"" maxlength=""43"" name=""eNetworkLogon"" type=""text"" " &_ 
            "value=""" & sNetworkLogon & """ onchange=""ChangeNetworkLogon()"" />"
        Else
          fNetworkLogon.InnerHTML = "<select class=""textbox128"" name=""eNetworkLogon"" " &_
            "onchange=""ChangeNetworkLogon()"">" & aNetworkLogon & "</select>"
          If IsComputer(sCategory) Then
            Call SetCheck("Network logon")
            MACD.eNetworkLogon.Style.BackgroundColor = "yellow"
          End If
        End If
      Else
        fNetworkLogon.InnerHTML = ""
      End If
    End Sub

    ' =======================
    ' DISPLAY CHECKS
    ' =======================
    Sub DisplayChecks
      If strDisplayChecks <> "" Then
        tTxt.InnerHTML = "Please check the following field(s): " & Left(strDisplayChecks, Len(strDisplayChecks) - 2)
      Else
        tTxt.InnerHTML = ""
      End If
    End Sub

    ' =======================================================================================================================
    '
    ' CHANGES FROM SCREEN
    '
    ' =======================================================================================================================

    ' =======================
    ' ASSETNAME CHANGE
    ' =======================
    Sub ChangeAssetName
      If IfEdit(iSecHPEngineer) Then
        If (sAssetName <> Trim(MACD.eAssetName.Value)) Or DoValidation("AssetName") Then
          sAssetName = Trim(MACD.eAssetName.Value)
          'Asset name can not be empty if the asset is an active computer.
          If IsComputer(sCategory) And IsEmpty(sAssetName) And _
            Not ((sInvoiceStatus = "Obsolete" And sAssetStatus ="Retired") Or Left(sInvoiceStatus, 8) = "Awaiting") Then
            Call SetCheck("Asset name")
            MACD.eAssetName.Style.BackgroundColor = "yellow"
'            tTxt.InnerHTML = "Asset name (Computer name) can not be empty unless the asset status is 'Retired' and " &_
'              "Invoice status is 'Obsolete' or the asset status is 'Awaiting receipt' which only can be used before " &_
'              "the computer is delivered to the customer"
            bError = True
          'Computer names can not have a space in the name
          ElseIf IsComputer(sCategory) And InStr(sAssetName, " ") > 0 Then
            Call SetCheck("Asset name")
            MACD.eAssetName.Style.BackgroundColor = "yellow"
            bError = True
          Else
            Call ClearCheck("Asset name")
            MACD.eAssetName.Style.BackgroundColor = MACD.Style.BackgroundColor
          End If
        End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' SERIAL NUMBER CHANGE
    ' =======================
    Sub ChangeSerialNumber
      If IfEdit(iSecHPEngineer) Then
        If (sSerialNumber <> Trim(MACD.eSerialNumber.Value)) Or DoValidation("SerialNumber") Then
          sSerialNumber = Trim(MACD.eSerialNumber.Value)
          If IsEmpty(sSerialNumber) Then
            Call SetCheck("Serial number")
            MACD.eSerialNumber.Style.BackgroundColor = "yellow"
          Else
            Call ClearCheck("Serial number")
            MACD.eSerialNumber.Style.BackgroundColor = MACD.Style.BackgroundColor
          End If
        End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' COUNTRY CHANGE
    ' =======================
    Sub ChangeCountry
      Dim bGetData
      If IfEdit(iSecHPSuperUser) And bEdit Then
        If (sCountry <> MACD.eCountry.Value) Or DoValidation("Country") Then
          'Refresh data in the fields?
          bGetData = (sCountry <> MACD.eCountry.Value)
          sCountry = MACD.eCountry.Value

          If sCountry = "-" Or IsEmpty(sCountry) Then
            Call SetCheck("Country")
            MACD.eCountry.Style.BackgroundColor = "yellow"
            bDataError = True
          Else
            Call ClearCheck("Country")
            MACD.eCountry.Style.BackgroundColor = MACD.Style.BackgroundColor
            If bGetData Then 
              Call GetCountryCodeFromDB

              If (sCountry <> "United States") Then
                If bDisplayState Then DisplayState
              Else
                'Now the country is changed to the US, display the state and evaluate it
                SetValidation("State")
                Call ChangeState
              End If

              sLocation = ""
              Call GetLocationFromDB
              Call DisplayLocation
              SetValidation("Location")
              Call ChangeLocation

              'Country change will have impact on the OpCo
              Call GetOpCoFromDB
              Call DisplayOpCo
              SetValidation("OpCo")
              Call ChangeOpCo

            End If 'If bGetData
          End If
        End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' STATE CHANGE
    ' =======================
    Sub ChangeState
      Dim bGetData

      If bDisplayState Then
        If sCountry <> "United States" Then Call DisplayState
      Else
        If sCountry = "United States" Then Call DisplayState
      End If

      If (IfEdit(iSecHPSuperUser) Or Left(sUID, 10) = "HPInstall_") And bEdit And (sCountry = "United States") Then
        If (sState <> MACD.eState.Value) Or DoValidation("State") Then
          bGetData = (sState <> MACD.eState.Value)
          sState = MACD.eState.Value
          If sState = "-" Or IsEmpty(sState) Then
            Call SetCheck("State")
            MACD.eState.Style.BackgroundColor = "yellow"
            bDataError = True
          Else
            Call ClearCheck("State")
            MACD.eState.Style.BackgroundColor = MACD.Style.BackgroundColor
            If bGetData Then
              Call GetLocationFromDB
              Call DisplayLocation

              SetValidation("Location")
              Call ChangeLocation
              
              Call GetOpCoFromDB
              Call DisplayOpCo
              SetValidation("OpCo")
              Call ChangeOpCo
            End If 'If bGetData
          End If
        End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' LOCATION CHANGE
    ' =======================
    Sub ChangeLocation
      Dim bGetData
      If IfEdit(iSecSLDEEngineer) And bEdit Then
        If (sLocation <> MACD.eLocation.Value) Or DoValidation("Location") Then
          bGetData = (sLocation <> MACD.eLocation.Value)
          sLocation = MACD.eLocation.Value
          If sLocation = "-" Or IsEmpty(sLocation) Then
            Call SetCheck("Location")
            MACD.eLocation.Style.BackgroundColor = "yellow"
            bDataError = True
          Else
            If bGetData Then
              Call ClearCheck("Location")
              MACD.eLocation.Style.BackgroundColor = MACD.Style.BackgroundColor
              'The change is approved, now find out if the OpCo needs to change
              strSQL = "SELECT ISNULL([OpCoID], -1) AS 'OpCoID' FROM acSite WHERE [acSite] = '" & sLocation & "'"
              objRs.Open strSQL, objConn

              If Not objRs.EOF Then
                If objRs("OpCoID") <> sOpCo And objRs("OpCoID") > -1 Then
                  sOpCo = Right("0000" & objRs("OpCoID"), 4)
                  objRs.Close
                  Call GetOpCoFromDB
                  Call GetOpCoNameFromDB
                  ClearCheck("OpCo")
                  Call DisplayOpCo
                  Call ChangeOpCo
                Else
                  Call DisplayOpCo
                  SetValidation("OpCo")
									err.Clear
									on error resume next
									objRs.Close
									on error goto 0
                  Call ChangeOpCo
									on error resume next
                  objRs.Close
									on error goto 0
                End If
              Else
                objRs.Close
              End If
            End If 'If bGetData
          End If
        End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' CHANGE REFERENCE CHANGE
    ' =======================
    Sub ChangeChangeReference
      If (sChangeReference = "") Then
        sChangeReference = MACD.eChangeReference.Value
        If sChangeReference = "-" Or IsEmpty(sChangeReference) Then
          Call SetCheck("Change reference")
          MACD.eChangeReference.Style.BackgroundColor = "yellow"
          bDataError = True
        Else
          Call ClearCheck("Change reference")
          MACD.eChangeReference.Style.BackgroundColor = MACD.Style.BackgroundColor
        End If
      End If 'If changed
    End Sub

    ' =======================
    ' INSTALLATION DATE
    ' =======================
    Sub ChangeInstallationDate
      If IfEdit(iSecHPSuperUser) And bEdit Then 
        sInstallationDate = MACD.eInstallationDate.Value
        If IsEmpty(sInstallationDate) Then
          MACD.eInstallationDate.Style.BackgroundColor = "yellow"
          SetCheck("Installation date")
          bDataError = True
        Else
          Call ClearCheck("Installation date")
          MACD.eInstallationDate.Style.BackgroundColor = MACD.Style.BackgroundColor
        End If
      End If
    End Sub

    ' =======================
    ' CATEGORY CHANGE
    ' =======================
    Sub ChangeCategory
      Dim bGetData

      If IfEdit(iSecHPSuperUser) And bEdit Then
        If (sCategory <> MACD.eCategory.Value) Or DoValidation("Category") Then
          bGetData = (sCategory <> MACD.eCategory.Value)
          sCategory = MACD.eCategory.Value
          If sCategory = "-" Or IsEmpty(sCategory) Then
            Call SetCheck("Category")
            MACD.eCategory.Style.BackgroundColor = "yellow"
            bDataError = True
          Else
            If bGetData Then
              Call ClearCheck("Category")
              MACD.eCategory.Style.BackgroundColor = MACD.Style.BackgroundColor
              Call GetCategoryFromDB
              Call DisplayCategory

              Call GetBrandFromDB
              Call DisplayBrand
              Call ChangeBrand

              Call GetModelFromDB
              Call DisplayModel
              SetValidation("Model")
              Call ChangeModel

              Call GetCatalogueReferenceFromDB
              Call DisplayCatalogueReference
              Call ChangeCatalogueReference

              Call ChangeAssetName
              Call ChangeBrand
              Call GetRadiaStatus
              Call DisplaySupervisor
              SetValidation("NetworkLogon")
              Call ChangeNetworkLogon
            End If 'If bGetData
          End If
        End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' BRAND CHANGE
    ' =======================
    Sub ChangeBrand
      Dim bGetData

      If bEdit And (IfEdit(iSecHPSuperUser) Or bDummyRecord) Then
        If (sBrand <> MACD.eBrand.Value) Or DoValidation("Brand") Then
          bGetData = (sBrand <> MACD.eBrand.Value)
          sBrand = MACD.eBrand.Value
          If sBrand = "-" Or IsEmpty(sBrand) Then
            Call SetCheck("Brand")
            MACD.eBrand.Style.BackgroundColor = "yellow"
            bDataError = True
          Else
            If bGetData Then
              Call ClearCheck("Brand")
              MACD.eBrand.Style.BackgroundColor = MACD.Style.BackgroundColor
              
              Call GetBrandFromDB
              Call DisplayBrand
              
              Call GetModelFromDB
              Call DisplayModel
              SetValidation("Model")
              Call ChangeModel
              
              Call GetCatalogueReferenceFromDB
              Call DisplayCatalogueReference
              Call ChangeCatalogueReference
             End If
          End If 'If bGetData
        End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' MODEL CHANGE
    ' =======================
    Sub ChangeModel
      Dim bGetData
      
      If bEdit And (IfEdit(iSecHPSuperUser) Or bDummyRecord) Then
        If (sModel <> MACD.eModel.Value) Or DoValidation("Model") Then
          bGetData = (sModel <> MACD.eModel.Value)
          sModel = MACD.eModel.Value
          If sModel = "-" Or IsEmpty(sModel) Then
            Call SetCheck("Model")
            MACD.eModel.Style.BackgroundColor = "yellow"
            bDataError = True
          Else
            If bGetData Then
              Call ClearCheck("Model")
              MACD.eModel.Style.BackgroundColor = MACD.Style.BackgroundColor
              
              Call GetModelFromDB
              Call DisplayModel
              
              Call GetCatalogueReferenceFromDB
              Call DisplayCatalogueReference
              Call ChangeCatalogueReference
            End If
          End If 'If bGetData
        End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' CATALOGUE REFERENCE
    ' =======================
    Sub ChangeCatalogueReference
      If bEdit And (IfEdit(iSecHPSuperUser) Or bDummyRecord) Then
        If sCatalogueReference <> MACD.eCatalogueReference.Value Then
          sCatalogueReference = MACD.eCatalogueReference.Value

          Call GetBillingTierFromDB
          Call DisplayBillingTier
        End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' NEW SERIALNO CHANGE
    ' =======================
    Sub ChangeNewSerialNo
      If IfEdit(iSecSLDEEngineer) And bEdit Then
        If sMRNewSerialNo <> MACD.eReplacedWith.Value Or DoValidation("Replaced with") Then
					sMRNewSerialNo = MACD.eReplacedWith.Value
					If (IsEmpty(sMRNewSerialNo) Or sMRNewSerialNo = "-") Then
						SetCheck("Replaced with")
						MACD.eReplacedWith.Style.BackgroundColor = "yellow"
						ClearCheck("Backup time")
						ClearCheck("Action status")
						Call DisplayBackupTime
						Call DisplayActionStatus
						
						bDataError = True
					Else
						ClearCheck("Replaced with")
						MACD.eReplacedWith.Style.BackgroundColor = MACD.Style.BackgroundColor
						
						If sMRNewSerialNo <> "-0-" Then
							SetCheck("Backup time")
							Call DisplayBackupTime
							Call ChangeBackupTime
						End If
					End If 'If Check
				End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' BACKUP TIME CHANGE
    ' =======================
    Sub ChangeBackupTime
      If IfEdit(iSecSLDEEngineer) And bEdit Then
        If sMRBackupTime <> MACD.eBackupTime.Value Or DoValidation("Backup time") Then
					sMRBackupTime = MACD.eBackupTime.Value
					If (sMRBackupTime = "" Or sMRBackupTime = "-") Then
						SetCheck("Backup time")
						MACD.eBackupTime.Style.BackgroundColor = "yellow"
						bDataError = True
					Else
						ClearCheck("Backup time")
						MACD.eBackupTime.Style.BackgroundColor = MACD.Style.BackgroundColor
						
						SetCheck("Action status")
						Call DisplayActionStatus
						Call ChangeActionStatus
					End If 'If Check
				End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' ACTION STATUS CHANGE
    ' =======================
    Sub ChangeActionStatus
      If IfEdit(iSecSLDEEngineer) And bEdit Then
        If sMRActionStatus <> MACD.eActionStatus.Value Or DoValidation("Action status") Then
					sMRActionStatus = MACD.eActionStatus.Value
					If (IsEmpty(sMRActionStatus) Or sMRActionStatus = "-") Then
						SetCheck("Action status")
						MACD.eActionStatus.Style.BackgroundColor = "yellow"
						
						If Left(sChangeReference, Len(strChangeID)) = strChangeID Then
							SetCheck("Change reference")
							MACD.eChangeReference.Value = ""
							
							Call ChangeChangeReference
						End If
						bDataError = True
					Else
						ClearCheck("Action status")
						MACD.eActionStatus.Style.BackgroundColor = MACD.Style.BackgroundColor

						ClearCheck("Change reference")
						MACD.eChangeReference.Value = strChangeID
						Call ChangeChangeReference

            Call SetValidation("NetworkDomain")
            Call ChangeNetworkDomain

            Call SetValidation("NetworkLogon")
            Call ChangeNetworkLogon
					End If 'If Check
				End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' OPCO CHANGE
    ' =======================
    Sub ChangeOpCo
      Dim bGetData
      'An engineer should be able to change the OpCo if it is empty, 0000 or 9000
      If (IfEdit(iSecHPAdmin) Or sOpCo = "-" Or sOpCo = "0000" Or sOpCo = "9000" Or IsEmpty(sOpCo) Or bCanUpdateOpCo) And bEdit Then
        If (sOpCo <> MACD.eOpCo.Value) Or DoValidation("OpCo") Then
          bGetData = (sOpCo <> MACD.eOpCo.Value)
          sOpCo = MACD.eOpCo.Value

          If (sOpCo = "-" Or IsEmpty(sOpCo)) Then
            Call SetCheck("OpCo")
            MACD.eOpCo.Style.BackgroundColor = "yellow"
            bDataError = True
          Else
            If bGetData Then
              Call ClearCheck("OpCo")
              MACD.eOpCo.Style.BackgroundColor = MACD.Style.BackgroundColor
              Call GetOpCoNameFromDB
              
              If sCountry = "United States" Then
                strSQL = "SELECT S2.[acSite] FROM acSite AS S1 JOIN acSite AS S2 ON ISNULL(S1.[DestinationID], S1.[ID]) = " &_
                  "S2.[ID] WHERE S1.[OpCoID] = " & sOpCo & " GROUP BY S2.[acSite]"
                objRs.Open strSQL, objConn
                
                If Not objRs.EOF Then
                  sState = Right(objRs("acSite"), 2)
                  sLocation = objRs("acSite")
                  aLocation = "<option value=""" & sLocation & """>" & sLocation & "</option>"
                  
                  Call DisplayState
                  Call DisplayLocation
                  
                  ClearCheck("State")
                  ClearCheck("Location")
                End If
                
                objRs.Close
              End If
            End If 'If bGetData
          End If
        End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' PURCHASE DATE
    ' =======================
    Sub ChangePurchaseDate
      If IfEdit(iSecHPSuperUser) And bEdit Then 
        sPurchaseDate = MACD.ePurchaseDate.Value
        If IsEmpty(sPurchaseDate) Then
          MACD.ePurchaseDate.Style.BackgroundColor = "yellow"
          SetCheck("Purchase date")
          bDataError = True
        Else
          Call ClearCheck("Purchase date")
          MACD.ePurchaseDate.Style.BackgroundColor = MACD.Style.BackgroundColor
        End If
      End If
    End Sub

    ' =======================
    ' INVOICE STATUS CHANGE
    ' =======================
    Sub ChangeInvoiceStatus
      Dim bGetData
      If bEdit And (IfEdit(iSecHPSuperUser) Or bDummyRecord) Then
        If (sInvoiceStatus <> MACD.eInvoiceStatus.Value) Or DoValidation("InvoiceStatus") Then
          bGetData = (sInvoiceStatus <> MACD.eInvoiceStatus.Value)
          sInvoiceStatus = MACD.eInvoiceStatus.Value
          If sInvoiceStatus = "In contract, for refresh" And sChangeReference = "" Then
						MACD.eChangeReference.Value = "#####"
						Call ChangeChangeReference
          End If
          If sInvoiceStatus = "-" Or IsEmpty(sInvoiceStatus) Then
            Call SetCheck("Invoice status")
            MACD.eInvoiceStatus.Style.BackgroundColor = "yellow"
            bDataError = True
          Else
            If bGetData Then
              Call ClearCheck("Invoice status")
              MACD.eInvoiceStatus.Style.BackgroundColor = MACD.Style.BackgroundColor
              Call GetInvoiceStatus
              'In the past, the asset name was cleared when an asset was set to obsolete
              Call ChangeAssetName
              Call ChangeOpCo

              Call SetValidation("NetworkDomain")
              Call ChangeNetworkDomain

              Call SetValidation("NetworkLogon")
              Call ChangeNetworkLogon
            End If 'If bGetData
          End If
        End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' ASSET STATUS CHANGE
    ' =======================
    Sub ChangeAssetStatus
      Dim bGetData
      If IfEdit(iSecSLDEEngineer) And bEdit Then
        If (sAssetStatus <> MACD.eAssetStatus.Value) Or DoValidation("AssetStatus") Then
          bGetData = (sAssetStatus <> MACD.eAssetStatus.Value)
          sAssetStatus = MACD.eAssetStatus.Value
          If sAssetStatus = "-" Or IsEmpty(sAssetStatus) Then
            Call SetCheck("Asset status")
            MACD.eAssetStatus.Style.BackgroundColor = "yellow"
            bDataError = True
          Else
            If bGetData Then
              Call ClearCheck("Asset status")
              MACD.eAssetStatus.Style.BackgroundColor = MACD.Style.BackgroundColor
              Call GetAssetStatus
              Call ChangeAssetName
              Call ChangeOpCo

              Call SetValidation("NetworkDomain")
              Call SetValidation("NetworkLogon")
              Call DisplaySupervisor
              Call ChangeNetworkDomain
              Call ChangeNetworkLogon
            End If
          End If
        End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' RADIA STATUS CHANGE
    ' =======================
    Sub ChangeRadiaStatus
      If IfEdit(iSecSLDEEngineer) And bEdit Then
        DoValidation("RadiaStatus")
        sRadiaStatus = MACD.eRadiaStatus.Value
      End If 'If edit
    End Sub

    ' =======================
    ' DOMAIN CHANGE
    ' =======================
    Sub ChangeNetworkDomain
      If IfEdit(iSecSLDEEngineer) And bEdit And IsUserMandatory Then
        If (sNetworkDomain <> MACD.eNetworkDomain.Value) Or DoValidation("NetworkDomain") Then
          sNetworkDomain = MACD.eNetworkDomain.Value
          If IsEmpty(sNetworkDomain) And IsUserMandatory Then
            Call SetCheck("Network domain")
            MACD.eNetworkDomain.Style.BackgroundColor = "yellow"
            bDataError = True
          Else
            Call ClearCheck("Network domain")
            MACD.eNetworkDomain.Style.BackgroundColor = MACD.Style.BackgroundColor
          End If
          Call SetValidation("NetworkLogon")
          Call ChangeNetworkLogon
        End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' LOGON CHANGE
    ' =======================
    Sub ChangeNetworkLogon
      Dim bGetData

      If IfEdit(iSecSLDEEngineer) And bEdit And IsUserMandatory Then
        If (mNetworkLogon <> MACD.eNetworkLogon.Value) Or DoValidation("NetworkLogon") Then
          bGetData = (mNetworkLogon <> MACD.eNetworkLogon.Value)
          mNetworkLogon = MACD.eNetworkLogon.Value
'          Call DisplaySupervisor
          If (IsEmpty(mNetworkLogon) Or InStr(mNetworkLogon, ",") > 0 Or InStr(mNetworkLogon, "\") > 0 Or InStr(mNetworkLogon, " ") > 0 Or InStr(mNetworkLogon, "@") > 0) And IsUserMandatory Then
            Call SetCheck("Network logon")
            MACD.eNetworkLogon.Style.BackgroundColor = "yellow"

            bDataError = True
          ElseIf InStr(mNetworkLogon, "%") > 0 Then
            Call SetCheck("Network logon")
            MACD.eNetworkLogon.Style.BackgroundColor = "yellow"
            GetNetworkLogonFromDB
            Call DisplaySupervisor
          Else
            Call ClearCheck("Network logon")
            MACD.eNetworkLogon.Style.BackgroundColor = MACD.Style.BackgroundColor

            If bGetData Then             
              ValidateNetworkLogonFromDB
            End If 'If bGetData
          End If
          If IsUserMandatory Then
            Call SetValidation("LastName")
            Call ChangeLastName

            Call SetValidation("FirstName")
            Call ChangeFirstName
          End If
        End If 'If changed
      End If 'If edit
      'Focus on last name field
    End Sub

    ' =======================
    ' BUNDLE CHANGE
    ' =======================
    Sub ChangeBundle
      If IfEdit(iSecHPSuperUser) And bEdit Then
        If MACD.eBundle.Checked Then
          sBundle = "Yes"
        Else
          sBundle = "No"
        End If
      End If 'If edit
    End Sub

    ' =======================
    ' LASTNAME CHANGE
    ' =======================
    Sub ChangeLastName
      If IfEdit(iSecSLDEEngineer) And bEdit And IsUserMandatory Then
        If (sLastName <> MACD.eLastName.Value) Or DoValidation("LastName") Then
          sLastName = MACD.eLastName.Value
          If IsEmpty(sLastName) And IsUserMandatory Then
            Call SetCheck("Last name")
            MACD.eLastName.Style.BackgroundColor = "yellow"
            bDataError = True
          Else
            Call ClearCheck("Last name")
            MACD.eLastName.Style.BackgroundColor = MACD.style.backgroundColor
          End If
        End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' FIRSTNAME CHANGE
    ' =======================
    Sub ChangeFirstName
      If IfEdit(iSecSLDEEngineer) And bEdit And IsUserMandatory Then
        If (sFirstName <> MACD.eFirstName.Value) Or DoValidation("FirstName") Then
          sFirstName = MACD.eFirstName.Value
          If IsEmpty(sFirstName) And IsUserMandatory Then
            Call SetCheck("First name")
            MACD.eFirstName.Style.BackgroundColor = "yellow"
            bDataError = True
          Else
            Call ClearCheck("First name")
            MACD.eFirstName.Style.BackgroundColor = MACD.style.backgroundColor
          End If
        End If 'If changed
      End If 'If edit
    End Sub

    ' =======================
    ' PHONENUMBER CHANGE
    ' =======================
    Sub ChangePhoneNumber
			If IsUserMandatory Then
				sPhoneNumber = MACD.ePhoneNumber.Value
			End If
    End Sub

    ' =======================
    ' EMAILADDRESS CHANGE
    ' =======================
    Sub ChangeEMailAddress
			If IsUserMandatory Then
				sEMailAddress = MACD.eEMailAddress.Value
			End If
    End Sub

    ' =======================
    ' DEPARTMENT CHANGE
    ' =======================
    Sub ChangeDepartment
      If IfEdit(iSecSLDEEngineer) And bEdit And IsUserMandatory Then
        sDepartment = MACD.eDepartment.Value
      End If 'If edit
    End Sub

    ' =======================
    ' VALIDATE ALL FIELDS
    ' =======================
    Sub ValidateAllFields
      bDataError = False
      If bEdit Then
        strCheckQueue = "AssetName;SerialNumber;Country;Location;Category;Brand;Model;OpCo;InvoiceStatus;AssetStatus;" &_
          "RadiaStatus;NetworkDomain;NetworkLogon;LastName;FirstName;"
        If sCountry = "United States" Then SetValidation("State")
        Call ChangeAssetName
        Call ChangeSerialNumber
        Call ChangeCountry
        Call ChangeState
        Call ChangeLocation
        Call ChangeChangeReference
        Call ChangeCategory
        Call ChangeBrand
        Call ChangeModel
        Call ChangeCatalogueReference
        Call ChangeBundle
        If aMRNewSerialNo <> "" Then
					SetValidation("Replaced with")
					Call ChangeNewSerialNo
					If Left(sMRNewSerialNo, 1) <> "-" Then
						SetValidation("Backup time")
						Call ChangeBackupTime
						If sMRBackupTime <> "-" Then
							SetValidation("Action status")
							Call ChangeActionStatus
						End If
					End If
				End If
        Call ChangeOpCo
        Call ChangeInvoiceStatus
        Call ChangeAssetStatus
        Call ChangeRadiaStatus
        Call ChangeNetworkDomain
        Call ChangeNetworkLogon
        Call ChangeLastName
        Call ChangeFirstName
        Call ChangeDepartment
        If IfEdit(iSecHPSuperUser) And bEdit Then sAssetTag = MACD.eAssetTag.Value
        sLocationDetail = MACD.eLocationDetail.Value
        If IfEdit(iSecHPSuperUser) And bEdit Then 
          sInstallationDate = MACD.eInstallationDate.Value
          Call ChangeInstallationDate
        End If
        If IfEdit(iSecSLDESuperUser) And bEdit Then sCostLocation = MACD.eCostLocation.Value
        If IfEdit(iSecHPEngineer) Then
          sPurchaseDate = MACD.ePurchaseDate.Value
          Call ChangePurchaseDate
        End If
        Call DisplayChecks
      End If
    End Sub

    ' =======================================================================================================================
    ' BUTTONS
    ' =======================================================================================================================

    ' =======================
    ' VIEW ONLY
    ' =======================
    Sub DispDataView
      Call DisplayData(False)

      bttnSwitch.InnerHTML = "<input name=""buttonSwitch"" type=""button"" onclick=""DispDataEdit()"" value=""Edit"" />"      
    End Sub

    ' =======================
    ' EDIT DATA
    ' =======================
    Sub DispDataEdit
      Call DisplayData(True)
      
      bttnSwitch.InnerHTML = "<input name=""buttonSwitch"" type=""button"" onclick=""DispDataView()"" value=""View"" />"      
    End Sub

    ' =======================
    ' SUBMIT
    ' =======================
    Sub DoSubmit
      If bEdit Then
        Call ValidateAllFields
        If bDataError Then
					fSubmitHelp.InnerHTML = "<br />"
          Call DisplayChecks
        Else
					fSubmitHelp.InnerHTML = ""
          tTxt.InnerHTML = "Data will be saved"
          Call SaveDataAll
          rMRTitle.style.visibility = "hidden"
          rMRRefreshStatus.style.visibility = "hidden"
          rMRReplacedWith.style.visibility = "hidden"
          rMRBackupTime.style.visibility = "hidden"
          rMRActionStatus.style.visibility = "hidden"
          divForm.Style.Visibility = "hidden"
          divSave.Style.Visibility = "visible"
          sFormState = "Save"

          If sAutoIMACD = "Yes" Then
            Window.close
          Else
						If sLastURL <> "" Then
							Window.location = sLastURL
						Else
							Window.location = "http://google.com" '"http://ITAMWeb.corp.demb.com/acAssetlist.asp"
						End If
          End If
        End If
      Else
        divForm.Style.Visibility = "hidden"
        divSave.Style.Visibility = "visible"
        sFormState = "Save"
      End If
    End Sub

    ' =======================================================================================================================
    ' SUPPORT FUNCTION
    ' =======================================================================================================================

    ' =======================
    ' IS COMPUTER
    ' =======================
    Function IsComputer(pstrCategory)
      'This function is needed because an asset doesn't need an asset name when it is not a computer
      strCategory = pstrCategory
      strCategory = LCase(strCategory)
      If strCategory = "desktop computer" Then 
        IsComputer = True
      ElseIf strCategory = "laptop" Then 
        IsComputer = True
      ElseIf strCategory = "folio laptop" Then 
        IsComputer = True
      ElseIf strCategory = "table" Then 
        IsComputer = True
      ElseIf strCategory = "thin client" Then 
        IsComputer = True
      ElseIf strCategory = "netbook" Then 
        IsComputer = True
      ElseIf strCategory = "server" Then 
        IsComputer = True
      ElseIf strCategory = "virtual machine" Then 
        IsComputer = True
      Else
        IsComputer = False
      End If
    End Function
    
    ' =======================
    ' IS EMPTY
    ' =======================
    Function IsEmpty(sParameter)
      IsEmpty = IsNull(sParameter) Or LTRim(sParameter) = "" Or LTRim(sParameter) = "-"
    End Function

    ' =======================
    ' IS PRINTER
    ' =======================
    Function IsPrinter(sParameter)
      IsPrinter = (InStr(LCase(sParameter), "print") > 0) Or (LCase(sParameter) = "mfp")
    End Function
    
    ' =======================
    ' IS USER MANDATORY
    ' =======================
    Function IsUserMandatory()
      If sAssetStatus <> "In use" Then
        If sMRNewSerialNo <> "" And sMRNewSerialNo <> "-" And sMRNewSerialNo <> "-0-" Then
					IsUserMandatory = True
				Else
				  IsUserMandatory = False
				End If
      Else
        IsUserMandatory = True
      End If
    End Function
    
    ' =======================
    ' CALCULATE ACCESS
    ' =======================
    Function IfEdit(iCheckRight)
      IfEdit = (iRight And iCheckRight) = iCheckRight
    End Function
    
    ' =======================
    ' CALCULATE ROL
    ' =======================
    Function IsRole(iCheckRole)
      IsRole = (iRight Or iCheckRole) <= iCheckRole
    End Function
    
    ' =======================
    ' CHANGE HTML CHARACTERS
    ' =======================
    Function DisplayInHTML(sHTML)
      sHTML = Replace(sHTML, "&", "&amp;")
      sHTML = Replace(sHTML, "<", "&lt;")
      sHTML = Replace(sHTML, ">", "&gt;")
      DisplayInHTML = sHTML
    End Function
    
    ' =======================
    ' CHANGE HTML CHARACTERS
    ' =======================
    Function ReadFromHTML(sHTML)
      sHTML = Replace(sHTML, "&amp;", "&")
      sHTML = Replace(sHTML, "&lt;", "<")
      sHTML = Replace(sHTML, "&gt;", ">")
      ReadFromHTML = sHTML
    End Function

    ' =======================
    ' TODAY
    ' =======================
    ' Use: Today("date") will generate the current date
    '    : Today("time") will generate the current time
    Function Today(strType)
      strToday = Year(Now()) & "-" & Right("00" & Month(Now()), 2) & "-" & Right("00" & Day(Now()), 2)
      If strType = "time" Then
        strToday = strToday & " " & Right("00" & Hour(Now()), 2) & ":" & Right("00" & Minute(Now()), 2) & ":" &_
          Right("00" & Second(Now()), 2)
      End If
      Today = strToday
    End Function

    ' =======================
    ' CLEANUPSTR
    ' =======================
    Function CleanupStr(strValue)
      'Remove tabs and enter characters
      If strValue <> "" Then
        strValue = Replace(strValue, Chr(9), " ")
        strValue = Replace(strValue, Chr(10), " ")
        strValue = Replace(strValue, Chr(13), " ")
      End If
      CleanupStr = strValue
    End Function


    ' =======================
    ' SQLDate
    ' =======================
    Function SQLDate(dDate)
      If VarType(dDate) <= 1 Then
        SQLDate = ""
      Else
        SQLDate = Year(dDate) & "-" & Right("00" & Month(dDate), 2) & "-" & Right("00" & Day(dDate), 2)
      End If
    End Function

    ' =======================
    ' CStrSQL
    ' double single quoutes
    ' =======================
    Function CStrSQL(strLine)
      On Error Resume Next
      If IsNull(strLine) Then
        CStrSQL = ""
      Else
        CStrSQL = Replace(strLine, "'", "''")
      End If
      On Error Goto 0
    End Function

    ' =======================
    ' DoValidation
    ' =======================
    Function DoValidation(strName)
      If InStr(strCheckQueue, strName & ";") > 0 Then
        DoValidation = True
        strCheckQueue = Replace(strCheckQueue, strName & ";", "")
      Else
        DoValidation = False
      End If
      Call ChangeInstallationDate
      Call ChangePurchaseDate
    End Function
    
    ' =======================
    ' SetValidation
    ' =======================
    Sub SetValidation(strName)
      If InStr(strCheckQueue, strName & ";") = 0 Then
        strCheckQueue = strCheckQueue & strName & ";"
      End If
    End Sub    

    ' =======================
    ' ClearCheck
    ' =======================
    Sub ClearCheck(strName)
      Dim sCheckMessage, aCheckMessage
      If InStr(strDisplayChecks, strName & ", ") > 0 Then
        strDisplayChecks = Replace(strDisplayChecks, strName & ", ", "")
        Call DisplayChecks
      End If
    End Sub

    ' =======================
    ' SetChecks
    ' =======================
    Sub SetCheck(strName)
      If InStr(strDisplayChecks, strName & ", ") = 0 Then
        strDisplayChecks = strDisplayChecks & strName & ", "
        Call DisplayChecks
      End If
    End Sub

    ' =======================
    ' SetTitle
    ' =======================
    Sub SetTitle(strSubTitle)
      If strSubTitle = "" Then
        document.title = strTitle
      Else
        document.title = strTitle & " :: " & strSubTitle
      End If
    End Sub


  </script>

</head>
<body onunload="DBClose()">
  <!-- HEADER -->
  <table align="center" border="0" cellpadding="0" cellspacing="0" summary="Page header"
    width="970">
    <tr height="31">
      <td width="180" rowspan="2" valign="top">
        <a href="Index.asp">
          <img alt="Main Menu" border="0" height="76" src="Image/HPE.png" width="179" /></a></td>
      <!-- width=110 -->
      <td rowspan="2" width="5">
        &nbsp;</td>
      <td class="title" colspan="3" width="600">
        ITAM Web portal
      </td>
      <td rowspan="2" width="5">
        &nbsp;</td>
      <td class="rightheader">
      </td>
    </tr>
    <tr height="76">
      <td width="87"></td>
      <td width="50"> &nbsp;</td>
      <td class="subtitle" style="width: 500">
        <div id="divPageTitle">
          <span style="color: Red">Loading page...</span><br />
          <a id="spanHelp" style="font-size:small;visibility:hidden" href="">An error has occured</a>
          </div>
      </td>
      <td align="right" valign="top" style="width: 273">
        <% If Request("Auto") <> "Yes" Then %>
        <% If Session("UID") = "" Then %>
        <a href="Logon.asp">Login</a><br />
        <% else %>
        <a href="Logoff.asp">Logoff</a><br />
        <% end if %>
        <a href="Index.asp">Menu</a><br />
        <a href="javascript:history.go(-1);">Back</a>
        <% End If %>
      </td>
    </tr>
  </table>
  <!--== PAGE ==-->
  <div id="divOpen" style="visibility: visible; text-align: center">
  </div>
  <div id="divData" style="visibility: hidden; text-align: center">
    <h1 style="color: red">
      Error collecting data</h1>
  </div>
  <div id="divForm" style="visibility: hidden">
    <form id="MACD" method="post" action="">
      <input name="urlAuto" type="hidden" value="<%= Request("Auto") %>" />
      <input name="urlLastURL" type="hidden" value="<%= Request.ServerVariables("HTTP_REFERER") %>" />
      <input name="urlUID" type="hidden" value="<%= Session("UID") %>" />
      <input name="urlChangeID" type="hidden" value="<%= Session("ChangeID") %>" />
      <input name="urlPageAction" type="hidden" value="<%= Session("PageAction") %>" />
      <input name="urlEngineer" type="hidden" value="<%= Session("Engineer") %>" />
      <table id="maintable" align="center" border="1" bordercolor="#000000" cellpadding="0"
        cellspacing="5" width="970">
        <tr valign="top">
          <td valign="top">
            <!--== ADMINISTRATION AREA ==-->
            <table id="Administration" border="1" bordercolor="#ffffff" cellpadding="0" width="405">
              <tr class="menutitle" bordercolor="#000000" style="height: 23">
                <th colspan="3">
                  Administration</th>
              </tr>
              <tr style="height: 23">
                <td style="width: 105">
                  <div id="tAssetName">
                    Asset name
                  </div>
                </td>
                <td style="width: 5">
                  :</td>
                <td>
                  <div id="fAssetName" title="Only computers need to have an asset name when there not 'retired/obsolete'
                  or 'waiting receipt'">
                    acAssetList->ComputerName
                  </div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tSerialNumber">
                    Serial number</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fSerialNumber">
                    acAssetList->SerialNo</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tAssetTag">
                    Asset tag</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fAssetTag" title="Asset tag is not mandatory, when it is not provided, AssetCenter will create an asset tag.">
                    acAssetList->AssetTag</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tCountry">
                    Country</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fCountry">
                    acAssetList->LocationCountry</div>
                </td>
              </tr>
              <tr style="height:23">
                <td>
                  <div id="tState">
                    State</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fState">
                    StateList</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tLocation">
                    Location</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fLocation">
                    acAssetList->LocationName</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tLocationDetail">
                    Location detail</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fLocationDetail">
                    acAssetList->fv_SLDE_LocDetail</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tInstallationDate">
                    Installation date</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fInstallationDate">
                    acAssetList->DTInstall</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tChangeReference">
                    Change reference</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fChangeReference">
                    acAssetList->...</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tLastScanDate">
                    Last scan date</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fLastScanDate">
                    acAssetList->DTLastScan</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tTimeToRetirement">
                    Time to retirement</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fTimeToRetirement">
                    acAssetList->TimeToRetirement</div>
                </td>
              </tr>
            </table>
          </td>
          <td valign="top">
            <!--== HARDWARE AREA ==-->
            <table id="Hardware" border="1" bordercolor="#ffffff" cellpadding="0" width="405">
              <tr bordercolor="#000000" class="menutitle" style="height: 23">
                <th colspan="3">
                  Hardware</th>
              </tr>
              <tr style="height: 23">
                <td style="width: 105">
                  <div id="tCategory">
                    Category</div>
                </td>
                <td style="width: 5">
                  :</td>
                <td>
                  <div id="fCategory">
                    acAssetList->CategoryName</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tBrand">
                    Brand</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fBrand">
                    acAssetList->Brand</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tModel">
                    Model</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fModel">
                    acAssetList->Model</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tCatalogueReference">
                    Catalogue ref.</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fCatalogueReference">
                    Catalogue reference</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tBillingTier">
                    Billing tier</div>
                </td>
                <td>:</td>
                <td>
                  <div id="fBillingTier">
                    acAssetList->BillingTier*</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tBundle">
                    Bundle</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fBundle">
                    Catalogue reference</div>
                </td>
              </tr>
              <tr id="rMRTitle" bordercolor="#000000" class="menutitle" style="height: 23; visibility:hidden">
								<th colspan="3">
									Managed Refresh &nbsp; <a id="lMRTitle" target="_blank" href=""><img alt="More information for this section" src="Image/IconHelp.gif" style="border:0" /></a>
								</th>
              </tr>
              <tr id="rMRRefreshStatus" style="height: 23; visibility: hidden">
                <td>
                  <div id="tMRRefreshStatus">
                    Refresh status</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fMRRefreshStatus">
                    under development</div>
                </td>
              </tr>
              <tr id="rMRReplacedWith" style="height: 23; visibility: hidden">
                <td>
                  <div id="tMRReplacedWith">
                    Replaced with</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fMRReplacedWith">
                    under development</div>
                </td>
              </tr>
              <tr id="rMRBackupTime" style="height: 23; visibility: hidden">
                <td>
                  <div id="tMRBackupTime">
                    Backup time</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fMRBackupTime">
                    under development</div>
                </td>
              </tr>
              <tr id="rMRActionStatus" style="height: 23; visibility: hidden">
                <td>
                  <div id="tMRActionStatus">
                    Action status</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fMRActionStatus">
                    under development</div>
                </td>
              </tr>
            </table>
          </td>
          <td valign="top">
            <!--== SESSION INFO AREA ==-->
            <table border="1" bordercolor="#ffffff" cellpadding="0" width="143">
              <tr bordercolor="#000000" class="menutitle" style="height: 23">
                <th colspan="3">
                  <div id="fDetailLink">Info</div></th>
              </tr>
              <tr style="height: 23">
                <td style="width: 60; vertical-align: top">
                  Engineer</td>
                <td style="width: 3; vertical-align: top">
                  :</td>
                <td style="width: 80; vertical-align: top">
                  <%=Session("UID")%>
                </td>
              </tr>
              <tr>
								<td colspan="3">
									<div id="fSpecialStatus">
									</div>
								</td>
              </tr>
              <tr>
                <td colspan="3">
                  <div id="fDummyState">
                  </div>
                </td>
              </tr>
              <tr>
                <td colspan="3">
                  <div id="fEditState">
                  </div>
                </td>
              </tr>
              <tr>
                <td colspan="3">
                  <div id="fIMACDError">
                  </div>
                </td>
              </tr>
              <tr>
                <td colspan="3">
                  <div id="fInformation">
                  </div>
                </td>
              </tr>
              <tr>
								<td colspan="3">
									<div id="fTestArea"></div>
								</td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td valign="top">
            <!--== COSTING AREA ==-->
            <table id="Costing" border="1" bordercolor="#ffffff" cellpadding="0" width="405">
              <tr bordercolor="#000000" class="menutitle" style="height: 23">
                <th colspan="3">
                  Costing</th>
              </tr>
              <tr style="height: 23">
                <td style="width: 105">
                  <div id="tOpCo">
                    OpCo</div>
                </td>
                <td style="width: 5">
                  :</td>
                <td>
                  <div id="fOpCo">
                    OpCo</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tCostLocation">
                    Cost location</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fCostLocation">
                    CostLocation</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tPurchaseDate">
                    Acquisition date</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fPurchaseDate">
                    PurchageDate</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tInvoiceStatus">
                    Billing status</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fInvoiceStatus">
                    InvoiceStatus</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tAssetStatus">
                    Asset status</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fAssetStatus">
                    AssetStatus</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tRadiaStatus">
                    Radia status</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fRadiaStatus">
                    ScannerDesc</div>
                </td>
              </tr>
            </table>
          </td>
          <td valign="top">
            <!--== SUPERVISOR AREA ==-->
            <table id="Supervisor" border="1" bordercolor="#ffffff" cellpadding="0" width="405">
              <tr bordercolor="#000000" class="menutitle" style="height: 23">
                <th colspan="3">
                  Main user</th>
              </tr>
              <tr style="height: 23">
                <td style="width: 105">
                  <div id="tNetworkLogon">
                    Network logon</div>
                </td>
                <td style="width: 5">
                  :</td>
                <td>
                  <span id="fNetworkDomain">Network domain</span> <span id="fNetworkLogon" title="Use % as wildcard, the list is limited to 25 items.">
                    Network logon</span>
                </td>
              </tr>
<!--
              <tr style="height: 23">
                <td>
                  <div id="tEmployeeNumber">
                    Employee number</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fEmployeeNumber">
                    Employee number is not yet available</div>
                </td>
              </tr>
-->
              <tr style="height: 23">
                <td>
                  <div id="tUserFound">
                    &nbsp;</div>
                </td>
                <td>
                  &nbsp;</td>
                <td>
                  <div id="fUserFound">
                    &nbsp;</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tLastName">
                    Last name</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fLastName">
                    LastName</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tFirstName">
                    First name</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fFirstName">
                    FirstName</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tPhoneNumber">
                    Phone number</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fPhoneNumber">
                    PhoneNumber</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tEMailAddress">
                    E-mail address</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fEMailAddress">
                    EMailAddress</div>
                </td>
              </tr>
              <tr style="height: 23">
                <td>
                  <div id="tDepartment">
                    Department</div>
                </td>
                <td>
                  :</td>
                <td>
                  <div id="fDepartment">
                    Department</div>
                </td>
              </tr>
            </table>
          </td>
          <td bordercolor="#ffffff" align="center" valign="bottom">
            <!--== BUTTON AREA ==-->
            <div id="bttnSwitch">
            </div>
            <!--
            <div id="bttnClearUser" style="visibility:hidden">
              <input name="buttonClearUser" onclick="ClearUser()" type="button" value="Clear user" /></div>
            <br />
-->
            <div id="fSubmitHelp"></div>
            <br />
            <div id="bttnSubmit">
              <input name="buttonSubmit" type="button" onclick="DoSubmit()" value="Submit" accesskey="s"/></div>
          </td>
        </tr>
        <tr>
          <td colspan="2" bordercolor="#ffffff">
            <br />
            <div id="tTxt">
            </div>
          </td>
        </tr>
        <tr>
          <td bordercolor="#ffffff" colspan="2">
            <br />
            <div id="tDebug">
            </div>
          </td>
        </tr>
      </table>
    </form>
  </div>
  <div id="divSave" style="visibility: hidden">
  </div>

  <script language="vbscript" type="text/vbscript">Call DBOpen</script>
</body>
</html>
