<% Language = "VBScript" %>
<%
  ' AutoIMACD.asp
  ' Version 2.10 (2007-10-31, Kees Hiemstra)
  ' - AssetCenter 4.4 migration.
  ' Version 2.02 (2007-10-03, Kees Hiemstra)
  ' - Bug fix: The selection is not shown when only one record exists in the database and the computer name is equal.
  ' Version 2.01 (2007-09-15, Kees Hiemstra)
  ' - Added extra search argument to prevent looking up canceled AutoIMACD.
  ' Version 2.00 (2007-08-15, Kees Hiemstra)
  ' - Lookup computers with the same serial number and present to the engineer so he can select the
  '   existing computer in case of an replacement with another computer name.
  ' Version 1.03 (2007-07-31, Kees Hiemstra)
  ' - fEngineer should have been fEngineerID.
  ' Version 1.02 (2007-07-26, Kees Hiemstra)
  ' - Store all information to acMACD to check if the screens are neglected after the installation.
  ' Version 1.01 (2007-07-25, Kees Hiemstra)
  ' - Added fEngineer.
  ' Version 1.00 (2007-04-17, Kees Hiemstra)
  ' - Initial version.

  ' Parameters:
  ' ===========
  ' fVersion=
  ' fUsername=
  ' fComputerName=
  ' fSerialNo=
  ' fChassisType=
  ' fManufacturer=
  ' fModel=
  ' fCountryCode=
  ' fEngineerID=
  ' fAuthentication=

  Dim PageAction
  Dim AssetData(13)

  Dim strVersion
  strVersion = "2.02.0006/1.01"

  Dim objConn, objRs, strConnect, strSQL
  Dim strCategoryName, strAuthentication
  
  Select Case Request("fChassisType")
    Case 3
      strCategoryName = "Desktop computer"
    Case 4
      strCategoryName = "Desktop computer"
    Case 5
      strCategoryName = "Desktop computer"
    Case 6
      strCategoryName = "Desktop computer"
    Case 7
      strCategoryName = "Desktop computer"
    Case 8
      strCategoryName = "Laptop"
    Case 9
      strCategoryName = "Laptop"
    Case 10
      strCategoryName = "Laptop"
    Case 11
      strCategoryName = "Handheld"
    Case 14
      strCategoryName = "Laptop"
    Case Else
      strCategoryName = ""
  End Select
%>
<!--#include file="include\EnvironmentDefaults.asp"-->
<html lang="uk" xml:lang="uk" xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>IW19 - Auto IMACD v<%= strVersion %>
  </title>
  <link href="Include/Template.css" rel="stylesheet" type="text/css" />
</head>
<body>
  <!-- HEADER -->
  <br />
  <table align="center" border="0" cellpadding="0" cellspacing="0" summary="Page header"
    width="970">
    <tr height="31">
      <td width="87">
        <a href="Index.asp">
          <img alt="Main Menu" border="0" height="31" src="Image/SaraLee.gif" width="87" /></a></td>
      <!-- width=110 -->
      <td rowspan="2" width="5">
        &nbsp;</td>
      <td class="title" colspan="3" width="600">
        MDE Web Interface
      </td>
      <td rowspan="2" width="5">
        &nbsp;</td>
      <td class="rightheader">
      </td>
    </tr>
    <tr height="76">
      <td width="87">
        <a href="Index.asp">
          <img alt="Main Menu" border="0" height="76" src="Image/HPInvent.gif" width="87" /></a></td>
      <td width="50">
        &nbsp;</td>
      <td class="subtitle" style="width: 500">
        <div id="divPageTitle">
          <span style="color: Red">AutoIMACD</span></div>
      </td>
      <td align="right" style="width: 50" valign="top">
      </td>
      <td align="right" style="width: 273" valign="top">
      </td>
    </tr>
  </table>
  <!--== PAGE ==-->
  <div id="divForm">
    <%
    ' Get the current page status/action
    PageAction = Request("PageAction")

    ' Open the database connection
    Set objConn = CreateObject("ADODB.Connection") 'Define object for connection
    Set objRs = CreateObject("ADODB.Recordset")    'Define object for record set

    objConn.Open strConnect
    %>
    <form id="AutoIMACD" action="" methode="post">
      <input name="fVersion" type="hidden" value="<%=Request("fVersion") %>" />
      <input name="fUsername" type="hidden" value="<%=Request("fUsername") %>" />
      <input name="fComputerName" type="hidden" value="<%=Request("fComputerName") %>" />
      <input name="fSerialNo" type="hidden" value="<%=Request("fSerialNo") %>" />
      <input name="fChassisType" type="hidden" value="<%=Request("fChassisType") %>" />
      <input name="fManufacturer" type="hidden" value="<%=Request("fManufacturer") %>" />
      <input name="fModel" type="hidden" value="<%=Request("fModel") %>" />
      <input name="fCountryCode" type="hidden" value="<%=Request("fCountryCode") %>" />
      <input name="fEngineerID" type="hidden" value="<%=Request("fEngineerID") %>" />
      <input name="fAuthentication" type="hidden" value="<%=Request("fAuthentication") %>" />
      <table id="Table1" align="center" cellpadding="0"
        cellspacing="5" width="970">
        <tr>
          <td>
            There is already an existing entry in the asset management database with the serial
            number
            <%=Request("fSerialNo") %>. 
            Please select which model to edit or press new to create a new entry.
          </td>
        </tr>
      </table>
      <table id="maintable" align="center" border="1" bordercolor="#000000" cellpadding="0"
        cellspacing="5" width="970">
        <tr valign="top">
          <td valign="top">
            <%
            ' Try to get the country code from either the parameters or the computer name
            ' Name convention: SL<contry code><last 10 digits of the serial number>
            strCountryCode = Request("fCountryCode")
            If strCountryCode = "" Then
              iLen = Len(Request("fSerialNo"))
              If iLen > 10 Then iLen = 10
              If Left(Request("fComputerName"), 2) = "SL" And _ 
                Right(Request("fSerialNo"), iLen) = Right(Request("fComputerName"), iLen) Then
                strCountryCode = Mid(Request("fComputerName"), 3, 2)
              End If
            End If

            ' The country code can be numeric (Microsoft location number)
            strCountry = ""
            If strCountryCode <> "" Then
              If IsNumeric(strCountryCode) Then
                strSQL = "SELECT [Country], [CountryCode] FROM RadiaRIMProd.dbo.udHRDataCountry WHERE [Active] = 1 AND [CountryNumber] = " &_
                  strCountryCode
              Else      
                strSQL = "SELECT [Country], [CountryCode] FROM RadiaRIMProd.dbo.udHRDataCountry WHERE [Active] = 1 AND [CountryCode] = '" &_
                  strCountryCode & "'"
              End If
              objRs.Open strSQL, objConn
              If Not objRs.EOF Then
                strCountry = objRs("Country")
                strCountryCode = objRs("CountryCode")
              End If
              objRs.Close
            End If
            
            If Request("SaveMACD") <> "0" Then
              ' Store the information in acMACD to investigate if engineers closes the IMACD web form without saving.
              ' This is on request of Wieger Ponstein and Jozef Balaz.
              strSQL = "INSERT INTO RadiaRIMProd.dbo.acMACD (" &_
                  "[TypeRequest]," &_
                  "[Category]," &_
                  "[Brand]," &_
                  "[Model]," &_
                  "[NewSerialNr]," &_
                  "[NewAssetName]," &_
                  "[CountryOfLocation]," &_
                  "[NTLogon]," &_
                  "[EWM]," &_
                  "[InstallDate]," &_
                  "[OnSiteEng]" &_
                ") VALUES (" &_
                  "'AutoIMACD'," &_
                  "'" & strCategoryName & "'," &_
                  "'" & Request("fManufacturer") & "'," &_
                  "'" & Request("fModel") & "'," &_
                  "'" & Request("fSerialNo") & "'," &_
                  "'" & Request("fComputerName") & "'," &_
                  "'" & strCountry & "'," &_
                  "'.\" & Request("fEngineerID") & "'," &_
                  "'" & Request("fEngineerID") & " (RC: " & iRecCount & ", CT: " & Request("fChassisType") & ")'," &_
                  "'" & Year(Now()) & "-" & Right("00" & Month(Now()), 2) & "-" & Right("00" & Day(Now()), 2) & " " &_
                    Right("00" & Hour(Now()), 2) & ":" & Right("00" & Minute(Now()), 2) & ":" & Right("00" & Second(Now()), 2) & "'," &_
                  "'HPInstall_" & strCountryCode & "'" &_
                ")"
              objConn.Execute strSQL
            End If
            %>
            <!-- Prevent the creation of the MACD entry the next time -->
            <input type="hidden" name="SaveMACD" value="0" />


            <%
            If Request("SelectAutoIMACD") = "" Then
              ' Get the information from the AssetList when the selection is not done.
              strSQL = "SELECT * FROM RadiaRIMProd.dbo.webAssetListComplete WHERE [SerialNo] = '" & Request("fSerialNo") & "' AND " &_
                "[CategoryName] IN ('Desktop computer', 'Laptop', 'Netbook')"

                AssetData(0) = "NEW"
                AssetData(1) = Request("fSerialNo")
                AssetData(2) = Request("fComputerName")
                AssetData(3) = strCategoryName
                AssetData(4) = Request("fManufacturer")
                AssetData(5) = Request("fModel")
                AssetData(6) = ""
                AssetData(7) = strCountry
                AssetData(8) = ""
                AssetData(9) = ""
                AssetData(10) = ""
                AssetData(11) = ""
                AssetData(12) = ""

              ' Count the number of records in the selection and collect the data in memory.
              iRecCount = 0
              bFirstItemMatch = False
              objRs.Open strSQL, objConn
              While Not objRs.EOF
                iRecCount = iRecCount + 1

                If objRS("ACID") > 0 Then
                  AssetData(0) = AssetData(0) & "|[ACID]=" & objRS("ACID")
                ElseIf objRS("ChangeID") > 0 Then
                  AssetData(0) = AssetData(0) & "|[ChangeID]=" & objRS("ChangeID")
                Else
                  ' Added to prevent looking up canceled AutoIMACD
                  AssetData(0) = AssetData(0) & "|[ComputerName]='" & objRS("ComputerName") & "' AND [SerialNo]='" & objRS("SerialNo") & "'"
                End If

                If iRecCount = 1 And objRS("SerialNo") = Request("fSerialNo") And objRS("ComputerName") = Request("fComputerName") Then
                  bFirstItemMatch = True
                End If

                AssetData(1) = AssetData(1) & "|" & objRS("SerialNo") 
                AssetData(2) = AssetData(2) & "|" & objRS("ComputerName")
                AssetData(3) = AssetData(3) & "|" & objRS("CategoryName")
                AssetData(4) = AssetData(4) & "|" & objRS("Brand")
                AssetData(5) = AssetData(5) & "|" & objRS("ProductModel")
                AssetData(6) = AssetData(6) & "|" & objRS("AssetTag")
                AssetData(7) = AssetData(7) & "|" & objRS("LocationCountry")
                AssetData(8) = AssetData(8) & "|" & objRS("OpcoFull")
                AssetData(9) = AssetData(9) & "|" & objRS("LocationName")
                AssetData(10) = AssetData(10) & "|" & objRS("SupervisorName")
                AssetData(11) = AssetData(11) & "|" & objRS("SupervisorFirstName")
                AssetData(12) = AssetData(12) & "|" & objRS("SupervisorEMail")

                objRs.MoveNext
              Wend
              objRs.Close
            Else
              ' Prevent that a new entry is selected based on iRecCount = 0
              iRecCount = -1
            End If

            If iRecCount = 0 Or _
              (iRecCount = 1 And bFirstItemMatch) Or _
              Request("SelectAutoIMACD") <> "" Then

              If iRecCount = 0 Or Request("SelectAutoIMACD") = "NEW" Then
                ' Nothing to select or selected to create a new record, save the information webAssetList.
                strSQL = "INSERT INTO RadiaRIMProd.dbo.webAssetList (" &_
                  "[ComputerName]," &_
                  "[SerialNo]," &_
                  "[CategoryName]," &_
                  "[Brand]," &_
                  "[ProductModel]," &_
                  "[cf_HP_AssgnRead]," &_
                  "[MaintContractRef]," &_
                  "[LocationCountry]," &_
                  "[DTInstall]" &_
                  ") VALUES (" &_
                  "'" & Request("fComputerName") & "'," &_
                  "'" & Request("fSerialNo") & "'," &_
                  "'" & strCategoryName & "'," &_
                  "'" & Request("fManufacturer") &"'," &_
                  "'" & Request("fModel") & "'," &_
                  "'In use'," &_
                  "''," &_
                  "'" & strCountry & "'," &_
                  "'" & Year(Now()) & "-" & Right("00" & Month(Now()), 2) & "-" & Right("00" & Day(Now()), 2) & "'" &_
                  ")"
                objConn.Execute strSQL

                Session("ChangeID") = "ComputerName='" & Request("fComputerName") & "' AND " &_
                  "SerialNo='" & Request("fSerialNo") & "'"
              Else
                aDetail = Split(Request("SelectAutoIMACDID"), "|", -1)

                strSelectAutoIMACD = Request("SelectAutoIMACD")
                If bFirstItemMatch Then
                  aDetail = Split(AssetData(0), "|", -1)
                  iDetail = 1
                Else
                  aDetail = Split(Request("SelectAutoIMACDID"), "|", -1)
                  iDetail = CInt(Right(strSelectAutoIMACD, Len(strSelectAutoIMACD) - 7))
                End If

                ' An existing record is selected. This record need to be updated.

                strSQL = "UPDATE RadiaRIMProd.dbo.webAssetList " &_
                  "SET [ComputerName]='" & Request("fComputerName") & "'," &_
                  "[cf_HP_AssgnRead]='In use'," &_
                  "[LocationCountry]='" & strCountry & "'" &_
                  "WHERE " & aDetail(iDetail)
                objConn.Execute strSQL

                Session("ChangeID") = aDetail(iDetail)
              End If

              Set objConn = Nothing
              Set objRs = Nothing

              Session("PageAction") = "ADD"
              Session("UID") = "HPInstall_" & strCountryCode
              Session("Engineer") = Request("fEngineerID")
              Response.Redirect "acIMACD.asp?Auto=Yes"
            Else
              %>
              <!-- Save the data from the hidden table row -->
              <input name="SelectAutoIMACDID" type="hidden" value="<%=AssetData(0) %>" />

              <table summary="show all records found">
                <%

                Response.Write("<tr bgcolor=""silver""><td></td>")
                For j = 1 To iRecCount
                  Response.Write("<th>existing record (" & j & ")</th>")
                Next
                Response.Write("<th>scanned info</th></tr>")

                ' Display all rows
                aHeader = Split("hidden|Serial number|Computer name|Category|Brand|Model|Asset tag|Country|Opco|Location|Last name|First name|Email", "|", -1)
                For i = 1 To 12
                  Response.Write("<tr bgcolor=""")
                  If i Mod 2 = 0 Then
                    Response.Write("#F4FAFF")
                  Else
                    Response.Write("#EAF5FF")
                  End If
                  Response.Write("""><td>" & aHeader(i) & "</td>")
                  aDetail = Split(AssetData(i), "|", -1)
                  For j = 1 To iRecCount
                    Response.Write("<td>")
                    If aDetail(j) = aDetail(0) Then
                      Response.Write("<strong>")
                    End If
                    Response.Write(aDetail(j) & "&nbsp;")
                    If aDetail(j) = aDetail(0) Then
                      Response.Write("</strong>")
                    End If
                    Response.Write("</td>")
                  Next
                  Response.Write("<td>" & aDetail(0) & "&nbsp;</td>")
                  Response.Write("</tr>")
                Next
                
                aDetail = Split(AssetData(0), "|", -1)
                Response.Write("<tr><td>&nbsp;</td>")
                For j = 1 To iRecCount
                  Response.Write("<td align=""center""><input type=""submit"" name=""SelectAutoIMACD"" value=""Select " & j & """ /></td>")
                Next
                Response.Write("<td align=""center""><input type=""submit"" name=""SelectAutoIMACD"" value=""NEW"" /></td>")
                Response.Write("</tr>")
                %>
              </table>
              <%
            End If

            Set objConn = Nothing
            Set objRs = Nothing
            %>
          </td>
        </tr>
      </table>
    </form>
  </div>
</body>
</html>
