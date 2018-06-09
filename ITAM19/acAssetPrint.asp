<% Language="VBScript" %>
<% Response.buffer = True %>
<% Response.Expires = 0 %>
<% Response.CacheControl = "no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %>
<html lang="uk" xml:lang="uk" xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>WI19 - Print asset information</title>
  <link href="Include/Print.css" rel="stylesheet" type="text/css" />
</head>
<body>
  <!--#include file="include/dbFunction.asp"-->
  <!--#include file="include/globalVars.asp"-->
  <%
  Call acOpenDB()

  If Trim(Request.QueryString("ChangeID")) <> "" Then
    strSQL = "[ChangeID] = " & Trim(Request.QueryString("ChangeID"))
  ElseIf Trim(Request.QueryString("InternalTag")) <> "" Then
    strSQL = "[InternalTag] = '" & Trim(Request.QueryString("InternalTag")) & "'"
  End If
  
  If strSQL <> "" Then
    strSQL = "SELECT *, ISNULL([ChangeID], 0) AS 'ChangeCommitted' FROM RadiaRIMProd.dbo.webAssetList WHERE " & strSQL
    objRs.Open strSQL, objConn
    
    If Not objRs.EOF Then
  %>
  <table border="0" cellpadding="2" cellspacing="1" width="500">
    <tr style="font-weight: bold;">
      <td class="Separator" colspan="3">
        Administration</td>
    </tr>
    <tr>
      <td width="125">
        Asset name</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("ComputerName") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Serial number</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("SerialNo") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Asset tag</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("AssetTag") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Country</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("LocationCountry") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Location</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("LocationName") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Location detail</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("fv_SLDE_LocDetail") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Installation date</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("DTInstall") %>
      </td>
    </tr>
    <tr style="font-weight: bold;">
      <td class="Separator" colspan="3">
        <br />
        Costing</td>
    </tr>
    <tr>
      <td width="125">
        OpCo</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("fv_SLDE_BUL") & " " & objRs("fv_SLDE_OpCo") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Cost location</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("CostcenterTitle") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Purchase date</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("DTAcquisition") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Invoice status</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("fv_SLDE_BillingStatus") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Asset status</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("cf_HP_AssgnRead") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Radia status</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("ScannerDesc") %>
      </td>
    </tr>
    <tr style="font-weight: bold;">
      <td class="Separator" colspan="3">
        <br />
        Hardware</td>
    </tr>
    <tr>
      <td width="125">
        Category</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("CategoryName") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Brand</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("Brand") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Model</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("ProductModel") %>
      </td>
    </tr>
    <tr style="font-weight: bold;">
      <td class="Separator" colspan="3">
        <br />
        Main user</td>
    </tr>
    <tr>
      <td width="125">
        Network logon</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("SupervisorUserLogin") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Last name</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("SupervisorName") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        First name</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("SupervisorFirstName") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Phone number</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("SupervisorPhone") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        E-mail address</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("SupervisorEmail") %>
      </td>
    </tr>
    <tr>
      <td width="125">
        Department</td>
      <td width="10">
        :</td>
      <td>
        <%=objRs("SupervisorTitle") %>
      </td>
    </tr>
  </table>
  <%
      If objRs("ChangeCommitted") > 0 Then
  %>
  <p>
    This information is reflecting changes entered at
    <%=objRs("DTMutation") %>
    and this has not yet been synchronised with AssetCenter.</p>
  <%
      End If
    Else
  %>
  <p>
    There is no information available for this query.</p>
  <%
    End If
  Else
  %>
  <p>
    There is no information to look for.</p>
  <%
  End If
  %>

  <script language="javascript" type="text/javascript">
    window.print();
  </script>

</body>
</html>
