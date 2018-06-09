<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.buffer = True%>
<%Response.Expires = 0%>

<HTML>
<HEAD>
<TITLE>IW19 - ShowIMACD <%=Request("ChangeID")%></TITLE>
</HEAD>
<LINK rel="stylesheet" type="text/css" href="Include/Template.css" />
<BODY leftmargin="2" topmargin="2">

<!--#include file="Include\dbFunction.asp"-->
<!--#include file="Include\pcFunction.asp"-->

<%
Dim strSQL
Call acOpenDB()

strSQL = "SELECT CAST([ActionDate] AS smalldatetime) AS 'Date', ISNULL(NULLIF([NewSerialNr], ''), [OldSerialNr]) AS 'SerialNumber', ISNULL(NULLIF([NewAssetName], ''), [OldAssetName]) AS 'AssetName', CAST([InstallDate] AS smalldatetime) AS 'InstallationDate', CAST([PurchaseDate] AS smalldatetime) AS 'AcquisitionDate', * FROM RadiaRIMProd.dbo.acMACD WHERE [ChangeID] = " & Request("ChangeID")
objRs.Open strSQL, objConn

%>

<table width="500" border="0" cellpadding="2" cellspacing="1">
	<tr><td colspan="2" class="subtitle">Show IMACD details</td></tr>
	<tr><td>Date </td><td><%=CISODateTime(objRs("Date")) %></td></tr>
	<tr><td>IMACD type</td><td><%=objRs("TypeRequest") %></td></tr>
	<tr><td>Internal tag</td><td><%=objRs("InternalTag") %></td></tr>
	<tr><td>Change ID</td><td><%=objRs("ChangeID") %></td></tr>
	<tr><td>Transfer date</td><td><%=objRs("TransDate") %></td></tr>
	<tr><td>Engineer</td><td><%=objRs("OnSiteEng") %></td></tr>
	<tr><td colspan="2"><hr /></td></tr>
	<tr><td>Reference</td><td><%=objRs("EWM") %></td></tr>
	<tr><td>Serial number</td><td><%=objRs("SerialNumber") %></td></tr>
	<tr><td>Asset name</td><td><%=objRs("AssetName") %></td></tr>
	<tr><td>Radia on</td><td><%=objRs("RadiaOn") %></td></tr>
	<tr><td>Category</td><td><%=objRs("Category") %></td></tr>
	<tr><td>Brand</td><td><%=objRs("Brand") %></td></tr>
	<tr><td>Model</td><td><%=objRs("Model") %></td></tr>
	<tr><td>Country</td><td><%=objRs("CountryOfLocation") %></td></tr>
	<tr><td>Location</td><td><%=objRs("LocationOfAsset") %></td></tr>
	<tr><td>Detailed location</td><td><%=objRs("DetailLocation") %></td></tr>
	<tr><td>Billing status</td><td><%=objRs("InvoiceType") %></td></tr>
	<tr><td>Asset status</td><td><%=objRs("Status") %></td></tr>
	<tr><td>OpCo name</td><td><%=objRs("OpCo") %></td></tr>
	<tr><td>Last name</td><td><%=objRs("UserLName") %></td></tr>
	<tr><td>First name</td><td><%=objRs("UserFName") %></td></tr>
	<tr><td>Account</td><td><%=objRs("NTLogon") %></td></tr>
	<tr><td>E-mail address</td><td><%=objRs("UserEMail") %></td></tr>
	<tr><td>Cost center</td><td><%=objRs("CostLoc") %></td></tr>
	<tr><td>Installation date</td><td><%=objRs("InstallationDate") %></td></tr>
	<tr><td>Acquisition date</td><td><%=objRs("AcquisitionDate") %></td></tr>
</table>

<%
objRs.Close

Call acCloseDB()
%>

</BODY>
</HTML>
