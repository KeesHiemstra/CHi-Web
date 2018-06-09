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

Dim PageTitle
Dim MenuSize
Dim strSQL

PageTitle = "Asset lookup/Configuration details"
MenuSize="Large"
%>

<!--#include file="Include/pageHeader.asp"-->
<!--#include file="Include/dbfunction.asp"-->
<!--#include file="Include/pcFunction.asp"-->
<br/>
<table align="center" cols="4" width="970">
<tr><td colspan="4">note: wildcards are denoted by %</td></tr>
<%
'-----------------------------------------------------------------------------
' This page does search for a pc from the SQL database
' based on search criteria it will show a list or transfer control to a detail page
' Version History
' 22 September 2003
'-----------------------------------------------------------------------------

'-----------------------------------------------------------------------------
' Declaration of variables
'-----------------------------------------------------------------------------
Dim bSubmitted '-- see if we come from submit or not
Dim iCountPC   '-- count of number of pc's that match criteria
Dim strPCName  '-- pc identifier

Call acOpenDB()

'-----------------------------------------------------------------------------
' Subroutines
'-----------------------------------------------------------------------------
Sub ShowMainForm(strUserFName, strUserLName, strAssetName, strSerial)
%>
   <form action="pcLookup.asp" method="post">

   <tr><td width="30%"/><td width="30%"/><td /></tr>
   <tr>
     <td height="15" class="menutitle" width="600" colspan="4">
      &nbsp;
     </td>
    </tr>
   <tr><% ShowUserLName(trim(strUserLName)) %></tr>
   <tr><% ShowUserFName(trim(strUserFName))%></tr>
   <tr><% ShowTextAssetName strAssetName %></tr>
   <tr><% ShowTextSerial trim(strSerial) %></tr>
   <tr><td></td><td> <input type="submit" name="frmSearch" value="Search"/></td></tr>
   <tr>
     <td height="15" class="menutitle" width="600" colspan=4>
      &nbsp;
     </td>
    </tr>
   </form>
<%
End Sub

bSubmitted= (Request.Form("frmSearch")="Search")

If (not bSubmitted) Then
    '-----------------------------------------------------------------------------
    ' If we are not submitting data then we display a totally blank form
    '-----------------------------------------------------------------------------
    ShowMainForm "","","",""
Else
    '-------------------------------------------------------------------------------
    ' If we enter the page from a submission we process the form
    ' - if no pc qualifies then give back a search page
    ' - if only one PC qualifies we have sufficient data and can show it
    ' - if less than 100 pc's qualify we can show a list
    ' - if more than 100 pc's qualify then we will ask to refine search
    '-------------------------------------------------------------------------------
    '----first see if we can get a hitcount of matching pc's
   strSQL="SELECT COUNT(*) As Frequency FROM RadiaRIMProd.dbo.acAssetList WHERE "
    If Request.Form("frmUserFName")="" Then
        strSQL=strSQL & " SuperVisorFirstName LIKE " & Chr(39) & "%" & Chr(39) & " AND "
    Else
        strSQL=strSQL & " SuperVisorFirstName LIKE " & Chr(39) & trim(Request.Form("frmUserFName")) & Chr(39) & " AND "
    End if
    If Request.Form("frmUserLName")="" Then
        strSQL=strSQL & " SuperVisorName LIKE " & Chr(39) & "%" & Chr(39) & " AND "
    Else
        strSQL=strSQL & " SuperVisorName LIKE " & Chr(39) & Request.Form("frmUserLName") & Chr(39) & " AND "
    End if
    If Request.Form("frmAssetName")="" Then
       strSQL=strSQL & " Computername LIKE " & Chr(39) & "%" & Chr(39) & " AND "
    Else
       strSQL=strSQL & " Computername LIKE " & Chr(39) & trim(Request.Form("frmAssetName")) & Chr(39) & " AND "
    End if
    If Request.Form("frmSerial") = "" Then
       strSQL=strSQL & " SerialNo LIKE " & Chr(39) & "%" & Chr(39)
    Else
      If LCase(Request.Form("frmSerial")) = "::emptysn" Then
        strSQL = strSQL & " SerialNo = ''"
      Else
        strSQL = strSQL & " SerialNo LIKE " & Chr(39) & Request.Form("frmSerial") & Chr(39)
      End If
    End if
    objRs.Open strSQL, objConn
    iCountPC=objRs("Frequency")
    objRs.Close
    If  iCountPC=0 Then
        '------------ nothing found-----------------------------------------------------------------------------
        Response.Write "<H3>No values found, please change search criteria</H3>"
        ShowMainForm Request.Form("frmUserFName"),Request.Form("frmUserLName"),Request.Form("frmAssetName"),Request.Form("frmSerial")
    Else    If iCountPC>0 and iCountPC<101 Then
        '------------ a list found------------------------------------------------------------------------------
        Response.Write "<H3>A list found, refine criteria or select asset</H3>"
        ShowMainForm Request.Form("frmUserFName"),Request.Form("frmUserLName"),Request.Form("frmAssetName"),Request.Form("frmSerial")
        '------------ show the list
        strSQL="SELECT * FROM RadiaRIMProd.dbo.acAssetList WHERE "
        If Request.Form("frmUserFName")="" Then
            strSQL=strSQL & " SuperVisorFirstName LIKE " & Chr(39) & "%" & Chr(39) & " AND "
        Else
            strSQL=strSQL & " SuperVisorFirstName LIKE " & Chr(39) & trim(Request.Form("frmUserFName")) & Chr(39) & " AND "
        End if
        If Request.Form("frmUserLName")="" Then
          strSQL=strSQL & " SuperVisorName LIKE " & Chr(39) & "%" & Chr(39) & " AND "
        Else
          strSQL=strSQL & " SuperVisorName LIKE " & Chr(39) & Request.Form("frmUserLName") & Chr(39) & " AND "
        End if
        If Request.Form("frmAssetName")="" Then
           strSQL=strSQL & " Computername LIKE " & Chr(39) & "%" & Chr(39) & " AND "
        Else
           strSQL=strSQL & " Computername LIKE " & Chr(39) & trim(Request.Form("frmAssetName")) & Chr(39) & " AND "
        End if
        If Request.Form("frmSerial")="" Then
           strSQL=strSQL & " SerialNo LIKE " & Chr(39) & "%" & Chr(39)
        Else
          If LCase(Request.Form("frmSerial")) = "::emptysn" Then
            strSQL = strSQL & " SerialNo = ''"
          Else
            strSQL = strSQL & " SerialNo LIKE " & Chr(39) & Request.Form("frmSerial") & Chr(39)
          End If
        End if
        strSQL=strSQL & " ORDER By ComputerName"
        Response.Write "<tr><td><b>Last name</b></td><td><b>First name</b></td><td><b>Asset name</b></td><td><b>Serial number</b></td></tr>"
        objRs.Open strSQL, objConn
        While Not objRs.EOF
            Response.Write "<tr><td>" & objRs("SuperVisorName")& "</td>"
            Response.Write "<td>" & objRs("SuperVisorFirstName")& "</td>"
            Response.Write "<td><A href=""" & "pcdetail.asp?frmInternalTag="& objRs("InternalTag")& """>"& objRs("ComputerName")& "</A></td>"
            Response.Write "<td>" & objRs("SerialNo")& "</td></tr>"
            objRs.MoveNext
        Wend
        objRs.Close
        Response.Write "<tr><td height=""15"" class=""menutitle"" width=""600"" colspan=4>&nbsp;</td></tr>"
    Else
      '------------ too much found----------------------------------------------------------------------------
      Response.Write "<H3>Too much entries found, please refine your criteria</H3>"
      ShowMainForm Request.Form("frmUserFName"),Request.Form("frmUserLName"),Request.Form("frmAssetName"),Request.Form("frmSerial")
    ' finish of complex If statement
            End If
    End if

End If '---------If (bSubmitted)---------

Call acCloseDB()
%>
</table>

<!--#include file="Include\pageFooter.asp"-->