<%
'*********************************************************************************************
' Display User first name
'*********************************************************************************************
Public Sub ShowUserFName(strValue)
  Response.Write "<td>First name</td><td><input type=""text"" name=""frmUserFName"" " & _
    "value=""" & trim(strValue) & """ size=""30""/></td>"
End Sub

'*********************************************************************************************
' Display User last name
'*********************************************************************************************
Public Sub ShowUserLName(strValue)
  Response.Write "<td>Last name</td><td><input type=""text"" name=""frmUserLName"" " & _
    "value=""" & trim(strValue) & """ size=""30""/></td>"
End Sub
 
'*********************************************************************************************
' Text Asset name
'*********************************************************************************************
Public Sub ShowTextAssetName(strValue)
  Response.Write "<td>Asset name</td><td><input type=""text"" name=""frmAssetName"" " & _
    "value=""" & trim(strValue) & """ size=""30""/></td>"
End Sub
  
'*********************************************************************************************
' Display Serial Number
'*********************************************************************************************
Public Sub ShowTextSerial(strValue)
  Response.Write "<td>Serial number</td><td><input type=""text"" name=""frmSerial"" " & _
    "value=""" & trim(strValue) & """ size=""30""/></td>"
End Sub

'*********************************************************************************************
' Display NT Logon
'*********************************************************************************************

Public Sub ShowTextAssetName(strValue)
	Response.Write "<td>Asset name</td><td><input type=""text"" name=""frmAssetName"" " & _
    "value=""" & trim(strValue) & """ size=""30""/></td>"
End Sub 
%>