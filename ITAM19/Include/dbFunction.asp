<%
'-------------------------------------------------------------------------------
' Database functions
'-------------------------------------------------------------------------------
Dim objConn             'Connection object
Dim objRs               'Recordset object
Dim strConnect          'Connection string to ODBC
Dim strCustomer					'Abbriviation of the customer
Dim bConn2DB            'Connection result

'-- Defaults%>
<!--#include file="dbConnection.asp"--><%
bConn2DB = acConnectDB()

Function acConnectDB()
  bConn2DB = (UCase(TypeName(objConn)) = "CONNECTION")
  acConnectDB = bConn2DB
End Function

Function acOpenDB()
  Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open strConnect
  objConn.CommandTimeout = 600

  If acConnectDB() Then
    Set objRs = Server.CreateObject("ADODB.Recordset")
  End If
  acOpenDB = bConn2DB
End Function

Function acCloseDB()
  Set objRs = Nothing
  Set strConnect = Nothing
  If bConn2DB Then
     objConn.Close
     bConn2DB = False
     Set objConn = Nothing
  End If
End Function

'-----------------------------------------------------------------------------------------
'Returns the dt parameter as string with the format YYYY-MM-DD HH:NN:SS or empty
'-----------------------------------------------------------------------------------------
Function CISODateTime(dt)
  On Error Resume Next
  CISODateTime = CStr(Year(dt)) & "-" & Right("0" & CStr(Month(dt)), 2) & "-" &_
    Right("0" & CStr(Day(dt)), 2) & " " & Right("0" & CStr(Hour(dt)), 2) & ":" &_
    Right("0" & CStr(Minute(dt)), 2) & ":" & Right("0" & CStr(Second(dt)), 2)
  If Err.Number <> 0 Then
    CISODateTime = ""
  End If
  On Error Goto 0
End Function

'-----------------------------------------------------------------------------------------
'Returns the dt parameter as string with the format YYYY-MM-DD or empty
'-----------------------------------------------------------------------------------------
Function CISODate(dt)
  On Error Resume Next
  CISODate = CStr(Year(dt)) & "-" & Right("0" & CStr(Month(dt)), 2) & "-" &_
    Right("0" & CStr(Day(dt)), 2)
  If Err.Number <> 0 Then
    CISODate = ""
  End If
  On Error Goto 0
End Function

'-----------------------------------------------------------------------------------------
'Returns the dt parameter as string with the SQL format 'YYYY-MM-DD HH:NN:SS' or NULL
'-----------------------------------------------------------------------------------------
Function CSQLDateTime(dt)
Dim strDate
	strDate = CISODateTime(dt)
	If strDate = "" Then
		CSQLDateTime = "NULL"
	Else
		If Year(dt) < 1753 Then
			'SQL can't handle dates before 1753-01-01 00:00:00.000
			CSQLDateTime = "NULL"
		Else
			CSQLDateTime = "'" & strDate & "'"
		End If
	End If
End Function

%>

<%
 '-------------------------------------------------------------------------------
 ' Security functions
 '-------------------------------------------------------------------------------

 '-------------------------------------------------------------------------------
 '  function: GetCountry
 '  input
 '    userid , string   : the userid that must be checked
 '  output
 '   string: the country for the user
 '
 '-------------------------------------------------------------------------------
 Function GetCountry( userid)
  Dim secConn
  Dim objSQL
  Dim secSql
  '----- Dim and set the connection --------------------------------------------
	Set secConn = Server.CreateObject("ADODB.Connection")
  Set secSql  = Server.CreateObject("ADODB.Recordset")
 	secConn.Open strSecConn
 	'----- loop throught all the roles for this user------------------------------
  secSql.Open "SELECT * FROM RadiaRIMProd.dbo.webusers WHERE upper(uid) ="& CHR(39) & ucase(userid) & CHR(39), secConn
  if not secSql.EOF  then
    GetCountry=secSql("Country")
  else
    GetCountry=""
  end if
  '----- close the connections --------------------------------------------------
  secSql.Close
  Set secSql = Nothing
  if ucase(TypeName(secConn)) = "CONNECTION" then
    secConn.Close
    Set secConn = Nothing
  end if
 End Function
 %>
