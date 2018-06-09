<%@LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%PageTitle = "Logon"%>
<%MenuSize = "Small"%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<!--#include file="include/dbFunction.asp"-->
<!--#include file="include/globalVars.asp"-->
<%
 ' There are a 4 situations in which the user opens this page:
 ' 1. The user directly opens this page. Then all variables are empty
 ' 2. The user opens an other page and is not logged on yet. That page wil set the AfterLoginGoto variable and redirect to this page.
 ' 3. The user just entered the correct username/password combination
 ' 4. The user just entered an incorrect username/password cobination

 ' What should this page do in each case?

 ' 1. Set the AfterLoginGoto variable to refresh_frst.asp and Show the login page.
 '    After a successfull login this page will redirect to the menu page refresh_frst.asp.
 '    * In the future the page should redirect directly to the Main menu. Login will occure after the mainmenu if required.

 ' 2. Show the login page and let the user try to login. After that option 3 or 4 will occur. (Page will be reloaded)

 ' 3. Check the username/password combination. This is in this case correct so redirect to the page stored in the variable AfterLoginGoto

 ' 4. Check the username/password combination. This is in this case incorrect so display an error message and let the user try again.

 ' If the Incorrect Login Count is empty then set it to "0"
 ' Define the next page after successfull login if the user opens this page from any other website or from the favorites.
If Session("AfterLoginGoto") = "" Then Session("AfterLoginGoto") = "Index.asp"

 ' Store the UserID into a session variable to make it available on other pages during that session.
Dim strUserPwd
Dim strUserName
Dim strSQL

Call acOpenDB

strUserName = trim(request("UID"))
strUserPwd = request("password")
If strUserName <> "" and strUserPwd <> "" then
  'If the user has entered a UserID and a Password, check if the user has access
  'Call acOpenDB()
  ' Check to see if this userID exists in your database.
  objRs.Open "SELECT * FROM RadiaRIMProd.dbo.webUsers WHERE UID='" & strUserName & "'", objConn, 3, 3
  If Not objRs.EOF then
    Session("fMessage") = ""
    If (StrComp(strUserPwd, objRs("password"), 1) = 0) Then
      ' Password is correct. Set some session variables and redirect the user the page which is stored in the variable AfterLoginGoto.

'      objRs("LogonCount") = objRs("LogonCount") + 1
'      objRs("DTLastLogon") = now() 
'      objrs.update

      objConn.Execute "UPDATE RadiaRIMProd.dbo.webUsers SET [LogonCount] = [LogonCount]+1, [DTLastLogon]=GETDATE() WHERE [UID]='" & strUserName & "'"
     
      Session("SecHP") = objRs("SecHP")
      Session("SecAdmin") = objRs("SecAdmin")
      Session("SecSLDE") = objRs("SecSLDE")
      Session("SecGuest") = objRs("SecGuest")
      Session("CountryAccess") = objRs("CountryAccess")
      Session("AuthLevel") = objRs("AuthLevel")
      Session("SearchResult") = objRs("SearchResult")
      Session("UID")= UCase(strUserName)
      Response.Redirect Session("AfterLoginGoto")
      Response.End
    End If
  End If
  ' Username is not found in the database or the password does not match with the username
  Session("fMessage") = "Username and/or Password incorrect"
  Session("IncorrectLoginCount") = cint(session("IncorrectLoginCount")) + 1
  Session("UID") = ""
  ' Check how many times the user tried to login with an incorrect username/password combination
  If session("IncorrectLoginCount") = "3" then
    %><script>alert('You tried to login three times, session aborted');</script><%
    Session("fMessage") = ""
    Session.abandon
    Response.End
  End if
  objRs.Close
End If
%>

<!--#include file="include/pageHeader.asp"-->
<table id="maintable" align="center" cols=3 width=970 border=0 cellpadding="0" cellspacing="0" summary="Page body">
  <tr><td width=92 height=50 >&nbsp;</td><td width=600 >&nbsp;</td><td>&nbsp;</td></tr>
  <tr><td /><td valign=top>

  <form name="form" method="post" action="">

    <table align="center" border=1 cellpadding=5 cellspacing=0 summary="Logon">
      <tr bordercolor="#ffffff" height="25">
        <td class="menutitle" width="102" colspan="2">Logon</td>
      </tr>
      <tr bordercolor="#000000" height="75">
        <td width="150">
          Username:<br><br><br>
          Password:
        </td>
        <td width="150" align="center">
          <input type="text" class="textbox" size="20" name="UID" value=<%=request("UID")%>><br><br>
          <input type="password" name="password" class="textbox" size="20">
        </td>
      </tr>
      <tr bordercolor="#ffffff" align="right" height="15">
        <td colspan="2" align="right">
          <span class="alert">&nbsp;<%=session("fMessage")%></span>
        </td>
      </tr>
      <tr bordercolor="#ffffff" align="right" height="15">
        <td colspan="2">
					<input type="submit" value="Logon" size="20" class="LoginBtn">
        </td>
      </tr>
    </table>
  </form><script type="text/javascript">document.form.UID.focus();</script>

<!--
	<div style="font-size:large">
		The ITAMWeb portal is currently down for maintenance
	</div>
-->
  </td><td /></tr>
  <tr><td colspan=3>&nbsp;</td></tr>
    <tr><td></td><td>
  <%
    objRs.Open "SELECT * FROM RadiaRIMProd.dbo.webNewsShow WHERE Category IS NULL", objConn
    If Not objRs.EOF Then
  %>
    <table width=600 align="center" summary="RadiaInfo news">
    <tr><th colspan=2 bgcolor="silver">RadiaInfo news</th></tr>
    <%
      While Not objRs.EOF
        Response.Write ("<tr bgcolor=""yellow""><td valign=""top"">" & strMonth(Month(objRs("DTCreation"))) & "&nbsp;" & Day(objRs("DTCreation")) & "</td>")
        Response.Write ("<td>" & objRs("NewsHeader") & "</td></tr>")
        objRs.MoveNext
      Wend
    %>
    </table>
  <%
    End If
    objRs.Close
  %>

  <%
		strSQL = "IF NOT EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[webLockedAccounts]')) "
		strSQL = strSQL & "SELECT TOP 0 -1 AS 'LockedAccounts' ELSE SELECT * FROM dbo.webLockedAccounts"

    objRs.Open strSQL, objConn
    If Not objRs.EOF Then
  %>
    <table width=600 align="center" summary="Account locks"><tr><td>The number of Active Directory account lockouts of today:
    <%
      Response.Write(objRs("LockedAccounts"))
    %>
    </td></tr></table>
  <%
    End If
    objRs.Close
  %>
  <!-- Closing maintable -->
  </td><td></td></tr>
</table>

<!--#include file="include/pageFooter.asp"-->
