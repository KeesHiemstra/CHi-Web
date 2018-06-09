<%
session("UID") = ""
Session.Contents.RemoveAll()
Session.Abandon
response.redirect "Logon.asp"
%>