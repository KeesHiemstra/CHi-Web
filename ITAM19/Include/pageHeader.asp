<html>
<head>
  <title>ITAMWeb web portal</title>
  <script src="Include/faq.js" type="text/javascript"></script>

  <script type="text/javascript">    
    if (typeof i18n == 'undefined') var i18n = {};
    i18n.TEXT_OPEN_ALL = "Open all";
    i18n.TEXT_CLOSE_ALL = "Close all";
  </script>

  <script language="JavaScript1.2" type="text/javascript">
    function starthelp(strTopic)
    {
      HelpWindow = window.open("help.asp?topic=" + strTopic, "HelpWindow", "resizable=1, menubar=0, toolbar=0, location=0, status=0, scrollbars=1, width=600, height=550");
      HelpWindow.moveTo(0, 0);
      if (window.focus) {HelpWindow.focus()}
    }
    function printasset(strSearch)
    {
      PrintWindow = window.open("acAssetPrint.asp?" + strSearch, "PrintWindow", "resizable=1, menubar=0, toolbar=0, location=0, status=0, scrollbars=1, width=600, height=750");
      PrintWindow.moveTo(0, 0);
      if (window.focus) {PrintWindow.focus()}
    }
    function showimacd(strSearch)
    {
      PrintWindow = window.open("acShowIMACD.asp?" + strSearch, "ShowIMACD", "resizable=1, menubar=0, toolbar=0, location=0, status=0, scrollbars=1, width=600, height=550");
      PrintWindow.moveTo(0, 0);
      ShowIMACD.focus()
    }
  </script>

  <script type="text/javascript" language='javascript1.2' src="Include/PopCalendar.js"></script>
  <link rel="stylesheet" type="text/css" href="Include/Template.css" />
</head>
<body bgcolor="white" leftmargin="0" topmargin="0" onload="if (document.getElementById('maintable') != null) {document.getElementById('maintable').height=document.body.clientHeight-141}">
<script language="JavaScript" type="text/javascript">
<!--
document.title = "IW19 - <%=PageTitle%>"
//-->
</script>

<br>
<table align="center" width="970" border="0" cellpadding="0" cellspacing="0" summary="Page header">

  <tr height="31">
    <!-- td width="87" rowspan="2" valign="top"><a href="Index.asp"><img src="Image/HPInvent.gif" width="87" height="76" border="0" alt="Main menu"></a></!--> <!-- width=110 -->
    <td width="180" rowspan="2" valign="top"><a href="Index.asp"><img src="Image/HPE.png" width="179" height="71" border="0" alt="Main menu"></a></td> <!-- width=110 -->
    <td width="5" rowspan="2">&nbsp;</td>
    <td width="508" class="title" colspan="3">
      ITAM Web portal
    </td>
    <td width="5" rowspan="2">&nbsp;</td>
    <td class="<%Response.Write "rightheader"%>"></td>
  </tr>

  <tr height="76" >
    <td width="87"></td>
    <td width="50">&nbsp;</td>
    <td width="500" class="subtitle" >
     <%=PageTitle%>&nbsp;
    </td>
    <td width="273" valign="top" align="right">
      <% if session("UID") = "" then %>
      <a href="Logon.asp">Logon</a><br/>
      <% else %>
      <a href="Logoff.asp">Logoff</a><br/>
      <% end if %>
      <a href="Index.asp">Menu</a><br>
      <a href="javascript:history.go(-1);">Back</a>
    </td>
  </tr>
</table>