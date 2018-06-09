<%@  language="VBSCRIPT" %>
<% Option Explicit%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = 0%>
<%
Dim PageTitle
Dim MenuSize
Dim strSQL, strLink
PageTitle = "Main menu"
MenuSize = "Small"
%>
<!--#include file="Include/pageHeader.asp"-->
<!--#include file="Include/dbFunction.asp"-->
<%
' Check if the user is logged otherwise redirect to the login page
If Session("UID") = "" Then
  Session("AfterLoginGoto") = Request.ServerVariables("SCRIPT_NAME")
  Response.Redirect "Logon.asp"
  Response.End
End if

If Request("PageAction") <> "" Then
  Session("PageAction") = Request("PageAction")
  Session("SortOrder") = ""  ' Reset Sort order to avoid sorting on a nonexisting column
  Response.Redirect "acAssetList.asp"
  Session.Abandon
  Response.End
End if

%>
<table id="maintable" align="center" cols="3" width="970" border="0" cellpadding="0"
  cellspacing="0" summary="Page body">
  <tr>
    <td width="92" height="50">
      &nbsp;</td>
    <td width="600">
      &nbsp;</td>
    <td>
      &nbsp;</td>
  </tr>
  <tr>
    <td />
    <td valign="top">
      <% If Session("SecAdmin") <> 0 Or Session("SecHP") > 0 Then %>
<!--      <p style="background-color: powderblue">
        Take note of the move from EWM to HPSC. As from September 30, please STOP opening
        new cases in EWM. All case handling will be in HPSC henceforth. All HP transaction
        must be ONLY via HPSC.
        <br />
        Please share this with all your teams. We propose that all teams start using HPSC
        from September 22nd and move completely to HPSC by Sept 30th latest. All EWM cases
        opened before September 30th must be closed in EWM by November 10.
        <br />
        All teams must manage your queues appropriately and we must ensure no backlogs.
        <br />
        If anyone of your team have access issues to HPSC, please Darryl Lloyd.
      </p>-->
      <% End If %>
      <form name="form" action="">
        <input type="hidden" name="submenu" />
        <input type="hidden" name="reportname" />
        <input type="hidden" name="PageAction" />
        <table width="600" cols="1" align="center" border="0" cellpadding="0" cellspacing="0"
          summary="main menu">
          <tr height="15">
            <td class="menutitle">
              &nbsp;</td>
          </tr>
          <tr>
            <td height="10" />
          </tr>
          <% If Session("SecAdmin")<>0 Or Session("SecHP")<>0 Or Session("SecSLDE") <> 0 Then %>
          <tr>
            <td class="menuitems" height="20">
              <a href="javascript:document.form.submenu.value='AdminPages';document.form.submit();">
                <b>Asset admininistration</b></a>
            </td>
          </tr>
          <% if request("submenu") = "AdminPages" then %>
          <tr valign="top">
            <td>
              <table width="100%" summary="Admin menu items">
                <tr>
									<td>
										<ul style="list-style-image:url(/image/MenuIconMenu.gif)">
											<li><a href="javascript:document.form.PageAction.value='INFO';document.form.submit();">View asset details</a></li>

											<% If Session("SecAdmin") <> 0 Or Session("SecHP") >= 1 Or Session("SecSLDE") >= 1 Then %>
												<li><a href="javascript:document.form.PageAction.value='EDIT';document.form.submit();">Edit existing asset</a></li>
											<% End If %>

			                <li><a href="pcLookup.asp">Asset lookup and configuration details</a></li>
										</ul>
                  </td>
                </tr>
								<% If Session("SecAdmin") <> 0 Or Session("SecHP") > 3  Then %>
									<tr>
										<td class="menusubtitle">Special actions</td>
									</tr>
									<tr>
										<td>
											<ul style="list-style-image:url(/image/MenuIconMenu.gif)">
												<% If Session("SecAdmin") <> 0 Or Session("SecHP") > 3  Then %>
													<li><a href="javascript:document.form.PageAction.value='ADD';document.form.submit();">Add new asset</a></li>
												<% End If %>
												<% If Session("SecAdmin") <> 0 Or Session("SecHP") >= 255 Or Session("SecSLDE") >= 255 Then %>
													<li><a href="javascript:document.form.PageAction.value='DELETE';document.form.submit();">Cancel asset modification</a></li>
												<% End If %>
												<% If Session("SecAdmin") <> 0 Or Session("SecHP") >= 15 Then %>
													<li><a href="acDuplicates.asp">Delete duplicate assets/serial numbers</a></li>
												<% End If %>
											</ul>
										</td>
									</tr>
							<% End If %>
							</table>
            </td>
          </tr>
          <% end if %>
          <% end if %>
          <!-- ----------------------------------------------------------------------------------------------------------------------------- -->
          <% if Session("SecAdmin")<>0 or Session("SecHP")<>0 or Session("SecSLDE")<>0 then %>
          <tr>
            <td class="menuitems" height="20">
              <a href="javascript:document.form.submenu.value='ReportsSL';document.form.submit();">
                <b>Reports</b></a>
            </td>
          </tr>
          <% if request("submenu") = "ReportsSL" then %>
          <tr valign="top">
            <td>
              <table width="100%" summary="menu items">
                <tr>
                  <td />
                </tr>
                <tr>
                  <td>
										<ul style="list-style-image:url(/image/MenuIconMenu.gif)">
											<li><a href="javascript:document.form.submenu.value='SoftwareReportsSL';document.form.submit();"><b>Software reports</b></a></li>
                    </ul>
                   </td>
                </tr>

                <tr>
                  <td class="menusubtitle">Asset reports</td>
                </tr>

                <tr>
									<td>
										<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											<li><a href="/Reports/Computer_in_contract.xlsx" target="_blank">Computers in contract</a></li>
											<li><a href="/Reports/Computer_in_stock.xlsx" target="_blank">Computers on stock</a></li>
											<li><a href="/Reports/Computer_in_contract_changes.xlsx" target="_blank">Computers with changed contract status since the last invoice</a></li>
											<!-- <li><a href="/Reports/Printer_in_contract.xlsx" target="_blank">Printers in contract</a></li> -->
										</ul>
									</td>
                </tr>

                <tr>
                  <td class="menusubtitle">Computer reports</td>
                </tr>

                <tr>
									<td>
										<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											<li><a href="/Reports/Computer_in_contract_OpCo_Difference.xlsx" target="_blank">Computers in contract with differences between the Asset OpCo and User OpCo</a></li>
											<li><a href="/Reports/Computer_not_audited_in_Radia.xlsx" target="_blank">Computers not audited for 30 days or more</a></li>
										</ul>
									</td>
                </tr>
<!--
                <tr>
                  <td class="menusubtitle">OS Patching (XP)</td>
                </tr>

                <tr>
                  <td>
										<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											<li><a href="/Reports/Monthly_patch_installation_48h.xlsx" target="_blank">Patch installations 48 hours after releasing patches (published monthly)</a></li>
											<li><a href="/Reports/Monthly_patch_installation.xlsx" target="_blank">Patch installations (published monthly)</a></li>
										</ul>
                  </td>
                </tr>
-->

                <tr>
                  <td class="menusubtitle">Account reports</td>
                </tr>

                <tr>
									<td>
										<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											<li><a href="/Reports/Computer_exceptions.xlsx" target="_blank">Computers with exceptions on user accounts</a></li>
											<li><a href="/Reports/Logon_history.xlsx" target="_blank">Logon history over the last 6 months</a></li>
											<!-- <li><a href="/Reports/User_rights_exceptions.xlsx" target="_blank">User rights exceptions</a></li> -->
										</ul>
									</td>
                </tr>

								<% If strCustomer = "DEMB" Then %>
									<tr>
										<td colspan="2" class="menusubtitle">External reports</td>
									</tr>
                
									<tr>
										<td>
											<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
												<!--
												<li><a href="/Reports/Assets_LostStolen.xlsx" target="_blank">Assets lost and stolen (published monthly)</a></li>
												-->
												<li><a href="Reports/Quarterly_disposed_assets.xlsx" target="_blank">Disposed assets (published quartaly)</a></li>
												<li><a href="Reports/Quarterly_disposed_assets_EnvironmentalCert.pdf" target="_blank">Disposed assets environmental certificate (published quartaly)</a></li>
											</ul>
										</td>
									</tr>
								<% End If %>

<!--
                <tr>
                  <td />
                </tr>
                <tr>
                  <td>
										<ul style="list-style-image:url(/image/MenuIconMenu.gif)">
											<li>
		                    <a href="javascript:document.form.submenu.value='SpecialReportsSL';document.form.submit();"><b>Special reports</b></a>
											</li>
											<li>
												<a href="javascript:document.form.reportname.value='';document.form.submenu.value='ReportArchive';document.form.submit();"><b>Report archive</b></a>
											</li>
                    </ul>
                   </td>
                </tr>
-->
              </table>
            </td>
          </tr>
          <% End If %>

          <% if request("submenu") = "SoftwareReportsSL" then %>
          <tr valign="top">
						<td>
              <table width="100%" summary="menu items">
                <tr>
                  <td style="width:33%" />
                  <td style="width:33%" />
                  <td />
                </tr>
								<tr>
									<td colspan="3">
										<ul style="list-style-image:url(/image/MenuIconMenu.gif)">
											<li><a href="javascript:document.form.submenu.value='ReportsSL';document.form.submit();"><b>Other reports</b></a></li>
										</ul>
									</td>
								</tr>

                <tr>
                  <td colspan="3" class="menusubtitle">Summaries</td>
                </tr>

                <tr>
									<td colspan="3">
										<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											<li><a href="/Reports/Application_count.xlsx" target="_blank">Application counts</a></li>
										</ul>
									</td>
                </tr>

								<% If strCustomer = "DEMB" Then %>
									<tr>
										<td align="center" class="menusubtitle">APJ</td>
										<td align="center" class="menusubtitle">EMEA</td>
										<td align="center" class="menusubtitle">America</td>
									</tr>
									<tr valign="top">
										<td>
											<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
												<li><a href="/Reports/Software_AU.xlsx" target="_blank">Australia</a></li>
												<li><a href="/Reports/Software_CN.xlsx" target="_Blank">China</a></li>
												<li><a href="/Reports/Software_HK.xlsx" target="_blank">Hong Kong</a></li>
												<li><a href="/Reports/Software_ID.xlsx" target="_blank">Indonesia</a></li>
												<li><a href="/Reports/Software_NZ.xlsx" target="_blank">New Zealand</a></li>
												<li><a href="/Reports/Software_TH.xlsx" target="_blank">Thailand</a></li>
											</ul>
										</td>
		
										<td>
											<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
												<li><a href="/Reports/Software_AT.xlsx" target="_blank">Austria</a></li>
												<li><a href="/Reports/Software_BY.xlsx" target="_blank">Belarus</a></li>
												<li><a href="/Reports/Software_BE.xlsx" target="_blank">Belgium</a></li>
												<li><a href="/Reports/Software_BG.xlsx" target="_blank">Bulgaria</a></li>
												<li><a href="/Reports/Software_CZ.xlsx" target="_blank">Czech Republic</a></li>
												<li><a href="/Reports/Software_FR.xlsx" target="_blank">France</a></li>
												<li><a href="/Reports/Software_GE.xlsx" target="_blank">Georgia</a></li>
												<li><a href="/Reports/Software_DE.xlsx" target="_blank">Germany</a></li>
												<li><a href="/Reports/Software_GR.xlsx" target="_blank">Greece</a></li>
												<li><a href="/Reports/Software_HU.xlsx" target="_blank">Hungary</a></li>
												<li><a href="/Reports/Software_IE.xlsx" target="_blank">Ireland</a></li>
												<li><a href="/Reports/Software_IT.xlsx" target="_blank">Italy</a></li>
												<li><a href="/Reports/Software_KZ.xlsx" target="_blank">Kazakhstan</a></li>
												<li><a href="/Reports/Software_LT.xlsx" target="_blank">Lithuania</a></li>
												<li><a href="/Reports/Software_MA.xlsx" target="_blank">Morocco</a></li>
												<li><a href="/Reports/Software_NL.xlsx" target="_blank">Netherlands</a></li>
												<li><a href="/Reports/Software_Nordics.xlsx" target="_blank">Nordics (Denmark, Latvia, Norway, Sweden)</a></li>
												<li><a href="/Reports/Software_PL.xlsx" target="_blank">Poland</a></li>
												<li><a href="/Reports/Software_PT.xlsx" target="_blank">Portugal</a></li>
 												<li><a href="/Reports/Software_RO.xlsx" target="_blank">Romania</a></li>
												<li><a href="/Reports/Software_RU.xlsx" target="_blank">Rusia</a></li>
												<li><a href="/Reports/Software_SK.xlsx" target="_blank">Slovakia</a></li>
												<li><a href="/Reports/Software_ZA.xlsx" target="_blank">South Africa</a></li>
												<li><a href="/Reports/Software_ES.xlsx" target="_blank">Spain</a></li>
												<li><a href="/Reports/Software_CH.xlsx" target="_blank">Switzerland</a></li>
												<li><a href="/Reports/Software_TR.xlsx" target="_blank">Turkey</a></li>
												<li><a href="/Reports/Software_UA.xlsx" target="_blank">Ukraine</a></li>
												<li><a href="/Reports/Software_GB.xlsx" target="_blank">United Kingdom</a></li>
											</ul>
										</td>

										<td>
											<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
												<li><a href="/Reports/Software_BR.xlsx" target="_blank">Brazil</a></li>
											</ul>
										</td>
									</tr>
								<% End If %>

								<% If strCustomer = "HBC" Then %>
									<tr>
										<td align="center" class="menusubtitle">America</td>
										<td align="center" class="menusubtitle"></td>
										<td align="center" class="menusubtitle"></td>
									</tr>
									<tr valign="top">
										<td>
											<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
												<li><a href="/Reports/Software_US_A.xlsx" target="_blank">US (A-Site specific)</a></li>
												<li><a href="/Reports/Software_US_B.xlsx" target="_blank">US (B-Site specific)</a></li>
												<li><a href="/Reports/Software_US_C.xlsx" target="_blank">US (C-Site specific)</a></li>
												<li><a href="/Reports/Software_US_X.xlsx" target="_blank">US (Not site specific)</a></li>
											</ul>
										</td>
		
										<td>
											<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											</ul>
										</td>

										<td>
											<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											</ul>
										</td>
									</tr>
								<% End If %>

<!--
                <tr>
                  <td colspan="3">
										<ul style="list-style-image:url(/image/MenuIconMenu.gif)">
											<li>
		                    <a href="javascript:document.form.submenu.value='SpecialReportsSL';document.form.submit();"><b>Special reports</b></a>
											</li>
										</ul>
									</td>
                </tr>
-->
             </table>
            </td>
          </tr>
          <% end if %>

          <% if request("submenu") = "SpecialReportsSL" then %>
          <tr>
            <td height="1">
              <table width="100%" summary="Special reports menu items">
                <tr>
                  <td />
                </tr>

                <tr>
                  <td>
										<ul style="list-style-image:url(/image/MenuIconMenu.gif)">
											<li><a href="javascript:document.form.submenu.value='ReportsSL';document.form.submit();"><b>Reports</b></a></li>
                    </ul>
                   </td>
                </tr>

                <tr>
                  <td class="menusubtitle">Special reports</td>
                </tr>

                <tr valign="top">
                  <td>
										<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											<li><a href="Reports/Computer_registration_exceptions.xls" target="_blank">Computer registration exceptions</a></li>
											<li><a href="Reports/Disk_space_on_logical_drives.xls" target="_blank">Disk space on logical drives</a></li>
											<li><a href="Reports/VDI_IBM_in_contract.xls" target="_blank">Virtual machines assigned to IBM</a></li>
											<li><a href="Reports/SAPVersions.xls" target="_blank">SAP versions Report</a></li>
											<li><a href="Reports/Assets_to_be_deleted_from_AD.xls" target="_blank">Assets to be deleted from AD</a></li>
										</ul>
                  </td>
                </tr>
                <tr>
                  <td class="menusubtitle">Quarterly published external reports</td>
                </tr>
                <tr>
									<td>
										<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											<li><a href="Reports/Quarterly_disposed_assets.xlsx" target="_blank">Disposed assets</a></li>
											<li><a href="Reports/Quarterly_disposed_assets_EnvironmentalCert.pdf" target="_blank">Disposed assets environmental certificate</a></li>
										</ul>
									</td>
								</tr>
<!--
                <tr>
                  <td class="menusubtitle">External reports</td>
                </tr>
                <tr>
									<td>
										<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											<li><a href="/Reports/SSPR.xls" target="_blank">SSPR report</a></li>
										</ul>
									</td>
								</tr>
-->
              </table>
            </td>
          </tr>
          <% end if %>
          <% if request("submenu") = "ReportArchive" then %>
          <tr valign="top">
						<td>
              <table width="100%" summary="menu items">
                <tr/>
								<tr>
									<td>
										<ul style="list-style-image:url(/image/MenuIconMenu.gif)">
											<li><a href="javascript:document.form.submenu.value='ReportsSL';document.form.submit();"><b>Other reports</b></a></li>
										</ul>
									</td>
								</tr>

								<% If Request("reportname") = "" Then %>
									<tr>
										<td class="menusubtitle">Select report archive</td>
									</tr>

									<tr>
										<td>
											<ul style="list-style-image:url(/image/MenuIconMenu.gif)">
												<li><a href="javascript:document.form.submenu.value='ReportArchive';document.form.reportname.value='Assets_in_contract.xls';document.form.submit();"><b>Assets in contract</b></a></li>
												<li><a href="javascript:document.form.submenu.value='ReportArchive';document.form.reportname.value='Assets_in_contract_changes.xls';document.form.submit();"><b>Assets in contract change report</b></a></li>
												<li><a href="javascript:document.form.submenu.value='ReportArchive';document.form.reportname.value='Quarterly_disposed_assets.xls';document.form.submit();"><b>Disposed assets</b></a></li>
												<li><a href="javascript:document.form.submenu.value='ReportArchive';document.form.reportname.value='Quarterly_disposed_assets_DiskwipeCertificate.pdf';document.form.submit();"><b>Disposed assets Diskwipe Certificate</b></a></li>
												<li><a href="javascript:document.form.submenu.value='ReportArchive';document.form.reportname.value='Quarterly_disposed_assets_EnvironmentalCertificate.pdf';document.form.submit();"><b>Disposed assets Environmental Certificate</b></a></li>
											</ul>
										</td>
									</tr>

								<% Else %>
									<tr>
										<td class="menusubtitle">Select archive of <%=Request("reportname") %></td>
									</tr>

									<tr>
										<td>
											<%
											acOpenDB()
											strSQL = "SELECT TOP 13 * FROM webReportArchive WHERE [ReportName] = '" & Request("reportname") & "' ORDER BY [DTCreation] DESC"
											
											objRs.Open strSQL, objConn
											If objRs.EOF Then
												%> No documents <%
											Else
												%> <ul style="list-style-image:url(/image/MenuIconExternal.gif)"> <%
												While Not objRs.EOF
													%> <li><a href="<%=Replace(objRs("ArchiveName"), "\\DESHNGAPS073\Reports$", "") %>" target="_blank"><%=objRs("ArchiveDesc") %></a></li> <%
												
													objRs.MoveNext
												Wend
												%> </ul> <%
											End If
											objRs.Close
											
											acCloseDB()
											%>

										</td>
									</tr>
									<tr>
										<td>
											<ul style="list-style-image:url(/image/MenuIconMenu.gif)">
												<li><a href="javascript:document.form.reportname.value='';document.form.submenu.value='ReportArchive';document.form.submit();"><b>Report archive</b></a></li>
											</ul>
										 </td>
									</tr>
								<% End If %>

	            </table>
            </td>
          </tr>
          <% end if %>

          <% end if %>
          <!-- --HP Reports----------------------------------------------------------------------------------------------------------------- -->
          <% If Session("SecAdmin") <> 0 OR Session("SecHP") <> 0 Then %>
          <tr>
            <td class="menuitems" height="20">
              <a href="javascript:document.form.submenu.value='ReportsHP';document.form.submit();">
                <b>Reports for HP</b></a>
            </td>
          </tr>
          <% If request("submenu") = "ReportsHP" Then %>
          <tr valign="top">
            <td>
              <table summary="MDE Reports for HP menu items" width="100%">
								<tr>
									<td />
                </tr>
                <tr>
                  <td>
										<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											<li><a href="/ReportsHP/Computer_not_in_contract.xlsx" target="_blank" type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">Computers not in contract (HP)</a></li>
											<li><a href="/ReportsHP/Computer_exceptions_HP.xlsx" target="_blank">Computer registration exceptions (HP)</a></li>
											<li><a href="/ReportsHP/Computer_not_audited_in_Radia_HP.xlsx" target="_blank">Computers not audited for 30 days or more</a></li>
											<li><a href="/ReportsHP/Logon_history_HP.xlsx" target="_blank">Logon history over the last 6 months</a></li>

											<li><a href="/ReportsHP/Asset_duplicates.xlsx" target="_blank">Duplicate assets registration</a></li>
										</ul>
                  </td>
                </tr>

								<tr>
									<td>
										<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											<li><a href="/ReportsHP/Managed_refresh_current_wave.xlsx" target="_blank">Managed Refresh progress report for current wave</a></li>
											<li><a href="/ReportsHP/Managed_refresh_completed.xlsx" target="_blank">Managed Refresh completed</a></li>
											<li><a href="/ReportsHP/Managed_refresh_last_waves.xlsx" target="_blank">Managed Refresh progress report for previous 2 waves</a></li>
										</ul>
									</td>
								</tr>

<!--
								<tr>
									<td>
										<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											<li><a href="/ReportsHP/Managed_reimage.xlsx" target="_blank">Reimage Windows 7</a></li>
										</ul>
									</td>
								</tr>

								<tr>
									<td>
										<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											<li><a href="/Reports/HP/Managed_Refresh_Wave_Legacy-04_(HP).xls" target="_blank">Legacy Refresh progress report for Wave 04</a></li>
											<li><a href="/Reports/HP/Managed_Refresh_Wave_Legacy-03_(HP).xls" target="_blank">Legacy Refresh progress report for Wave 03</a></li>
										</ul>
									</td>
								</tr>

								<tr>
									<td>
										<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											<li><a href="/Reports/HP/Early_Termination_Finnigan_(HP).xls" target="_blank">Early Termination Finnegan</a></li>
										</ul>
									</td>
								</tr>
-->
<!--
                <tr>
                  <td>
										<ul style="list-style-image:url(/image/MenuIconMenu.gif)">
											<li><a href="javascript:document.form.submenu.value='SpecialReportsHP';document.form.submit();"><b>Special reports</b></a></li>
										</ul>
									</td>
                </tr>
-->
              </table>
            </td>
          </tr>
          <% end if %>
          <% end if %>
          <% if request("submenu") = "SpecialReportsHP" then %>
          <tr>
            <td>
              <table summary="Special reports menu items" width="100%">
                <tr>
                  <td />
                </tr>
                <tr>
                  <td>
										<ul style="list-style-image:url(/image/MenuIconMenu.gif)">
											<li><a href="javascript:document.form.submenu.value='ReportsHP';document.form.submit();"><b>Reports</b></a></li>
										</ul>
                  </td>
                </tr>
								<tr>
									<td class="menusubtitle">Special reports</td>
								</tr>
								<tr>
									<td>
										<ul style="list-style-image:url(/image/MenuIconExternal.gif)">
											<!--<li><a href="/Reports/AssetDetailsComputers.xls" target="_blank">Asset details in contract (Computers only)</a></li>-->
											<!--<li><a href="/Reports/AssetDetailsAll.xls" target="_blank">Asset details in contract (All assets)</a></li>-->
											<!--<li><a href="/Reports/HP/Assets_as_MACD_(HP).xls" target="_blank">Assets in MACD template format (HP)</a></li>-->
											<li><a href="/Reports/SAVReport.xls" target="_blank">Details on Symantec AntiVirus per computer</a></li>
										</ul>
									</td>
								</tr>
							</table>
            </td>
          </tr>
          <% end if %>
          <!-- ----------------------------------------------------------------------------------------------------------------------------- 
          <% 
          If Session("SecAdmin")<>0 or Session("SecHP")<>0 or Session("SecSLDE")<>0 Then %>
          <tr>
            <td class="menuitems" height="20">
              &raquo; <a href="javascript:document.form.submenu.value='TECHNICAL';document.form.submit();">
                Technical support</a>
            </td>
          </tr>
          <%
            If request("submenu") = "TECHNICAL" Then
              If Session("SecAdmin")<>0 or Session("SecHP")<>0 then
                response.redirect "SupportHP.asp"
                response.end
              Else
                If  Session("SecSLDE")<>0 then
                  response.redirect "SupportSLI.asp"
                  response.end
                Else
                  response.redirect "support/support1.asp"
                  response.end
                End If
              End If
            End If
          End if %>
          -->
          <!-- ----------------------------------------------------------------------------------------------------------------------------- -->
          <% 
          If Session("SecAdmin")<>0 Or Session("SecHP")<>0 Or Session("SecSLDE") <> 0 Then %>
					<!--
          <tr>
            <td class="menuitems" height="20">
              <a href="javascript:document.form.submenu.value='RadiaInfo';document.form.submit();">
                <b>RadiaInfo web portal information</b></a>
            </td>
          </tr>
					-->
          <% 
            If Request("submenu") = "RadiaInfo" Then %>
          <tr valign="top">
            <td>
              <table summary="RadiaInfo menu items" width="100%">
                <tr>
                  <td />
                </tr>

								<tr>
									<td>
										<ul style="list-style-image:url(/image/MenuIconMenu.gif)">
											<li><a href="WPIChangeLog.asp">Change log</a></li>
											<li><a href="WPI_FAQ.asp">Frequently Asked Questions</a></li>
											<li><a href="/Docs/RadiaInfo web portal manual.pdf" target="_blank">RadiaInfo web portal manual</a></li>
									</ul>
									</td>
								</tr>
              </table>
            </td>
          </tr>
          <%
            End If %>
          <% 
          End If %>
          <!-- ----------------------------------------------------------------------------------------------------------------------------- -->
          <tr>
            <td height="15" class="menutitle" width="600">
              &nbsp;
            </td>
          </tr>
        </table>
      </form>
			<% 
				If Request("submenu") = "" Then
					If (Now() > #12/21/2010#) And (Now() < #12/27/2010#) Then
					%>
						<br />
						<p style="text-align:center;font-size:18pt;color:Red">Merry Christmas<br />
						<img src="Image/xmas_tree.gif" style="text-align:center" alt="X-mas tree" />
						<br />from the HP RadiaInfo team
						</p>
					<%
					End If
					If (Now() > #01/01/2012#) And (Now() < #01/07/2012#) Then
					%>
						<br />
						<p style="text-align:center;font-size:18pt;color:Red">Happy new year<br />
						<img src="Image/Champagne.gif" style="text-align:center" alt="Champagne" height="50%" />
						<br />from the HP RadiaInfo team
						</p>
					<%
					End If
				End If
			%>
    </td>
    <td />
  </tr>
  <tr>
    <td colspan="3">
      &nbsp;</td>
  </tr>
</table>
<!--#include file="Include/pageFooter.asp"-->
