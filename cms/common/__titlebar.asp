<!-- START TITLE BAR -->
<tr>
	<td class="title" width="<%=CMS_PAGE_WIDTH%>" align="right">
		<span style="padding:5 0 5 15;float:left;"><img src="<%=IMAGE_PATH & "/" & COMPANY_LOGO%>"></span>
		<span style="padding:15 10 15 0;float:right;"><br>
			<span class="page-title"><%=COMPANY_NAME%></span><br>
			<span class="sub-title"><%=PAGE_TITLE%><%
			If Cbool(Instr(PAGE_URL_FILE,"page-photo-report")) Then
				If WEB_PAGE_NAME <> "" Then Response.Write(" < " & WEB_PAGE_NAME & " >")
			End If
			%></span>
		</span>
	</td>
</tr>
<!-- END TITLE BAR -->