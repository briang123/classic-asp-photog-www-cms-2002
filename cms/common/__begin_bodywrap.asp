	<tr>
		<td width="100%">
			<table width="100%" border="0" cellpadding="0" cellspacing="0" height="100%">
				<tr>
					<td id="leftnav">
						<!-- #include virtual="/cms/sidenav.asp" -->
					</td>
					<td id="mainbody">		
						<!-- #include virtual="/cms/common/__toolbar.asp" -->											
						<table border="0" cellpadding="0" cellspacing="0" width="<%=CMS_PAGE_WIDTH%>">
							<tr>
								<td width="50" valign="top"><br><img src="<%=CMS_IMAGE_PATH %>/<%=PAGE_IMAGE%>"></td>
								<td style="width:auto;">
									<p class="admin-instruction"><% If IS_EDIT_PAGE Then echo(EDIT_INSTRUCTIONS) Else echo(REPORT_INSTRUCTIONS) %></p>
									<p style="color:red;"><%If StringNotEmptyOrNull(displayMessage) Then echo("<br>" & displayMessage)%></p>
