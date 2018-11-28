<% Call HTMLComment("FOOTER SECTION",1) %>
				<tr bgcolor="<%=strFooterColor%>">
					<td width="990px" colspan="100%">
						<table border="<%=intBorder%>" width="100%" cellpadding="0" cellspacing="0">
							<tr>
								<td width="50%" align="left">&nbsp;<%=Application("__AppCopyright")%></td>
								<td width="50%" align="right"><a href="./contact.asp">
									<img src="<%=Application("InvoiceAppImagePath")%>/main_send.gif" border="0">Contact Us</a>&nbsp;|&nbsp;
									phone: <%=Application("__AppCompanyPhone")%>&nbsp;|&nbsp;
									fax: <%=Application("__AppCompanyFax")%>
								</td>
							</tr>
						</table>					
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<% 
Call HTMLComment("FOOTER SECTION",2) 
Call HTMLComment("MAIN SECTION",2)
Call CloseDB() 
%>