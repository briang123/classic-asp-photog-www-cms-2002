<!-- START PAGE BODY TOOLBAR -->	
<% 
Dim IS_EDIT_PAGE
IS_EDIT_PAGE = Not IS_REPORT_PAGE

If IS_EDIT_PAGE Then %>
<div style="border-bottom:1px solid <%=abelard_border_color%>;width:100%;padding-bottom:20px;">
	<div style="align:left;width:<%=CMS_PAGE_WIDTH%>;">	
		<span style="float:right;padding-right:10px;">
			<A href="<%=REPORT_PAGE%>" class="menu" title="Return to <%=PAGE_TYPE%> Report">
				<IMG height="16" alt="View Report" src="<%=CMS_IMAGE_PATH%>/discthrd.gif" width="16" hspace="5" align="absmiddle">View <%=PAGE_TYPE%> Report
			</A>
		</span>
		<span style="float:right;padding-right:10px;">
			<A href="#" onClick="return checkForm();" class="menu" title="Save and Close">
				<IMG height="16" alt="Save and Close" src="<%=CMS_IMAGE_PATH%>/saveitem.gif" width="16" hspace="5" align="absmiddle">Save and Close
			</A>
		</span>
	</div>
</div>
<% ElseIf IS_REPORT_PAGE Then %>
<div style="border-bottom:1px solid <%=abelard_border_color%>;width:100%;padding-bottom:20px;">
	<div style="align:left;width:<%=CMS_PAGE_WIDTH%>;">
		<span style="float:right;padding-right:10px;">
			<A href="<%=EDIT_PAGE%>" class="menu" title="Add <%=PAGE_TYPE%>"><IMG height="16" alt="Create New" src="<%=CMS_IMAGE_PATH%>/newitem.gif" width="16" hspace="5" align="absmiddle">Add <%=PAGE_TYPE%></A>
		</span>
	</div>
</div>
<% End If %>
<!-- END PAGE BODY TOOLBAR -->