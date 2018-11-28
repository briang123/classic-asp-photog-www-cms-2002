<!-- START HEADER BAR -->
<% If loginPage Then %>
	<tr>
		<td class="header">
			<ul>
				<li class="inline"><a href="http://www.juliestarkphotography.com" target="_self">View Website</a></li>
				<li class="inline"><a href="http://support.newleafhosting.com" target="_blank">New Leaf Technical Support</a></li>				
			</ul>
		</td>
	</tr>
<% Else %>
	<tr>
		<td class="header">
			<ul>
				<li class="inline"><a href="<%=CMS_ROOT_PATH%>/logout.asp">Logout</a></li>
				<li class="inline"><a href="<%=CMS_ROOT_PATH%>/gallery-report.asp">Home</a></li>
				<li class="inline"><a href="<%=DOMAIN_NAME%>" target="_blank">View Website</a></li>
				<li class="inline"><a href="http://support.newleafhosting.com" target="_blank">New Leaf Technical Support</a></li>				
				<li class="inline"><a href="http://mail.juliestarkphotography.com" target="_blank">WebMail</a></li>				
				<li class="inline"><a href="http://support.newleafhosting.com/webstats" target="_blank">Site Statistics</a></li>				
				<li class="inline">[Site Statistics:: Site ID=500; User Name=julie@juliestarkphotography.com; Password=julie]</li>				
			</ul>

		</td>
	</tr>
<% End If %>
<!-- END HEADER BAR -->
