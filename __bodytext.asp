<td class="body-content" valign="top" width="575">
	<table cellspacing="0" border="0" class="body-content-liner" width="575">
		<tr>
			<td align="left" width="575">
				<h1><%=BODY_TITLE%></h1>
				<div style="overflow:auto;height:282;width:500;text-align:justify;line-height:20px;padding:3 5 3 5;SCROLLBAR-FACE-COLOR: #E4E1BD;SCROLLBAR-HIGHLIGHT-COLOR:#B8B37C; SCROLLBAR-SHADOW-COLOR: #B8B37C; SCROLLBAR-3DLIGHT-COLOR: #F5F5DD; SCROLLBAR-ARROW-COLOR: #B8B37C;SCROLLBAR-TRACK-COLOR: #F5F5DD; SCROLLBAR-DARKSHADOW-COLOR: #F5F5DD; SCROLLBAR-BASE-COLOR: #F5F5DD;">
					<% 
					If Instr(PAGE_URL_FILE,"about") Then
						Set oAbout = New cAbout
						Set collAbout = New cAbout
						Call collAbout.GetAboutText()
						For Each oAbout In collAbout.Abouts.Items		
							echo(oAbout.AboutText)
							Set oAbout = Nothing
						Next
						Set collAbout = Nothing						
					ElseIf Instr(PAGE_URL_FILE,"details") Then
						Set oDetails = New cDetails
						Set collDetails = New cDetails
						Call collDetails.GetDetails()
						For Each oDetails In collDetails.Details.Items		
							echo(oDetails.DetailsText)
							Set oDetails = Nothing
						Next
						Set collDetails = Nothing
					End If
					%>
				</div>
			</td>
		</tr>
	</table>
</td>
