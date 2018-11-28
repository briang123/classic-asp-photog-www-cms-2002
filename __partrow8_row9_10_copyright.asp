		<td><img src="images/generic_11.jpg" width="50" height="359" alt=""></td>
		<td><img src="images/spacer.gif" width="1" height="359" alt=""></td>
	</tr>
	<tr><!-- ROW 9 -->
		<td colspan="13"><img src="images/generic_12.jpg" width="899" height="35" alt=""></td>
		<td><img src="images/spacer.gif" width="1" height="35" alt=""></td>
	</tr>
	<tr><!-- ROW 10 -->
		<td><img src="images/spacer.gif" width="23" height="1" alt=""></td>
		<td><img src="images/spacer.gif" width="37" height="1" alt=""></td>
		<td><img src="images/spacer.gif" width="125" height="1" alt=""></td>
		<td><img src="images/spacer.gif" width="69" height="1" alt=""></td>
		<td><img src="images/spacer.gif" width="89" height="1" alt=""></td>
		<td><img src="images/spacer.gif" width="88" height="1" alt=""></td>
		<td><img src="images/spacer.gif" width="117" height="1" alt=""></td>
		<td><img src="images/spacer.gif" width="93" height="1" alt=""></td>
		<td><img src="images/spacer.gif" width="69" height="1" alt=""></td>
		<td><img src="images/spacer.gif" width="41" height="1" alt=""></td>
		<td><img src="images/spacer.gif" width="96" height="1" alt=""></td>
		<td><img src="images/spacer.gif" width="2" height="1" alt=""></td>
		<td><img src="images/spacer.gif" width="50" height="1" alt=""></td>
		<td><img src="images/spacer.gif" width="1" height="1" alt=""></td>
	</tr>
</table>
<table class="copyright-row" align="center">
<tr>
    <td align="left" style="padding-left:20;" width="430"><a href="http://www.newleaftechinc.com" target="_blank">developed by New Leaf Technologies, Inc</a></td>
    <td align="right" width="450"><%=RenderCopyright(Year(Now))%></td>
</tr>
</table>
<% 
Function RenderCopyright(year)
    if year = 2006 then
        RenderCopyright = "Copyright&copy; 2006, Stark Photography, All rights reserved"
    else
        RenderCopyright = "Copyright&copy; 2006 - " & year & " Stark Photography, All rights reserved"
    end if
End Function
%>
