<%
IsLoggedIn = False
If GetSessionVariable("PROOF_GALLERY_LOCKED") = "NO" Then 
    IsLoggedIn = True
End If
%>
	<tr><!-- ROW 1 -->
		<td valign="top"  colspan="2" rowspan="6"><img src="images/generic_01.jpg" width="60" height="158" alt=""></td>
		<td valign="top"  rowspan="6"><img src="images/logo.jpg" width="125" height="158" alt=""></td>
		<td valign="top"  colspan="10"><img src="images/generic_02.jpg" width="714" height="25" alt=""></td>
		<td valign="top" ><img src="images/spacer.gif" width="1" height="25" alt=""></td>
	</tr>
	<tr><!-- ROW 2 -->
		<td valign="top"  rowspan="5"><img src="images/generic_03.jpg" width="69" height="133" alt=""></td>
		<td valign="top"  colspan="5" style="height: 17px"><img src="images/generic_04.jpg" width="456" height="17" alt=""></td>
		<%  If WEB_PAGE_ID = 6 Then     'Login Page 
		        If IsLoggedIn Then      'Check if logged in
		            PageRedirect("logout.asp")
		        Else %>
		<td valign="top"  colspan="2" rowspan="2"><a href="login.asp"><img src="images/clientlogin-over.jpg" width="137" height="40" border="0" alt=""></a></td>
		<%      End If %>
		<%  ElseIf WEB_PAGE_ID = 8 Then 'Proofs Page
		        If IsLoggedIn Then      'Check if logged in (if not logged in and on proofs page, then redirect to login page) %>
		<td valign="top"  colspan="2" rowspan="2"><a href="proofs.asp"><img src="images/proofs-sel.jpg" width="137" height="40" border="0" alt=""></a></td>
		<%      Else 
		            PageRedirect("login.asp")
		        End If 
		    Else                        'Any page other than Login or Proofs page
		        If IsLoggedIn Then      'Check if logged in on any page other than Login or Proofs %>
        <!--<td valign="top"  colspan="2" rowspan="2"><a href="proofs.asp"><img src="images/proofs-sel.jpg" width="137" height="40" border="0" alt=""></a></td>-->
        <td valign="top"  colspan="2" rowspan="2"><a href="proofs.asp" onmouseover="changeImages('proofs', 'images/proofs-over.jpg'); return true;" onmouseout="changeImages('proofs', 'images/proofs-index_sel.jpg'); return true;" onclick="changeImages('proofs', 'images/proofs-over.jpg');"><img name="proofs" src="images/proofs-index_sel.jpg" width="137" height="40" border="0" alt=""></a></td>
		<%      Else %>
		<td valign="top"  colspan="2" rowspan="2"><a href="login.asp" onmouseover="changeImages('clientlogin', 'images/clientlogin-over.jpg'); return true;" onmouseout="changeImages('clientlogin', 'images/clientlogin.jpg'); return true;" onclick="changeImages('clientlogin', 'images/clientlogin-over.jpg');"><img name="clientlogin" src="images/clientlogin.jpg" width="137" height="40" border="0" alt=""></a></td>
		<%      End If 
		    End If %>
		<td valign="top"  colspan="2" rowspan="2"><img src="images/generic_05.jpg" width="52" height="40" alt=""></td>
		<td valign="top"  style="height: 17px"><img src="images/spacer.gif" width="1" height="17" alt=""></td>
	</tr>
	<tr><!-- ROW 3 -->
		<%  If WEB_PAGE_ID = 2 Then     'Home Page %>
		<td valign="top"  rowspan="2"><a href="home.asp"><img src="images/index-over.jpg" width="89" height="49" border="0" alt=""></a></td>
		<%  Else %>
		<td valign="top"  rowspan="2"><a href="home.asp" onmouseover="changeImages('index', 'images/index-over.jpg'); return true;" onmouseout="changeImages('index', 'images/index-proofs_sel.jpg'); return true;" onclick="changeImages('index', 'images/index-over.jpg');"><img name="index" src="images/index-proofs_sel.jpg" width="89" height="49" border="0" alt=""></a></td>
	    <%  End If %>
		<td valign="top"  colspan="4"><img src="images/generic.jpg" width="367" height="23" alt=""></td>
		<td valign="top" ><img src="images/spacer.gif" width="1" height="23" alt=""></td>
	</tr>
	<tr><!-- ROW 4 -->
	    <%  If WEB_PAGE_ID = 7 Then     'Gallery Page %>
	    <td valign="top"  rowspan="2"><a href="gallery.asp"><img src="images/gallery-over.jpg" width="88" height="60" border="0" alt=""></a></td>
	    <%  Else %>
		<td valign="top"  rowspan="2"><a href="gallery.asp" onmouseover="changeImages('gallery', 'images/gallery-over.jpg'); return true;" onmouseout="changeImages('gallery', 'images/gallery.jpg'); return true;" onclick="changeImages('gallery', 'images/gallery-over.jpg');"><img name="gallery" src="images/gallery.jpg" width="88" height="60" border="0" alt=""></a></td>
		<%  End If %>
	    <%  If WEB_PAGE_ID = 3 Then     'About Page %>
	    <td valign="top"  rowspan="2"><a href="about.asp"><img src="images/about-over.jpg" width="117" height="60"  alt="" /></a></td>
		<%  Else %>
		<td valign="top"  rowspan="2"><a href="about.asp" onmouseover="changeImages('about', 'images/about-over.jpg'); return true;" onmouseout="changeImages('about', 'images/about.jpg'); return true;" onclick="changeImages('about', 'images/about-over.jpg');"><img name="about" src="images/about.jpg" width="117" height="60" border="0" alt=""></a></td>
		<%  End If %>
	    <%  If WEB_PAGE_ID = 4 Then     'Session Details Page %>
	    <td valign="top"  rowspan="2"><a href="details.asp"><img src="images/sessions-over.jpg" width="93" height="60" border="0" alt=""></a></td>
		<%  Else %>
        <td valign="top"  rowspan="2"><a href="details.asp" onmouseover="changeImages('sessions', 'images/sessions-over.jpg'); return true;" onmouseout="changeImages('sessions', 'images/sessions.jpg'); return true;" onclick="changeImages('sessions', 'images/sessions-over.jpg');"><img name="sessions" src="images/sessions.jpg" width="93" height="60" border="0" alt=""></a></td>		
		<%  End If %>
	    <%  If WEB_PAGE_ID = 5 Then     'Contact Page %>
	    <td valign="top"  colspan="2" rowspan="2"><a href="contact.asp"><img src="images/contact-over.jpg" width="110" height="60" border="0" alt=""></a></td>
		<%  Else %>
        <td valign="top"  colspan="2" rowspan="2"><a href="contact.asp" onmouseover="changeImages('contact', 'images/contact-over.jpg'); return true;" onmouseout="changeImages('contact', 'images/contact.jpg'); return true;" onclick="changeImages('contact', 'images/contact-over.jpg');"><img name="contact" src="images/contact.jpg" width="110" height="60" border="0" alt=""></a></td>        
		<%  End If %>		
		<td colspan="3" rowspan="2" valign="top"><img src="images/generic_07.jpg" width="148" height="60" alt=""></td>
		<td valign="top"><img src="images/spacer.gif" width="1" height="26" alt=""></td>
	</tr>
	<tr><!-- ROW 5 -->
		<td rowspan="2" valign="top"><img src="images/generic_06.jpg" width="89" height="67" alt=""></td>
		<td valign="top"><img src="images/spacer.gif" width="1" height="34" alt=""></td>
	</tr>
	<tr><!-- ROW 6 -->
		<td colspan="8"><img src="images/generic_08.jpg" width="556" height="33" alt=""></td>
		<td valign="top"><img src="images/spacer.gif" width="1" height="33" alt=""></td>
	</tr>
	<tr><!-- ROW 7 -->
		<td colspan="13" valign="top"><img src="images/generic_09.jpg" width="899" height="9" alt=""></td>
		<td valign="top"><img src="images/spacer.gif" width="1" height="9" alt=""></td>
	</tr>
	<tr><!-- ROW 8 -->
		<td valign="top"><img src="images/generic_10.jpg" width="23" height="359" alt=""></td>	