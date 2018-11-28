<%
'get the individual photo (must have at least one)
DIM SIDEPHOTO
SIDEPHOTO = "1x1.gif"
Set oPhoto = New cPhotos
Set collPhotos = New cPhotos
Call collPhotos.GetSitePhotos()
For Each oPhoto In collPhotos.Photos.Items 
	If oPhoto.WebPageId = WEB_PAGE_ID And oPhoto.ActiveFlag = True Then
		SIDEPHOTO = oPhoto.LargeImage
        Set oPhoto = Nothing
        Exit For
	End If
Next
Set collPhotos = Nothing
%>
<td class="body-photo-display" valign="top" width="264">
	<table cellpadding="0" cellspacing="0" border="0" class="body-photo-display-liner" width="264">
		<tr>
			<td style="padding:5;" align="center" valign="middle">
				<% 
				Dim imgArrCounter
				If RANDOMIZE_SIDE_PHOTO Then %>
				<script type="text/javascript" language="javascript"><!--
				function sidePhotoRndImg() {
					var sidephoto= new Array()
				<%
				Set oPhoto = New cPhotos
				Set collPhotos = New cPhotos
				Call collPhotos.GetSitePhotos()
				imgArrCounter = 0
				For Each oPhoto In collPhotos.Photos.Items 
					If oPhoto.WebPageId = WEB_PAGE_ID And oPhoto.ActiveFlag = True Then
				%>
					sidephoto[<%=imgArrCounter%>]='<%=oPhoto.LargeImage%>';
				<%	imgArrCounter = imgArrCounter + 1
					End If
						Set oPhoto = Nothing
				Next
				Set collPhotos = Nothing
				%>
					var img = Math.floor(Math.random() * sidephoto.length);
					document.write('<img src=<%=IMAGE_PATH%>/'+sidephoto[img]+' width=264 height=326 alt=<%=ALT_IMAGE_TEXT%>>');
				}
				sidePhotoRndImg();
				//--></script>			
				<noscript><img src="<%=IMAGE_PATH%>/<%=SIDEPHOTO%>" width="264" height="326" alt="<%=ALT_IMAGE_TEXT%>"></noscript>
				<% ElseIf RUN_SIDE_PHOTO_SLIDE_SHOW Then %>

				<script type="text/javascript" language="javascript"><!--
					var slidespeed = <%=SLIDE_SHOW_SLIDE_SPEED%>;
					var slideimages = new Array();
				<%
				Set oPhoto = New cPhotos
				Set collPhotos = New cPhotos
				Call collPhotos.GetSitePhotos()
				imgArrCounter = 0
				For Each oPhoto In collPhotos.Photos.Items 
					If oPhoto.WebPageId = WEB_PAGE_ID And oPhoto.ActiveFlag = True Then	%>
					slideimages[<%=imgArrCounter%>]='<%=IMAGE_PATH & "/" & oPhoto.LargeImage%>';
				<%	SIDEPHOTO = oPhoto.LargeImage
				    imgArrCounter = imgArrCounter + 1
					End If
					Set oPhoto = Nothing
				Next
				Set collPhotos = Nothing
				%>
					var imageholder = new Array();
					var ie = document.all;
					for (i=0;i<slideimages.length;i++){
						imageholder[i] = new Image();
						imageholder[i].src = slideimages[i];
					}
				//--></script>
				<img src="<%=IMAGE_PATH%>/<%=SIDEPHOTO%>" name="slide" alt="<%=ALT_IMAGE_TEXT%>" width="264" height="326" style="filter:progid:DXImageTransform.Microsoft.Wipe(duration=<%=SLIDE_SHOW_DURATION%>,gradientsize=1.0, motion='<%=SLIDE_SHOW_MOTION%>', wipestyle=<%=SLIDE_SHOW_WIPE_STYLE%>)">				
				<script type="text/javascript" language="javascript"><!--
				var whichimage=0;
				var blenddelay=(ie) ? document.images.slide.filters[0].duration*1000 : 0;
				function slideit() {
					if (!document.images) return
					if (ie) document.images.slide.filters[0].apply();
					document.images.slide.src = imageholder[whichimage].src;
					if (ie) document.images.slide.filters[0].play();
					whichimage =(whichimage<slideimages.length-1) ? whichimage+1 : 0;
					setTimeout("slideit()",slidespeed+500);
				}
				slideit();
				//--></script>
				<% Else %>
				<img src="<%=IMAGE_PATH%>/<%=SIDEPHOTO%>" width="264" height="326" alt="<%=ALT_IMAGE_TEXT%>">
				<% End If %>
			</td>
		</tr>
	</table>
</td>