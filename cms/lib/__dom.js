//Get domain information
var domain = location.protocol+'//'+location.hostname;
var strQueryString=location.search;
var strQueryString=strQueryString.replace('?','');

//path information
var strRootPath='';
location.hostname=='localhost' ? strRootPath='/JulieStarkPhotography/' : strRootPath=domain;
var strCmsImagePath=strRootPath+'cms/images/';
var strProofRootPath=strRootPath+'secure/proofs/';
var strCmsCommonPath=strRootPath+'cms/common/';
var strPortfolioRootPath=strRootPath+'secure/portfolio/';

//page information
var strPhotoUploadForm=strCmsCommonPath+'__uploadForm.asp'+location.search;

//Add querystring information to an array in order to parse through
var qStr = new String(strQueryString);
var qArr = new Array();
var tempArr = qStr.split('&');
for (var i=0;i<tempArr.length;i++) {
	var strPair = new String(tempArr[i])
	var args = strPair.split('=');
	var key = args[0];var val = args[1];
	qArr.push(key);qArr.push(val);
}

//Get querystring values and set to global variables for later use
var siteId='';
var galleryId='';
var galleryName='';
var galleryLastName='';
var categoryId='';
var categoryName='';
var isDragDrop=true;
var lookupArr = new Array();
lookupArr = qArr.slice();
for (var i=0;i<qArr.length;i++) {
	switch(qArr[i]) { //get the querystring keyname and set the keyvalue by performing lookup on our copied array
		case 'gid':galleryId=lookupArr[i+1];break;
		case 'gn':galleryName=lookupArr[i+1];break;
		case 'gln':galleryLastName=lookupArr[i+1];break;
		case 'cid':categoryId=lookupArr[i+1];break;
		case 'cn':categoryName=lookupArr[i+1].replace(' ','');break;
	}
}

var photographer='julie';
var srcObj = new Object;	// the object that you are dragging:
var dummyObj;				// string to hold source of object being dragged:
var tempThumbImage='';		// string to hold the previously dropped thumbnail image
var tempLargeImage='';		// string to hold the previously dropped viewable image

var dom;
var domStyle;
var isDHTML = 0;
var isID = 0;
var isAll = 0;
var isLayers = 0;

if (document.getElementById) {
	isID = 1; 
	isDHTML = 1;
} else if (document.all) {
	isAll = 1; 
	isDHTML = 1;
} else {
	browserVersion = parseInt(navigator.appVersion);
	if ((navigator.appName.indexOf('Netscape') != -1) && (browserVersion == 4)) {
		isLayers = 1;
		isDHTML = 1;
	}
}

//object reference handler
function findDOM(objectID,withStyle) {
	if (withStyle == 1) {
		if (isID) { 
			return (document.getElementById(objectID).style);
		} else if (isAll) { 
			return (document.all[objectID].style); 
		} else if (isLayers) { 
			return (document.layers[objectID]); 
		}
	} else {
		if (isID) { 
			return (document.getElementById(objectID)); 
		} else if (isAll) {
			return (document.all[objectID]);
		} else if (isLayers) { 
			return (document.layers[objectID]);
		}
	}
}

//Hides/Shows the side navigation window
function toggle(oid) {
	if (isAll || isID) {
		domStyle = findDOM(oid,1);
		(domStyle.display=='block') ? domStyle.display='none' : domStyle.display='block';
	}
	return;
}

function toggleImageGallery(oid) {
	if (isAll || isID) {
		domStyle = findDOM(oid,1);
		((domStyle.visibility =='show') || (domStyle.visibility == 'visible')) ? domStyle.visibility = 'hidden' : domStyle.visibility='visible';
	}
	return;
}

function toggleMenuImage(oid,srcImg,trgImg) {
	if (isAll || isID) {
		dom = findDOM(oid);
		(dom.src==domain+strCmsImagePath+srcImg) ? dom.src=domain+strCmsImagePath+trgImg : dom.src=domain+strCmsImagePath+srcImg;
	}
}

function getCenteredWidth(w) {return (screen.width-w)/2;}
function getCenteredHeight(h) {return (screen.height-h)/2;}

function popup(url,win,h,w) {
	var winl=getCenteredWidth(w);
	var wint=getCenteredHeight(h)-50;
	var popupWin=window.open(url,win,'height='+h+',width='+w+',left='+winl+',top='+wint+',status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no');
}

function popupPhoto(url,win,features,w,h,isCenter) { 
	if(window.screen)if(isCenter)if(isCenter=='true'){
		var winl=getCenteredWidth(w);
		var wint=getCenteredHeight(h);
		features+=(features!='')?',':'';
		features+=',left='+winl+',top='+wint;
	}
	var popupWin=window.open(url,win,features+((features!='')?',':'')+'width='+w+',height='+h);
}

function connectPhotos() {
	var cnxnList=findDOM('selCnxn');
	var cnxn1=findDOM('LgName').innerText;
	var cnxn2=findDOM('ThbName').innerText;
	if ((cnxn1.length==1)||(cnxn2.length==1)) {
		alert('You must drag-n-drop a picture from both the viewable images and thumbnail images area before you are able to establish a connection between the two photographs.');
		return;
	} else if ((cnxn1.length>1)||(cnxn2.length>1)) {	
		var optTxt=cnxn1+'  <==>  '+cnxn2;
		for(var i=0;i<cnxnList.options.length;i++) { 
			if (cnxnList.options[i].text==optTxt) {
				return;
			}
		}		
		addOption(cnxnList,optTxt,optTxt)		
		resetDragDropElements();		
	} else {
		alert('You must drag-n-drop a picture from the viewable images and thumbnail images into the draggable area to establish a connection between the two photographs.');
		return;
	}
}

function resetConnections() {
	//remove all photograph relationships
	var cnxnList=findDOM('selCnxn');
	for (var i=cnxnList.options.length-1;i>=0;i--) {
		cnxnList.options[i] = null;
	}

	//clear out the viewing area for all connected photos in the workspace
	var cnxnPhotoList = findDOM('connectedPhotoList');
	cnxnPhotoList.innerHTML='';

	//make all previously shown photographs that were connected visible again
	var spanTags = document.all.tags('SPAN');
	for (var i=0;i<spanTags.length;i++) {
		if (spanTags(i).id.indexOf('container_') != -1) {
			spanTags(i).style.display='block';
		}
	}	

	resetDragDropElements();
}

function resetDragDropElements() {

	//reset the drag-n-drop backdrop and other elements to the default
	var backDrop=findDOM('BackDrop');
	backDrop.style.backgroundImage ='url('+strCmsImagePath+'ddbackdrop.gif)';

	var tprev=findDOM('thumbPreview');
	var tname=findDOM('ThbName');
	var lprev=findDOM('largePreview');
	var lname=findDOM('LgName');

	tprev.src=strCmsImagePath+'ThumbDrop.gif';
	lprev.src=strCmsImagePath+'LargeDrop.gif';
	lname.innerText=tname.innerText='';
	
	tempLargeImage='';
	tempThumbImage='';
}

function removeConnection(cnxn) {

	var cnxnList=findDOM('selCnxn');
	for (var i=cnxnList.options.length-1;i>=0;i--) {
		if (cnxnList.options[i].text == cnxn) {
			cnxnList.options[i] = null;
		}
	}

	//strip out both the viewable and thumbnail image;
	var strCnxn=new String(cnxn);
	var arrCnxn=strCnxn.split(' <==> ');
	var lrgImg=arrCnxn[0].replace('.jpg','').replace(' ','');
	var thumbImg=arrCnxn[1].replace('.jpg','').replace(' ','');	

	//make all previously shown photographs that were connected visible again
	var spanTags=document.all.tags('SPAN');
	for (var i=0;i<spanTags.length;i++) {
		if (spanTags(i).id.indexOf('container_')!=-1) {
			if (spanTags(i).id=='container_largeimg_'+lrgImg) {
				spanTags(i).style.display='block';
			} else if (spanTags(i).id=='container_thumb_'+thumbImg) {
				spanTags(i).style.display='block';
			}
		}
	}
	
	resetDragDropElements();
	viewConnectedPhotos();
}

function viewConnectedPhotos() {

	//grab a reference to the list of connected photographs
	var cnxnList=findDOM('selCnxn');
	//var list='';
	var listHtml='';
	
	//build the list of all connected photographs
	for(var i=0;i<cnxnList.options.length;i++) { 

		//strip out both the viewable and thumbnail image;
		var strCnxn=new String(cnxnList.options[i].text);
		var arrCnxn=strCnxn.split(' <==> ');
		var thumbImg=arrCnxn[1].replace(' ','');	

		listHtml+='<a href=\"#\" '
		listHtml+='onclick=\"removeConnection(\''+cnxnList.options[i].text+'\'); \"> '
		
		//listHtml+='<img align=\"middle\" style=\"cursor:pointer;cursor:hand;border:1 solid #666;\" vspace=\"1\" width=\"50\" height=\"50\" src=\"'+strProofRootPath+photographer+'/'galleryLastName+'/thumbs/'+thumbImg+'\"></a>&nbsp;'

		if (categoryId>0) {
			listHtml+='<img align=\"middle\" style=\"cursor:pointer;cursor:hand;border:1 solid #666;\" vspace=\"1\" width=\"50\" height=\"50\" src=\"'+strPortfolioRootPath+photographer+'/'+categoryName+'/thumbs/'+thumbImg+'\"></a>&nbsp;'			
		} else if(galleryId>0) {
			listHtml+='<img align=\"middle\" style=\"cursor:pointer;cursor:hand;border:1 solid #666;\" vspace=\"1\" width=\"50\" height=\"50\" src=\"'+strProofRootPath+photographer+'/'+galleryLastName+'/thumbs/'+thumbImg+'\"></a>&nbsp;'
		}
		listHtml+=cnxnList.options[i].text+'<br/>';
	}
	
	//grab a reference to the display list of connected photographs
	var cnxnPhotoList = findDOM('connectedPhotoList');
	
	//display the list of connected photographs	
	cnxnPhotoList.innerHTML=listHtml;
}

function checkForUnsubmittedConnections() {

	//grab reference to list of connected photographs
	var cnxnList=findDOM('selCnxn');
	
	//if there are connected photographs, we prevent user from uploading additional photographs (this would cause entire page to refresh prior to saving to the database)
	if (cnxnList.options.length>0) {
		alert('You must submit all your viewable/thumbnail photograph relationships prior to uploading new photos to your workspace.');
		return;
	} else {

		//if there are no photos connected, then open the photo upload center
		popup(strPhotoUploadForm,'Upload',500,500);
		return true;
	}
}

//strip the filename based on the path to the file
function stripFileName(s)  {
	lastSlash=s.lastIndexOf('/',s.length-1)
	return s.substring (lastSlash+1,s.length);
}


function startDrag(){  
    srcObj=window.event.srcElement;
    dummyObj=srcObj.outerHTML;
    with(window.event.dataTransfer) {
		setData('Text', window.event.srcElement.src);
    	effectAllowed='linkMove';
    	dropEffect='move';
	}
}

function enterDrag(){window.event.dataTransfer.getData('Text');}
function endDrag(){window.event.dataTransfer.clearData();}
function overDrag(){window.event.returnValue = false;}

function setPhotoSwapCache(rad,imgname) {

    if (isDragDrop){
        
	    //if dropping the thumbnail
	    if (srcObj.id.indexOf('thumb_') != -1) {
		    //if a thumbnail has already been dropped, then add it back to general population (tempThumbImage is global variable initialized to empty string)
		    if (tempThumbImage.length>0) {	
    			
    			//(stripFileName(tempThumbImage).replace('.jpg',''))
			    //get the name of the container which the thumbnail is embedded
			    var tempContainer='container_thumb_'+lcase(imgname);

			    //grab object reference to the thumbnail's container
			    var tempPhoto=findDOM(tempContainer,1);
    			
			    //reset the previously dropped photo by re-adding it to general population
			    tempPhoto.display='block';			
		    }
		    //update our temporary thumbnail image holder with the currently dropped image
		    tempThumbImage=srcObj.src;
    		
	    } else {
		    //if a large image has already been dropped, then add it back to general population (tempLargeImage is global variable initialized to empty string)	
		    if (tempLargeImage.length>0) {

                //(stripFileName(tempLargeImage).replace('.jpg',''))
			    //get the name of the container which the viewable image is embedded
			    var tempContainer='container_largeimg_'+lcase(imgname);

			    //grab object reference to the viewable image's container
			    var tempPhoto=findDOM(tempContainer,1);

			    //reset the previously dropped photo by re-adding it to general population
			    tempPhoto.display='block';			
		    }	
		    //update our temporary viewable image holder with the currently dropped image		
		    tempLargeImage=srcObj.src;
	    }
    } else {
        
    
    /*
	    //if connecting the thumbnail
	    if (rad.value.indexOf('thumb_') > -1) {   
		    //if a thumbnail has already been dropped, then add it back to general population (tempThumbImage is global variable initialized to empty string)
		    if (tempThumbImage.length>0) {	
    			
			    //get the name of the container which the thumbnail is embedded
			    var tempContainer='container_thumb_'+(stripFileName(tempThumbImage).replace('.jpg',''));

			    //grab object reference to the thumbnail's container
			    var tempPhoto=findDOM(tempContainer,1);
    			
			    //reset the previously dropped photo by re-adding it to general population
			    tempPhoto.display='block';			
		    }
		    
		    //update our temporary thumbnail image holder with the currently dropped image
            var srcImg=findDOM('thumb_'+imgname);
		    tempThumbImage=srcImg.src;
    		
	    } else {
		    //if a large image has already been dropped, then add it back to general population (tempLargeImage is global variable initialized to empty string)	
		    if (tempLargeImage.length>0) {

			    //get the name of the container which the viewable image is embedded
			    var tempContainer='container_largeimg_'+(stripFileName(tempLargeImage).replace('.jpg','')));

			    //grab object reference to the viewable image's container
			    var tempPhoto=findDOM(tempContainer,1);

			    //reset the previously dropped photo by re-adding it to general population
			    tempPhoto.display='block';			
		    }	
		    //update our temporary viewable image holder with the currently dropped image		
		    var srcImg=findDOM('largeimg_'+imgname);
		    tempLargeImage=srcImg.src;
	    }    */
    }
}

function radClick(rad,imgname) {

	isDragDrop=false;
	window.event.returnValue = false;
	var backDrop=findDOM('BackDrop',1);
	backDrop.backgroundImage ='url(/cms/images/ConnectedBackDrop.gif)';
	
	//if connecting the thumbnail
	if (rad.value.indexOf('thumb_') > -1) {   
		setPhotoSwapCache(rad,imgname);
		
		//grab object references to the drag-n-drop box elements
		var tprev=findDOM('thumbPreview');
		var tname=findDOM('ThbName');
        var srcImg=findDOM('thumb_'+imgname);

		//set the source of the thumbnail previewer to the new photograph
		tprev.src=srcImg.src;
		
		//set the text to be the image name
		tname.innerText=imgname;	    

	    //grab object reference to the container object which the image is embedded and remove it from the general population
	    var photo=findDOM('container_thumb_'+imgname,1);

	} else {
		setPhotoSwapCache(rad,imgname);
		
		//grab object references to the drag-n-drop box elements		
		var lprev=findDOM('largePreview');
		var lname=findDOM('LgName');
		var srcImg=findDOM('largeimg_'+imgname);

		//set the source of the viewable image previewer to the new photograph		
		lprev.src=srcImg.src;
		
		//set the text to be the image name		
		lname.innerText=imgname;

	    //grab object reference to the container object which the image is embedded and remove it from the general population
	    var photo=findDOM('container_largeimg_'+imgname,1);
	}
	photo.display='none';   
}


function drop() {
    if (isDragDrop){
	    window.event.returnValue = false;
	    var backDrop=findDOM('BackDrop',1);
	    backDrop.backgroundImage ='url(/cms/images/ConnectedBackDrop.gif)';

	    //if dropping the thumbnail
	    if (srcObj.id.indexOf('thumb_') != -1) {

		    setPhotoSwapCache(null,null);
    		
		    //grab object references to the drag-n-drop box elements
		    var tprev=findDOM('thumbPreview');
		    var tname=findDOM('ThbName');

		    //set the source of the thumbnail previewer to the new photograph
		    tprev.src=srcObj.src;
    		
		    //set the text to be the image name
		    tname.innerText=srcObj.fileName;
	    } else {

		    setPhotoSwapCache(null,null);
    		
		    //grab object references to the drag-n-drop box elements		
		    var lprev=findDOM('largePreview');
		    var lname=findDOM('LgName');

		    //set the source of the viewable image previewer to the new photograph		
		    lprev.src=srcObj.src;
    		
		    //set the text to be the image name		
		    lname.innerText=srcObj.fileName;
	    }
	    //grab object reference to the container object which the image is embedded and remove it from the general population
	    var photo=findDOM('container_'+srcObj.id,1);
	    photo.display='none';
    }
}

//selects all items in listbox to be submitted
function selectList(trgObj) {
	for(var i=0;i<trgObj.options.length;i++) { 
		if (trgObj.options[i]!=null) {
			trgObj.options[i].selected=true;
		}
	}
	return true;
}

//adds item to listbox
function addOption(trgObj,optTxt,optVal) {
    var newOpt = new Option(optTxt,optVal)
    var len= trgObj.options.length
    trgObj.options[len]=newOpt
}

//deletes item from listbox
function deleteOption(trgObj,el) {
	if (trgObj.options.length!=0) {trgObj.options[el]=null;}
}

//submits the photos to the database and completes the photo connection process
function submitForm() {
	var cnxnList = findDOM('selCnxn');
	if (cnxnList.options.length==0) {
		alert('You are trying to submit your connected photographs without having any connected. Please connect the photographs, then try submitting again.');
		return;
	}
	var connect = findDOM('hidConnect');
	connect.value='true';	
	selectList(cnxnList);
	document.forms[0].submit();
}

//toggles checkboxes to be checked/unchecked
function ToggleAll(e) {
	if (e.innerHTML=='Activate'){
		CheckAll();
		e.innerHTML='De-Activate';
	} else {
		ClearAll();
		e.innerHTML='Activate';
	}
}

//checks all the checkboxes
function CheckAll() {
	var frm = document.forms[0];
	var len = frm.elements.length;
	for (var i = 0; i < len; i++) {
		var e = frm.elements[i];
		if(e.type=='checkbox') e.checked=true;
	}
}

//un-checks all the checkboxes
function ClearAll() {
	var frm = document.forms[0];
	var len = frm.elements.length;
	for (var i = 0; i < len; i++) {
		var e = frm.elements[i];
		if(e.type=='checkbox') e.checked=false;
	}
}
