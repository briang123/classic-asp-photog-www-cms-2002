function newImage(arg) {
	if (document.images) {
		rslt = new Image();
		rslt.src = arg;
		return rslt;
	}
}

function changeImagesArray(array) {
	if (document.images && (preloadFlag == true)) {
		for (var i=0; i<array.length; i+=2) {
			document[array[i]].src = array[i+1];
		}
	}
}

function changeImages() {
	changeImagesArray(changeImages.arguments);
}

function toggleImages() {
	for (var i=0; i<toggleImages.arguments.length; i+=2) {
		if (selected == toggleImages.arguments[i]) {
		    changeImagesArray(toggleImages.arguments[i+1]);
		}
	}
}

var selected = '';
var preloadFlag = false;
function preloadImages() {
	if (document.images) {
		clientlogin_over = newImage("images/clientlogin-over.jpg");
		index_over = newImage("images/index-over.jpg");
		gallery_over = newImage("images/gallery-over.jpg");
		about_over = newImage("images/about-over.jpg");
		sessions_over = newImage("images/sessions-over.jpg");
		contact_over = newImage("images/contact-over.jpg");
		preloadFlag = true;
	}
}
