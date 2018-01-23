
function openhero(hid) {
 
 var openhero = window.open(hid,'','toolbars=0,scrollbars=1,location=0,statusbars=0,menubars=0,resizable=1,width=450,height=475');
	
} 

function showimage(imageid) {
 
 var openhero = window.open('~/ShowImage.aspx?imageid=' + imageid,'','toolbars=0,scrollbars=0,location=0,statusbars=0,menubars=0,resizable=0,width=450,height=300');
	
} 


function openemailwindow(hid) {
 
 var openhero = window.open(hid,'','toolbars=0,scrollbars=1,location=0,statusbars=0,menubars=0,resizable=0,width=440,height=475');
} 

function openWindow(hid) {
 
 var openWindow = window.open('modelnumberlocate.aspx?i=' + hid,'','toolbars=0,scrollbars=1,location=0,statusbars=0,menubars=0,resizable=0,width=440,height=475');
} 

function resizewindow(pagename,width, height)
{
	if (window.formcorey.hdnLoaded.value == "0")
	{
	var resizewindow = window.open(window.document.URL + "&loaded=1",'','toolbars=0,scrollbars=1,location=0,statusbars=0,menubars=0,resizable=0,width=400,height=400');
	var caller = resizewindow.opener
	resizewindow.resizeTo(width,height);
	caller.close();
	return false;
	}
	
}


function clearValue(item) {

item.value = '';	
}

function popup()
 {
 	window.open("http://www.netdeals.com/index.cfm?main=privacy","window","resizable=yes,menubar=yes,directories=no,status=yes,location=no,scrollbars=yes,width=650,height=400");
 }

 function confirmGRDelete()
 {
 var doIt = window.confirm("Are you sure you want to delete this Gift Registry?. Click Cancel to stop.");
if (doIt)
   return true;
else 
   return false;
}

function getAddresses() {
		var winstyle = "innerWidth=450,innerHeight=380,height=380,width=450,toolbar=no,menubar=no,location=no,scrollbars=yes";
		winstyle += ",left=" + (screen.width - 450)/2;
		winstyle += ",top=" + (screen.height - 380)/2;
		aWin = window.open( "../wishaddressbook.asp?strWID=<%=vstrWID%>", null, winstyle );
		if( aWin.opener == null )
			aWin.opener = self;
		aWin.focus();
	}

	function addAddresses( strAddresses ) {
		var toInput = document.forms.frmEmail.txtFEmail;
		toInput.value = toInput.value + ((toInput.value.length > 0) ? "," : "") + strAddresses;
	}

	function doUnload() {
	}
	
function addAddresses() {
	if(opener.addAddresses ){
		var eForm = document.forms.emailForm;
		var tempStr = "";
		for( var i = 0; i < eForm.elements.length; i++ ){
			if( eForm.elements[i].checked )
				tempStr += ((tempStr.length >  0) ? "," : "") + eForm.elements[i].value
		}
		if( tempStr.length > 0 )
			eForm.strAddresses.value = tempStr;
			opener.addAddresses( tempStr );
}
		window.close();
}
function deleteAddresses() {
	var eForm = document.forms.emailForm;
	eForm.submit();
	window.close();
}
function changeBox(cbox) {
	box = eval(cbox);
	box.checked = !box.checked;
}

function ShowAddressBook(container_id)
{
var sFeatures = "innerWidth=450,innerHeight=380,height=380,width=450,toolbar=no,menubar=no,location=no,scrollbars=yes"
		sFeatures += ",left=" + (screen.width - 450)/2;
		sFeatures += ",top=" + (screen.height - 380)/2;
	var opener = window.open ("GRAddressBook.aspx?cid=" + container_id,null,sFeatures)
	opener.focus();

}
	

function UpdateMasterEmailForm()
{
	var opener = window.opener;
	if (emailSelect.hdnSelectedValue.value != "")
	{
		try
		{
			var control = eval("opener.document.forms[0].txtFEmail")
			control.value = emailSelect.hdnSelectedValue.value 
			window.close();
		}
		catch(e)
		{
			alert("invalid selection");
		
		}
	}
}
