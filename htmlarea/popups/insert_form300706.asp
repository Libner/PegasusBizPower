<!--#include file="../../NETcom/connect.asp"-->
<%pageId = Request.QueryString("pageid")
If isNumeric(pageID) = false Then
	%>
	<script language=javascript>
	<!--
			window.alert("לפני הוספת טופס , שמור את הדף המעוצב");
			self.close();
	//-->
	</script>
	<%
Else

	set rs_page = con.GetRecordSet("Select Product_Id,FORM_LINK_IMAGE,LINK_IMAGE,LINK_IMAGE_ALIGN,LINK_TEXT,LINK_TEXT_ALIGN,LINK_BGCOLOR,LINK_FONT_TYPE,LINK_FONT_SIZE,LINK_FONT_COLOR,FORM_BGCOLOR,FORM_FONT_TYPE,FORM_FONT_SIZE,FORM_FONT_COLOR from pages where Page_Id=" & pageId)
	if not rs_page.eof then
		Product_Id = rs_page("Product_Id")
		FORM_LINK_IMAGE = Trim(rs_page("FORM_LINK_IMAGE"))
		ImgActualSize = rs_page.Fields("LINK_IMAGE").ActualSize
		LINK_IMAGE_ALIGN = Trim(rs_page("LINK_IMAGE_ALIGN"))
		LINK_TEXT = Trim(rs_page("LINK_TEXT"))
		LINK_TEXT_ALIGN = Trim(rs_page("LINK_TEXT_ALIGN"))
		LINK_BGCOLOR = Trim(rs_page("LINK_BGCOLOR"))
		LINK_FONT_TYPE = Trim(rs_page("LINK_FONT_TYPE"))
		LINK_FONT_SIZE = Trim(rs_page("LINK_FONT_SIZE"))
		LINK_FONT_COLOR = Trim(rs_page("LINK_FONT_COLOR"))
		FORM_BGCOLOR = Trim(rs_page("FORM_BGCOLOR"))
		FORM_FONT_TYPE = Trim(rs_page("FORM_FONT_TYPE"))
		FORM_FONT_SIZE = Trim(rs_page("FORM_FONT_SIZE"))
		FORM_FONT_COLOR = Trim(rs_page("FORM_FONT_COLOR"))
	end if	
	set rs_page = nothing
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD W3 HTML 3.2//EN">
<HTML  id=dlgBullet STYLE="width: 374px; height: 325px; ">
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
<meta http-equiv="MSThemeCompatible" content="Yes">
<TITLE>Insert Form</TITLE>
<style>
  html, body, button, div, input, select, fieldset { font-family: MS Shell Dlg; font-size: 8pt; position: absolute; };
</style>
<SCRIPT defer>
var ImageTofes = false ;
var Product_Id = '<%=Product_Id%>';
var FORM_LINK_IMAGE = '<%=FORM_LINK_IMAGE%>';
var ImgActualSize = '<%=ImgActualSize%>';
var LINK_IMAGE_ALIGN = '<%=LINK_IMAGE_ALIGN%>';
var LINK_TEXT = '<%=LINK_TEXT%>';
var	LINK_TEXT_ALIGN = '<%=LINK_TEXT_ALIGN%>';
var	LINK_BGCOLOR = '<%=LINK_BGCOLOR%>';
var	LINK_FONT_TYPE = '<%=LINK_FONT_TYPE%>';
var	LINK_FONT_SIZE = '<%=LINK_FONT_SIZE%>';
var	LINK_FONT_COLOR = '<%=LINK_FONT_COLOR%>';
var	FORM_BGCOLOR = '<%=FORM_BGCOLOR%>';
var	FORM_FONT_TYPE = '<%=FORM_FONT_TYPE%>';
var	FORM_FONT_SIZE = '<%=FORM_FONT_SIZE%>';
var	FORM_FONT_COLOR = '<%=FORM_FONT_COLOR%>';

function _CloseOnEsc() {
  if (event.keyCode == 27) { window.close(); return; }
}

function _getTextRange(elm) {
  var r = elm.parentTextEdit.createTextRange();
  r.moveToElementText(elm);
  return r;
}

window.onerror = HandleError

function HandleError(message, url, line) {
  var str = "An error has occurred in this dialog." + "\n\n"
  + "Error: " + line + "\n" + message;
  alert(str);
  window.close();
  return true;
}
function Init() {
  var elmSelectedImage;
  var htmlSelectionControl = "Control";
  var globalDoc = window.dialogArguments;
  var grngMaster = globalDoc.selection.createRange();
  
  if (globalDoc.selection.type == htmlSelectionControl) {
    if (grngMaster.length == 1) {
      elmSelectedImage = grngMaster.item(0);
      if (elmSelectedImage.tagName == "IMG" && elmSelectedImage.id == "tofes") {
        ImageTofes = true;
        //window.parent.partitle.innerHTML = 'Update Form';
        document.form1.selFormId.value = Product_Id;
        if (FORM_LINK_IMAGE == 'link')
        {	document.all.cbFormLink(0).checked = true;
			document.form1.bgrColor.value = LINK_BGCOLOR ;
			document.all.imgbgrColor.style.backgroundColor = LINK_BGCOLOR ;
			document.form1.textColor.value = LINK_FONT_COLOR ;
			document.all.imgtxtColor.style.backgroundColor = LINK_FONT_COLOR ;
			document.all.fsize.value = LINK_FONT_SIZE ;
			document.all.fstyle.value = LINK_FONT_TYPE ;
			document.all.LinkText.value = LINK_TEXT ;
			document.all.fTxtAlign.value = LINK_TEXT_ALIGN ;
        }else if (FORM_LINK_IMAGE == 'form'){
			document.all.cbFormLink(2).checked = true;
			document.form1.bgrColor.value = FORM_BGCOLOR ;
			document.all.imgbgrColor.style.backgroundColor = FORM_BGCOLOR ;
			document.form1.textColor.value = FORM_FONT_COLOR ;
			document.all.imgtxtColor.style.backgroundColor = FORM_FONT_COLOR ;
			document.all.fsize.value = FORM_FONT_SIZE ;
			document.all.fstyle.value = FORM_FONT_TYPE ;
        }else if (FORM_LINK_IMAGE == 'image'){ // image link
			document.all.cbFormLink(1).checked = true;
			document.all.fImgAlign.value = LINK_IMAGE_ALIGN ;
        }else{
			document.all.cbFormLink(0).checked = true;
        }
        setlink();
      }
    }
  }
  
  if(globalDoc.images("tofes") && !ImageTofes)
  {alert(".אין אפשרות להוספת שני טפסים לאותו דף פרסומי \n\n                                       .אנא מחק את הטופס הקיים");
  window.close();}
  // event handlers  
  document.body.onkeypress = _CloseOnEsc;
  btnOK.onclick = new Function("getHtmlForm()");
  selClick();
}

function getHtmlForm(){
	if (document.all.cbFormLink(1).checked){ //image link
		if(!document.form2.UploadFile1.value && !ImageTofes)
		{ 
		  alert("Image URL must be specified.");
		  document.form2.UploadFile1.focus();
		  return;
		}
		document.form2.action = 'Formadd.asp?pageid=' + hpageid.value + '&formId=' + document.form1.selFormId.value
		document.form2.submit();	
	}else{
		document.form1.action = 'Formadd.asp?pageid=' + hpageid.value
		document.form1.submit();	
	}
}

function btnOKClick() {
  var htmlSelectionControl = "Control";
  var globalDoc = window.dialogArguments;
  var grngMaster = globalDoc.selection.createRange();
  
  // delete selected content and replace with image
  if (globalDoc.selection.type == htmlSelectionControl || grngMaster.htmlText) {
	if(!ImageTofes)
	{
		if (!confirm("Overwrite selected content?")) 
		{ return; }
	}
    grngMaster.execCommand('Delete');
    grngMaster = globalDoc.selection.createRange();
  }
    
  grngMaster.pasteHTML(txtForm.value)
  window.close();
}
function selClick(){
	if(document.form1.selFormId.value)
	{ btnOK.disabled = false;}
	else
	{ btnOK.disabled = true;}
}
function setBgrColor()
{
	var oldcolor = document.all.bgrColor.value;
    var newcolor = showModalDialog("select_color.html", oldcolor, "resizable: no; help: no; status: no; scroll: no;");
    if (newcolor != null) { document.all.bgrColor.value = "#"+newcolor; document.all.imgbgrColor.style.backgroundColor = "#"+newcolor;}
}
function setTextColor()
{
	var oldcolor = document.all.textColor.value;
    var newcolor = showModalDialog("select_color.html", oldcolor, "resizable: no; help: no; status: no; scroll: no;");
    if (newcolor != null) { document.all.textColor.value = "#"+newcolor; document.all.imgtxtColor.style.backgroundColor = "#"+newcolor; }
}
function setlink(){
	if (document.all.cbFormLink[0].checked )
	{
		lgdText.innerText = ' Text Formating '
		document.all.LinkText.style.display = 'block';
		document.all.fTxtAlign.style.display = 'block';
		document.all.divLinkText.style.display = 'block';
		document.all.divtxtalign.style.display = 'block';
		document.all.divbgrColor.style.display = 'block';
		document.all.bgrColor.style.display = 'block';
		document.all.button4.style.display = 'block';
		document.all.divtxtColor.style.display = 'block';
		document.all.textColor.style.display = 'block';
		document.all.button3.style.display = 'block';
		document.all.divfsize.style.display = 'block';
		document.all.fsize.style.display = 'block';
		document.all.divstyle.style.display = 'block';
		document.all.fstyle.style.display = 'block';
		document.all.divUploadImg.style.display = 'none';
		document.all.UploadFile1.style.display = 'none';
		document.all.divimgalign.style.display = 'none';
		document.all.fImgAlign.style.display = 'none';		
	}
	else if (document.all.cbFormLink[1].checked)
	{
		lgdText.innerText = ' Image Layout '
		document.all.LinkText.style.display = 'none';
		document.all.fTxtAlign.style.display = 'none';
		document.all.divLinkText.style.display = 'none';
		document.all.divtxtalign.style.display = 'none';
		document.all.divbgrColor.style.display = 'none';
		document.all.bgrColor.style.display = 'none';
		document.all.button4.style.display = 'none';
		document.all.divtxtColor.style.display = 'none';
		document.all.textColor.style.display = 'none';
		document.all.button3.style.display = 'none';
		document.all.divfsize.style.display = 'none';
		document.all.fsize.style.display = 'none';
		document.all.divstyle.style.display = 'none';
		document.all.fstyle.style.display = 'none';
		document.all.divUploadImg.style.display = 'block';
		document.all.UploadFile1.style.display = 'block';
		document.all.divimgalign.style.display = 'block';
		document.all.fImgAlign.style.display = 'block';
	}
	else if (document.all.cbFormLink[2].checked)
	{
		window.alert("שימו לב : ספקי המייל האינטרנטיים \n\n<%=Space(17)%>('וכו Walla, Hotmail : כגון)\n\n<%=Space(14)%>.אינם תומכים באפשרות זו");
		lgdText.innerText = ' Text Formating '
		document.all.LinkText.style.display = 'none';
		document.all.fTxtAlign.style.display = 'none';
		document.all.divLinkText.style.display = 'none';
		document.all.divtxtalign.style.display = 'none';
		document.all.divbgrColor.style.display = 'block';
		document.all.bgrColor.style.display = 'block';
		document.all.button4.style.display = 'block';
		document.all.divtxtColor.style.display = 'block';
		document.all.textColor.style.display = 'block';
		document.all.button3.style.display = 'block';
		document.all.divfsize.style.display = 'block';
		document.all.fsize.style.display = 'block';
		document.all.divstyle.style.display = 'block';
		document.all.fstyle.style.display = 'block';
		document.all.divUploadImg.style.display = 'none';
		document.all.UploadFile1.style.display = 'none';
		document.all.divimgalign.style.display = 'none';
		document.all.fImgAlign.style.display = 'none';
	}	
}
</SCRIPT>
</HEAD>
<BODY id=bdy onload="Init()" style="background: threedface; color: windowtext;" scroll=no>
<INPUT ID=txtForm name=txtForm type=hidden value="">
<INPUT ID=hpageid name=hpageid type=hidden value="<%=pageId%>">
<form name="form1" ACTION="Formadd.asp" target="hidden_frame" METHOD="post">
<FIELDSET id=fldSelect style="padding-left:10px; left: 12px; top: 10px; width: 250px; height: 120px;">
<LEGEND id=lgdSelect> Select Form </LEGEND>
	<SELECT dir=rtl size=7 ID=selFormId name=selFormId tabIndex=20 style="width: 226px;left: 10px; top: 15px;" onclick="selClick();">	
	<%
		set rs_products = con.GetRecordSet("Select product_id, product_name FROM products WHERE product_number = '0' AND FORM_MAIL = '1' AND ORGANIZATION_ID=" & trim(Request.Cookies("biznetcom")("OrgID")) & " ORDER BY PRODUCT_ID")
		do while not rs_products.eof
		prod_Id = rs_products("product_id")
		product_name = rs_products("product_name")%>
		<OPTION value="<%=prod_Id%>" > <%=product_name%> </OPTION>
	<%	rs_products.MoveNext
		loop
	set rs_products=nothing%>	
	</SELECT>
</FIELDSET>
<DIV id=divFormLinkImage style="left: 12px; top: 142px;"> Show form :</DIV>
<INPUT type="radio" name=cbFormLink style="left: 80px; top: 140px;" value="link" onclick="setlink();" checked>
<DIV id=divLink style="left: 100px; top: 142px;">Text link</DIV>
<INPUT type="radio" name=cbFormLink style="left: 155px; top: 140px;" value="image" onclick="setlink();">
<DIV id=divLink style="left: 175px; top: 142px;">Image link</DIV>
<INPUT type="radio" name=cbFormLink style="left: 240px; top: 140px;" value="form" onclick="setlink();">
<DIV id="Div1" style="left: 260px; top: 142px;">Inside page</DIV>

<FIELDSET id=fldText style="left: 12px; top: 165px; width: 345px; height: 122px;">
<LEGEND id=lgdText> Text Formating </LEGEND>
<DIV id=divbgrColor style="left: 8px; top: 20px;"> Background color:</DIV>
<input type="text" name="bgrColor" value="" size=7 style="height: 22px; left: 100px; top: 15px;" maxlength=7>
<BUTTON style="top: 15px;left:160px ;width: 22px; height: 22px; border: 1px none;"  unselectable="on" onclick="setBgrColor();" id=button4 name=button4><img id=imgbgrColor style="background-color:#000000;" src="../images/selcolor.gif" border=0></BUTTON>

<DIV id=divtxtColor style="left: 195px; top: 20px;"> Text Color:</DIV>
<input type="text" name="textColor" value="" size=7 style="height: 22px; left: 252px; top: 15px;" maxlength=7>
<BUTTON style="top: 15px;left:312px ;width: 22px; height: 22px; border: 1px none;"  unselectable="on" onclick="setTextColor();" id=button3 name=button3><img id=imgtxtColor style="background-color:#000000;" src="../images/selcolor.gif" border=0></BUTTON>

<DIV id=divfsize style="left: 8px; top: 45px;"> Font Size:</DIV>
<SELECT size=1 ID=fsize name=fsize style="width: 80px;left: 100px; top: 42px;">
	<OPTION value="1" >1 (8 pt)</OPTION>
	<OPTION value="2" >2 (10 pt)</OPTION>
	<OPTION value="3" >3 (12 pt)</OPTION>
	<OPTION value="4"  selected>4 (14 pt)</OPTION>
	<OPTION value="5" >5 (18 pt)</OPTION>
	<OPTION value="6" >6 (24 pt)</OPTION>
	<OPTION value="7" >7 (36 pt)</OPTION>
</SELECT>

<DIV id=divstyle style="left: 195px; top: 45px;"> Font Style:</DIV>
<SELECT size=1 ID=fstyle name=fstyle style="width: 80px;left: 252px; top: 42px;">
	<OPTION value="Regular">Regular</OPTION>
	<OPTION value="I">Italic</OPTION>
	<OPTION value="STRONG" selected>Bold</OPTION>
	<OPTION value="U" >Underline</OPTION>
	<OPTION value="SUB" >Subscript</OPTION>
	<OPTION value="SUP" >Superscript</OPTION>
	<OPTION value="STRIKE" >Strikethrough</OPTION>
</SELECT>

<DIV id=divLinkText style="left: 8px; top: 71px;"> Link Text:</DIV>
<input type="text" name="LinkText" value="" style="width: 232px; height: 22px; left: 100px; top: 66px;" maxlength=255>

<DIV id=divtxtalign style="left: 8px; top: 96px;"> Text Align:</DIV>
<SELECT size=1 ID=fTxtAlign name=fTxtAlign style="width: 80px;left: 100px; top: 93px;">
	<OPTION value="left">left</OPTION>
	<OPTION value="center" selected>center</OPTION>
	<OPTION value="right" >right</OPTION>
</SELECT>
</FIELDSET>

</form>	
<form name="form2" ACTION="Formadd.asp" target="hidden_frame" METHOD="post" ENCTYPE="multipart/form-data">
	<DIV id=divUploadImg style="left: 23px; top: 185px; display: none">Upload Image:</DIV>
	<input TYPE="FILE" NAME="UploadFile1" ID="UploadFile1" style="height: 22px; left: 100px; top: 181px; width: 250px; display: none" tabIndex=5>

	<DIV id=divimgalign style="left: 23px; top: 210px; display: none"> Image Align:</DIV>
	<SELECT size=1 ID=fImgAlign name=fImgAlign style="width: 80px;left: 100px; top: 208px; display: none">
		<OPTION value="left">left</OPTION>
		<OPTION value="center" selected>center</OPTION>
		<OPTION value="right" >right</OPTION>
	</SELECT>
</form>
<BUTTON ID=btnOK style="left: 280px; top: 15px; width: 7em; height: 2.2em; " type=submit tabIndex=40 disabled>OK</BUTTON>
<BUTTON ID=btnCancel style="left: 280px; top: 50px; width: 7em; height: 2.2em; " type=reset tabIndex=1 onClick="window.close();">Cancel</BUTTON>
<%set con = nothing%>
</BODY>
</HTML>