<%serv_path = Request.ServerVariables("SERVER_NAME")
	if InStr(1,serv_path,"popeye") > 0 then
		str_mappath = "http://" & serv_path & "/eastronics/bullets/"
	else
		str_mappath = "http://" & serv_path & "/bullets/"
	end if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD W3 HTML 3.2//EN">
<HTML  id=dlgBullet STYLE="width: 292px; height: 224px; ">
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="MSThemeCompatible" content="Yes">
<TITLE>Insert Bullet</TITLE>
<style>
  html, body, button, div, input, select, fieldset { font-family: MS Shell Dlg; font-size: 8pt; position: absolute; };
</style>
<SCRIPT defer>

function _CloseOnEsc() {
  if (event.keyCode == 27) { window.close(); return; }
}

window.onerror = HandleError

function HandleError(message, url, line) {
  var str = "An error has occurred in this dialog." + "\n\n"
  + "Error: " + line + "\n" + message;
  alert(str);
  window.close();
  return true;
}


function btnOKClick() {
  var imgtext = '';
  var str_mappath = '<%=str_mappath%>';
  // error checking

  if (txtFileName.value=="") { 
    alert("Bullet Image must be specified.");
    return;
  }
  
  window.returnValue = str_mappath + txtFileName.value;
  window.close();
}
</SCRIPT>
</HEAD>
<BODY id=bdy style="background: threedface; color: windowtext;" scroll=no>
<DIV id=divTable style="left:-1.8em; top: 2.5em;">
<TABLE cellSpacing=0 cellPadding=2 border=0 align=left>
	<TR>
		<TD noWrap width="50"  align=right>
			<IMG src="../../Bullets/bul5.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul5.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul4.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul4.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul3.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul3.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul2.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul2.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul1.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul1.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
	</TR>
	<TR>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul10.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul10.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul9.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul9.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul8.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul8.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul7.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul7.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul6.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul6.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
	</TR>
	<TR>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul15.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul15.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul14.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul14.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul13.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul13.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul12.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul12.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul11.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul11.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
	</TR>
	<TR>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul20.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul20.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul19.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul19.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul18.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul18.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul17.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul17.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
		<TD noWrap width="50" align=right>
			<IMG src="../../Bullets/bul16.gif" border=0 align=absmiddle>&nbsp;<INPUT type=radio value="bul16.gif" name=bullet onclick="txtFileName.value=this.value;">
		</TD>
	</TR>
</TABLE>
</DIV>
<INPUT ID=txtFileName type=hidden value="">

<FIELDSET id=fldSpacing style="left: .9em; top: 1.0em; width: 25.08em; height: 11.2em;">
<LEGEND id=lgdSpacing> Select Bullet </LEGEND>
</FIELDSET>

<BUTTON ID=btnOK style="left: 18.9em; top: 13.25em; width: 7em; height: 2.2em; " type=submit tabIndex=40 onclick="btnOKClick()">OK</BUTTON>
<BUTTON ID=btnCancel style="left: 18.9em; top: 15.8em; width: 7em; height: 2.2em; " type=reset tabIndex=45 onClick="window.returnValue='';window.close();">Cancel</BUTTON>

</BODY>
</HTML>