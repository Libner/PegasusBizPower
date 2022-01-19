<%SERVER.ScriptTimeout=3000%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
</head>
<body>
<%	pageid = Request.QueryString("pageid")
	formId = Request.QueryString("formId")
			
	if formId <> "" then
		
		set con_image=Server.CreateObject("adodb.connection")
		con_image.open "FILE NAME="& server.MapPath("../../netReply.udl")
		set pr=Server.CreateObject("ADODB.Recordset")
		pr.Open "SELECT LINK_IMAGE FROM PAGES where Page_Id="&pageid,con_image,2,3
   		set upl=Server.CreateObject("SoftArtisans.FileUp") 
		if Trim(upl.Form("UploadFile1")) <> "" then
			upl.SaveAsBlob pr.Fields("LINK_IMAGE")
			pr.Update
			pr.Close
			
			set pr1=Server.CreateObject("ADODB.Recordset")  
			pr1.Open "SELECT LINK_IMAGE FROM PAGES where Page_Id="&pageid,con_image,2,3
			Set objImageGen = Server.CreateObject("softartisans.ImageGen")
			objImageGen.LoadImage pr1.Fields("LINK_IMAGE")	
			if  objImageGen.Width >=610  or objImageGen.Height >= 600 then			
				objImageGen.ResizeImage 610, 600, saiBicubic,1								
				objImageGen.SaveImage saiDatabaseBlob, objImageGen.ImageFormat,  pr.Fields(F)	
			 pr1.Update
			 pr1.Close  
			end if
			Set objImageGen = Nothing
			set pr1 = Nothing   
		end if	
		Set pr = Nothing
		set con_image = nothing
		fImgAlign = upl.Form("fImgAlign")
		upEx = "UPDATE PAGES SET Product_Id=" & formId & ", FORM_LINK_IMAGE='image'" &_
		", LINK_IMAGE_ALIGN='" & fImgAlign &_
		"', LINK_TEXT=''" &_
		", LINK_TEXT_ALIGN=''" &_
		", LINK_BGCOLOR=''" &_
		", LINK_FONT_TYPE=''" &_
		", LINK_FONT_SIZE=''" &_
		", LINK_FONT_COLOR=''" &_
		", FORM_BGCOLOR=''" &_
		", FORM_FONT_TYPE=''" &_
		", FORM_FONT_SIZE=''" &_
		", FORM_FONT_COLOR=''" 
		set upl=nothing
	else 'formId <> ""
		prod_id = Request.Form("selFormId")
		cbFormLink = Request.form("cbFormLink")
		bgcolor = trim(Request.form("bgrColor"))
		fcolor = trim(Request.form("textColor"))
		fsize = Request.form("fsize")
		fstyle = Request.form("fstyle")
		LinkText = trim(Request.form("LinkText"))
		fTxtAlign = Request.form("fTxtAlign")
	
		upEx = "UPDATE PAGES SET Product_Id=" & prod_id & ", FORM_LINK_IMAGE='" & cbFormLink
	
		if cbFormLink = "link" then
			upEx = upEx & "', LINK_IMAGE_ALIGN=''" &_
			", LINK_TEXT='" & LinkText &_
			"', LINK_TEXT_ALIGN='" & fTxtAlign &_
			"', LINK_BGCOLOR='" & sFix(bgcolor) &_
			"', LINK_FONT_TYPE='" & fstyle &_
			"', LINK_FONT_SIZE='" & fsize &_
			"', LINK_FONT_COLOR='" & sFix(fcolor) &_
			"', FORM_BGCOLOR=''" &_
			", FORM_FONT_TYPE=''" &_
			", FORM_FONT_SIZE=''" &_
			", FORM_FONT_COLOR=''"
		elseif cbFormLink = "form" then
			upEx = upEx & "', LINK_IMAGE_ALIGN=''" &_
			", LINK_TEXT=''" &_
			", LINK_TEXT_ALIGN=''" &_
			", LINK_BGCOLOR=''" &_
			", LINK_FONT_TYPE=''" &_
			", LINK_FONT_SIZE=''" &_
			", LINK_FONT_COLOR=''" &_
			", FORM_BGCOLOR='" & sFix(bgcolor) &_
			"', FORM_FONT_TYPE='" & fstyle &_
			"', FORM_FONT_SIZE='" & fsize &_
			"', FORM_FONT_COLOR='" & sFix(fcolor) &"'"
		end if
		
	end if 'formId <> ""
	
	upEx = upEx & " WHERE Page_Id=" & pageid
		
	'Response.Write upEx
	'Response.End
	if con.executeQuery(upEx) then
		if cbFormLink = "link" then
			htmlForm = "<IMG id=tofes src=../../images/linkform_image.gif align=baseline border=0>"
		elseif cbFormLink = "form" then
			htmlForm = "<IMG id=tofes src=../../images/form_image.gif align=baseline border=0>"
		else
			htmlForm = "<p align=" & fImgAlign & "><IMG id=tofes src=../../GetImage.asp?DB=Page&FIELD=LINK_IMAGE&ID=" & pageid & "  border=0></p>"	
		end if	
	%> 
	<SCRIPT LANGUAGE=javascript>
	<!--
		window.parent.main_frame.txtForm.value = "<%=htmlForm%>";
		window.parent.main_frame.btnOKClick();
	//-->
	</SCRIPT>
<%	else %>
	<SCRIPT LANGUAGE=javascript>
	<!--
		alert("437 Error !")
	//-->
	</SCRIPT>
<%	end if
	set con = nothing%>	
	
</body>
</html>
