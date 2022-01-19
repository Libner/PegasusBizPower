<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%
 'On Error Resume Next
  pageId=trim(Request("pageId"))
  prodId=trim(Request("prodId"))
  quest_id=trim(Request("quest_id"))
  If trim(quest_id) = "" And trim(prodId) <> "" Then
  	sqlstr = "Select QUESTIONS_ID FROM products WHERE PRODUCT_ID = " & prodId
	set rs_quest = con.getRecordSet(sqlstr)
	If not rs_quest.eof Then
		quest_id = trim(rs_quest(0))
	End if
  End If	
  'Response.Write prodId
  'Response.End
  PathCalImage = strLocal & "/Netcom/"
  UserID = trim(Request.Cookies("bizpegasus")("UserID"))
  OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))
  If trim(UserID) = "" or IsNumeric(UserID) = false Then
	UserID = trim(Request("U"))
  End If  	
  If trim(OrgID) = "" or IsNumeric(OrgID) = false Then
	OrgID = trim(Request("O"))
  End If  	
  CONTACT_ID = trim(Request("C"))
  If isNumeric(prodId) = true And isNumeric(pageId) = false Then
  sqlstr = "Select page_id FROM products where product_id = " & prodId
  set rspage = con.getREcordSet(sqlstr)
  If not rspage.eof Then
	pageId = trim(rspage(0))
  End If
  set rspage = Nothing	
  End If
  
  If trim(CONTACT_ID) <> "" Then
	url = strLocal & "/sale.asp?" & Encode("prodId=" & prodId & "&pageId=" & pageId & "&C=" & CONTACT_ID & "&O=" & OrgID)
    sqlstr = "Select PEOPLE_NAME FROM PEOPLES where PEOPLE_ID = " & cLng(CONTACT_ID)
	set rs = con.getREcordSet(sqlstr)
	If not rs.eof Then
		PeopleNAME = trim(rs(0))
	End If
	set rs = Nothing	
  End if
  
If isNumeric(pageId) And trim(pageId) <> "0" Then
  set pg=con.getRecordSet("SELECT Page_Title,Page_Source,Page_Lang,Product_ID,FORM_LINK_IMAGE FROM Pages WHERE Page_Id="&pageId&" ")
  If not pg.eof Then
	 pertitle=pg("Page_Title")	
	 pagesource=trim(pg("Page_Source"))
	 quest_id = trim(pg("Product_ID"))
	 formLink = trim(pg("FORM_LINK_IMAGE"))	 
	 PageLang = trim(pg("Page_Lang"))	 	
  End If 	 
  set pg = Nothing  
  poll_message = poll_message & strLocal &"/netcom/seker/seker.asp?" & Encode("P=" & trim(Request("prodId")) & "&C=" & trim(Request("C")) & "&O=" & OrgID)
    
  If Len(pagesource) > 0 And trim(quest_id) <> "" Then ' יש טופס בדף
	
	set prod = con.GetRecordSet("Select * from products where product_id=" & quest_id)
	if not prod.eof then
		productName=prod("Product_Name")				
		PRODUCT_DESCRIPTION = prod("PRODUCT_DESCRIPTION")
		attachment_form = prod("FILE_ATTACHMENT")
		attachment_title = prod("ATTACHMENT_TITLE")
		Langu = trim(prod("Langu"))
		if Langu = "eng" then
			dir_align = "ltr"
			td_align = "left"
			pr_language = "eng"
		else
			dir_align = "rtl"
			td_align = "right"
			pr_language = "heb"
		end if			
	End If		
	
	
	If trim(formLink) = "image" Then 'image  
		if prodId<>nil then
			set para=con.getRecordSet("select LINK_IMAGE,LINK_IMAGE_ALIGN from Pages where Page_Id="&pageId&" ")
			perpicture=para.Fields("LINK_IMAGE").ActualSize
			peralign=para("LINK_IMAGE_ALIGN")	   
			para.close
		end if
		form_content = "<a "
		If trim(CONTACT_ID) = "" Then
			form_content = form_content & " href=""#"" OnClick=""return false;""> "
		Else
			form_content = form_content &  " href="""&vFix(poll_message)& """ target=""_blank"">"
		End If
		form_content = form_content & " <img src="""&strLocal&"netcom/GetImage.asp?DB=Page&FIELD=LINK_IMAGE&ID="&pageId&""" border=""0""></a>"
	 
	ElseIf trim(formLink) = "link" Then 'text

		if pageId<>nil then
		set para=con.getRecordSet("select LINK_TEXT,LINK_TEXT_ALIGN,LINK_BGCOLOR,LINK_FONT_TYPE,LINK_FONT_SIZE,LINK_FONT_COLOR from Pages where Page_Id="&pageId&" ")
			pertext = para("LINK_TEXT")
			perTextImgAlign = trim(para("LINK_TEXT_ALIGN"))
			perbgcolor = trim(para("LINK_BGCOLOR"))
			pertype = trim(para("LINK_FONT_TYPE"))
			persize = trim(para("LINK_FONT_SIZE"))
			percolorname = trim(para("LINK_FONT_COLOR"))	
			percolor = trim(para("LINK_FONT_COLOR"))	   
		para.close
		end if
		If Len(pertext) = 0 Or isNULL(pertext) Then
			pertext  = productName
		End If	
		If Len(percolor) = 0 Or isNULL(percolor)  Then
			percolor = "#000000"
		End If
		If Len(pertype) = 0 Or isNULL(pertype)  Then
			pertype = "STRONG"
		End If	
		If Len(persize) = 0 Or isNULL(persize)  Then
			persize = "2"
		End If	
		If Len(perbgcolor) = 0 Or isNULL(perbgcolor)  Then
			perbgcolor = "transparent"
		End If
		If Len(perTextImgAlign) = 0 Or isNULL(perTextImgAlign)  Then
			perTextImgAlign = "center"
		End If		
		
		form_content = "<p align="""&perTextImgAlign&""" dir=rtl width=100% bgcolor="""&perbgcolor&""">"&_
		"&nbsp;<a bgcolor="""&perbgcolor&""" target=""_blank"" style=""letter-spacing:1px"""
		If trim(CONTACT_ID) = "" Then
			form_content = form_content & " href=""#"" OnClick=""return false;"" "
		Else
			form_content = form_content &  " href="&vFix(poll_message)
		End If
		form_content = form_content & "><font color="""&percolor&""" size="""&persize&""">"
		If trim(pertype) = "I" Then
			form_content = form_content & "<i>"
		ElseIf	trim(pertype) = "STRONG" Then
			form_content = form_content & "<b>"
		ElseIf	trim(pertype) = "SUB" Then
			form_content = form_content & "<SUB>"
		ElseIf	trim(pertype) = "SUP" Then
			form_content = form_content & "<SUP>"
		ElseIf	trim(pertype) = "STRIKE" Then
			form_content = form_content & "<STRIKE>"
		End If				
		
		form_content = form_content & pertext
		
		If trim(pertype) = "I" Then
			form_content = form_content & "</i>"
		ElseIf	trim(pertype) = "STRONG" Then
			form_content = form_content & "</b>"
		ElseIf	trim(pertype) = "SUB" Then
			form_content = form_content & "</SUB>"
		ElseIf	trim(pertype) = "SUP" Then
			form_content = form_content & "</SUP>"
		ElseIf	trim(pertype) = "STRIKE" Then
			form_content = form_content & "</STRIKE>"
		End If			
		
		form_content = form_content & "</font></a>&nbsp;</p>"
	Else
					
		if pageId<>nil then			
		set para=con.getRecordSet("select FORM_BGCOLOR,FORM_FONT_TYPE,FORM_FONT_SIZE,FORM_FONT_COLOR from Pages where Page_Id="&pageId&" ")
		If not para.eof Then
			perbgcolor = trim(para("FORM_BGCOLOR"))
			pertype = trim(para("FORM_FONT_TYPE"))
			persize = trim(para("FORM_FONT_SIZE"))
			percolor = trim(para("FORM_FONT_COLOR"))			
			If trim(perbgcolor) = "" Then
				perbgcolor = "#F0F0F0"
			End If	
			If trim(percolor) = "" Then
				percolor = "#000000"
			End If				
			pertitlecolor = calcColor(perbgcolor,-80)			
			persubjectcolor = calcColor(perbgcolor,-30)			
				
		Else
			perbgcolor = "#F0F0F0"
			pertype = "STRONG"
			persize = "2"
			percolor = "#000000"
			pertitlecolor = "#808080"
			persubjectcolor = "#D3D3D3"
		End If		
		para.close
		end if	
		form_content = createFormContent(pertype,perbgcolor,pertitlecolor,persize,percolor,persubjectcolor)	
	End If		
	set prod = nothing
    
    'Response.Write image_name
    'Response.End
    image_name = "id=""tofes"""
	ind = inStr(pagesource,image_name)
	If ind = 0 Then
		image_name = "id=tofes"
		ind = inStr(pagesource,image_name)
	End If
	
	If ind > 0 Then
	ind_start = inStrRev(pagesource,"<",ind)			
	ind_end = inStr(ind,pagesource,">")
	'Response.Write "ind = " & ind & " ind_start = " & ind_start & " ind_end = " & ind_end
	'Response.End
	length_to_cut = ind_end - ind_start + 1	
	string_to_replace = Mid(pagesource,ind_start,length_to_cut)
	'Response.Write form_content 
	'Response.End	
	pagesource = Replace(pagesource,string_to_replace,form_content)
	End if
End If
ElseIf trim(quest_id) <> "" Then 'רק טופס ללא דף	
	If quest_id <> "" Then
		set prod = con.GetRecordSet("Select * from products where product_id=" & quest_id)
		if not prod.eof then
		productName=prod("Product_Name")				
		PRODUCT_DESCRIPTION = prod("PRODUCT_DESCRIPTION")		
		Langu = trim(prod("Langu"))
		attachment_form = prod("FILE_ATTACHMENT")
		attachment_title = prod("ATTACHMENT_TITLE")
		if Langu = "eng" then
			dir_align = "ltr"
			td_align = "left"
			pr_language = "eng"
		else
			dir_align = "rtl"
			td_align = "right"
			pr_language = "heb"
		end if	
	 End If	
	 pagesource = createFormContent("","#E5E5E5","#808080","2","#000000")
	End If		
End If

function createFormContent(pertype,perbgcolor,titlecolor,persize,percolor,persubjectcolor)
	If isNumeric(persize) = false Then
		persize = 2
	Else
		persize = cInt(persize)
	End if	
	form_content = ""
		
	form_content = form_content &"<form action=""" & strLocal & "/netcom/seker/addappeal.asp"" id=""form1"" name=""form1"" method=""post"" target=""_blank"" onSubmit=""return click_fun()"">"&chr(13)&chr(10)
	form_content = form_content &"<table dir=ltr WIDTH=""550"" BGCOLOR="""&perbgcolor&""" ALIGN=center BORDER=""0"" bordercolor=""#000000"" style=""border-collapse:collapse;"" CELLPADDING=""1"" cellspacing=""0"">"&chr(13)&chr(10)
	form_content = form_content &"<tr><td align=center dir="&dir_align&" bgcolor="""&titlecolor&""" colspan=4><font color="""&percolor&""" size="""&persize&""">"&chr(13)&chr(10)
	If trim(pertype) = "I" Then
		form_content = form_content & "<i>"
	ElseIf trim(pertype) = "STRONG" Then
		form_content = form_content & "<b>"
	ElseIf	trim(pertype) = "SUB" Then
		form_content = form_content & "<SUB>"
	ElseIf	trim(pertype) = "SUP" Then
		form_content = form_content & "<SUP>"
	ElseIf	trim(pertype) = "STRIKE" Then
		form_content = form_content & "<STRIKE>"
	End If				
	
	form_content = form_content & "<b>" & productName & "</b>"
	
	If trim(pertype) = "I" Then
		form_content = form_content & "</i>"
	ElseIf	trim(pertype) = "STRONG" Then
		form_content = form_content & "</b>"
	ElseIf	trim(pertype) = "SUB" Then
		form_content = form_content & "</SUB>"
	ElseIf	trim(pertype) = "SUP" Then
		form_content = form_content & "</SUP>"
	ElseIf	trim(pertype) = "STRIKE" Then
		form_content = form_content & "</STRIKE>"
	End If			
	
	form_content = form_content & "</td></tr>"&chr(13)&chr(10)
	if PRODUCT_DESCRIPTION <> ""  then
		form_content = form_content & "<tr><td align=right colspan=4 style=""padding-right:15px;padding-left:15px"">"&chr(13)&chr(10)
		form_content = form_content & "<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=3><TR>"&chr(13)&chr(10)
		form_content = form_content & "<td class=""form_makdim"" dir="&dir_align&" align=right><font color="""&percolor&""">"
		If trim(pertype) = "I" Then
			form_content = form_content & "<i>"
		ElseIf	trim(pertype) = "STRONG" Then
			form_content = form_content & "<b>"
		ElseIf	trim(pertype) = "SUB" Then
			form_content = form_content & "<SUB>"
		ElseIf	trim(pertype) = "SUP" Then
			form_content = form_content & "<SUP>"
		ElseIf	trim(pertype) = "STRIKE" Then
			form_content = form_content & "<STRIKE>"
		End If				
		
		form_content = form_content & trim(PRODUCT_DESCRIPTION)
		
		If trim(pertype) = "I" Then
			form_content = form_content & "</i>"
		ElseIf	trim(pertype) = "STRONG" Then
			form_content = form_content & "</b>"
		ElseIf	trim(pertype) = "SUB" Then
			form_content = form_content & "</SUB>"
		ElseIf	trim(pertype) = "SUP" Then
			form_content = form_content & "</SUP>"
		ElseIf	trim(pertype) = "STRIKE" Then
			form_content = form_content & "</STRIKE>"
		End If			
		
		form_content = form_content & "</td></tr>"		
		If Len(attachment_title) > 0 Then
		form_content = form_content & "<TR><td class=""form_makdim"" dir=""rtl"" width=100% align=right>" &_
		"<a style=""color:black;font-weight:bolder;font-size:11pt"" href="&strLocal&"/download/products/"&attachment_form&" target=_blank>"& attachment_title&"</a></td></tr>"
		End If
		
			
		end if
					
	set fields=con.GetRecordSet("SELECT Field_Id,Field_Title,Field_Type,Field_Size,Field_Align,Field_Scale,FIELD_MUST FROM FORM_FIELD Where product_id=" & quest_id &" Order by Field_Order")
					
	do while not fields.EOF 
		Field_Id = fields("Field_Id")
		Field_Title = trim(fields("Field_Title"))
		Field_Title_Str = ""
		Field_Title_Str = Field_Title_Str & "<font color="""&percolor&""" size="""&persize-1&""">"
		If trim(pertype) = "I" Then
			Field_Title_Str = Field_Title_Str & "<i>"
		ElseIf	trim(pertype) = "STRONG" Then
			Field_Title_Str = Field_Title_Str & "<b>"
		ElseIf	trim(pertype) = "SUB" Then
			Field_Title_Str = Field_Title_Str & "<SUB>"
		ElseIf	trim(pertype) = "SUP" Then
			Field_Title_Str = Field_Title_Str & "<SUP>"
		ElseIf	trim(pertype) = "STRIKE" Then
			Field_Title_Str = Field_Title_Str & "<STRIKE>"
		End If				

		Field_Title_Str = Field_Title_Str & Field_Title
		
		If trim(pertype) = "I" Then
			Field_Title_Str = Field_Title_Str & "</i>"
		ElseIf	trim(pertype) = "STRONG" Then
			Field_Title_Str = Field_Title_Str & "</b>"
		ElseIf	trim(pertype) = "SUB" Then
			Field_Title_Str = Field_Title_Str & "</SUB>"
		ElseIf	trim(pertype) = "SUP" Then
			Field_Title_Str = Field_Title_Str & "</SUP>"
		ElseIf	trim(pertype) = "STRIKE" Then
			Field_Title_Str = Field_Title_Str & "</STRIKE>"
		End If						
		Field_Title_Str = Field_Title_Str & "</font>"&chr(13)&chr(10)
		
		Field_Type = fields("Field_Type")
		Field_Size = fields("Field_Size")
		Field_Align = trim(fields("Field_Align"))
		Field_Scale = fields("Field_Scale")
		
		if trim(Langu) = "eng" then
			form_content = form_content &"<tr><td width=""100%"" "
			if trim(Field_Type) = "10" then
			form_content = form_content &" colspan=2 class=""title_form"" style=""background:"&persubjectcolor&";"" "
			Else
			form_content = form_content &" class=""form"" "
			End If
			form_content = form_content &" style=""padding-left:10px"" align=left>"
			if trim(Field_Type) = "5" then
				form_content = form_content & "<font color="""&percolor&""" size="""&persize&""">" & con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng","","" & Field_Title & "") & "</font"
			end if
			form_content = form_content &"<b>"
			if FIELD_MUST then
			form_content = form_content &"<font color=red>*&nbsp;</font>"
			end if
			form_content = form_content & Field_Title_Str & "&nbsp;</b></td></tr>"
			if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then
			form_content = form_content & "<tr><td align=""left"" valign=middle width=100% class=""form"" style=""padding-left:10px"">"					
			form_content = form_content & "<font color="""&percolor&""" size="""&persize&""">" & con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng","","" & Field_Title & "") & "</font"
			form_content = form_content & "</td></tr>"
			end if
		 else
			form_content = form_content & "<tr><td width=""100%"" "
			if trim(Field_Type) = "10" then
			form_content = form_content & " colspan=2 class=""title_form"" style=""background:"&persubjectcolor&";"" "
			Else
			form_content = form_content & " class=""form"" "
			End If
			form_content = form_content & "  align=right valign=""top"" dir=ltr style=""padding-right:15px;"">"
			form_content = form_content & "<span dir=""rtl""><b>"
			if FIELD_MUST then
			form_content = form_content & "<font color=red>&nbsp;*&nbsp;</font>"
			end if
			form_content = form_content & Field_Title_Str & "</b></span>"
			if trim(Field_Type) = "5" then
			form_content = form_content & "<font color="""&percolor&""" size="""&persize&""">" & con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb","","" & Field_Title & "") & "</font>"
			end if
			form_content = form_content & "</td></tr>"
			if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then
			form_content = form_content & "<tr><td align=""right"" valign=middle width=100% class=""form"" style=""padding-right:15px;"">"					
			form_content = form_content & "<font color="""&percolor&""" size="""&persize&""">" & con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb","","" & Field_Title & "") & "</font>"
			form_content = form_content & "</td></tr>"
			end if
		end if
		fields.moveNext()
		loop
	set fields=nothing 
	
	form_content = form_content & "</TABLE></td></tr>"&chr(13)&chr(10)	
	form_content = form_content &"<tr><td colspan=""4"" height=""10"" width=""550"">"&chr(13)&chr(10)
	form_content = form_content &"<input type=""text"" style=""width:0"" name=""C"" value="&CONTACT_ID&">"&chr(13)&chr(10)
	form_content = form_content &"<input type=""text"" style=""width:0"" name=""P"" value="&prodId&">"&chr(13)&chr(10)
	form_content = form_content &"<input type=""text"" style=""width:0"" name=""O"" value="""&OrgID&""">"&chr(13)&chr(10)
	form_content = form_content &"<table width=""550"" align=""center"" cellpadding=""0"" cellspacing=""0"">"&chr(13)&chr(10)	
	form_content = form_content &"<tr><td width=""550"" bgcolor=""transparent"" height=""1"" colspan=""2""></td>"&chr(13)&chr(10)
	form_content = form_content &"</tr></table></td></tr>"&chr(13)&chr(10)
	if Langu = "eng" then
		form_content = form_content &"<tr><td colspan=""4"" align=center>"
		If trim(CONTACT_ID) <> "" Then
			form_content = form_content & "<input type=""image"" SRC=""" & strLocal & "/netcom/images/SEND-ENG.gif"" border=""0"" name=""submit_button"" id=""submit_button"" hspace=""60""></td></tr>"&chr(13)&chr(10)
		Else
			form_content = form_content & "<IMG SRC=""" & strLocal & "/netcom/images/SEND-ENG.gif"" border=""0"" OnClick=""return false;"" name=""submit_button"" id=""submit_button"" hspace=""60""></td></tr>"&chr(13)&chr(10)
		End If 		
	Else
		form_content = form_content &"<tr><td colspan=""4"" align=center>"
		If trim(CONTACT_ID) <> "" Then
			form_content = form_content & "<input type=""image"" SRC=""" & strLocal & "/netcom/images/b-send-a.gif"" border=""0"" name=""submit_button"" id=""submit_button"" hspace=""60""></td></tr>"&chr(13)&chr(10)
		Else
			form_content = form_content & "<IMG SRC=""" & strLocal & "/netcom/images/b-send-a.gif"" border=""0"" OnClick=""return false;"" name=""submit_button"" id=""submit_button"" hspace=""60""></td></tr>"&chr(13)&chr(10)
		End If 				
	End If
	form_content = form_content &"<tr><td colspan=""4"" height=""10""></td></tr>"&chr(13)&chr(10)	
	form_content = form_content & "</table></td></tr><tr><td height=10 nowrap></td></tr>"
	form_script = ""
    if Langu = "heb" then
		str_alert = "!!!נא למלא את השדה"
	Else
		str_alert = "Please provide the answer!!!"
	End If	
	form_script = form_script &"<SCRIPT LANGUAGE=javascript><!-- "&chr(13)&chr(10)&_	
	"function click_fun()"&chr(13)&chr(10)&_
	"{"&chr(13)&chr(10)
        sqlStr = "SELECT Field_Id,Field_Must,Field_Type FROM FORM_FIELD Where product_id=" & quest_id & " Order by Field_Order"
		'Response.Write sqlStr
		'Response.End
		set fields=con.GetRecordSet(sqlStr)
		do while not fields.EOF 
		  Field_Id = fields("Field_Id")				
		  Field_Must = trim(fields("Field_Must"))
		  Field_Type = trim(fields("Field_Type"))
		  'Response.Write  Field_Type
		  If Field_Must = "True"  Then		  
		  form_script = form_script & "field =  window.document.all(""field"&Field_Id&""");"&chr(13)&chr(10)
          If trim(Field_Type) = "8" Or trim(Field_Type) = "9" Or trim(Field_Type) = "11" Or trim(Field_Type) = "12" Then
		  form_script = form_script & "if(field != null)"&chr(13)&chr(10)&_
		  "{"&chr(13)&chr(10)&_
			"answered = false;"&chr(13)&chr(10)&_
			"for(j=0;j<field.length;j++)"&chr(13)&chr(10)&_
			"{"&chr(13)&chr(10)&_
				"if(field(j).checked == true)"&chr(13)&chr(10)&_
				"{"&chr(13)&chr(10)&_
					"answered = true;"&chr(13)&chr(10)&_
				"}"&chr(13)&chr(10)&_
			"}"&chr(13)&chr(10)&_
			"if(answered == false)"&chr(13)&chr(10)&_
			"{"&chr(13)&chr(10)&_
				"window.alert(""" & str_alert & """);"&chr(13)&chr(10)&_
				"field(0).focus();"&chr(13)&chr(10)&_
				"return false;"&chr(13)&chr(10)&_
			"}"&chr(13)&chr(10)&_
		  "}"&chr(13)&chr(10)
	      Else
	      form_script = form_script & "if((field != null) && document.all(""field"&Field_Id&""").value == '')"&chr(13)&chr(10)&_
	      "{ document.all(""field" & Field_Id & """).focus(); "&chr(13)&chr(10)&_
		  "  window.alert(""" & str_alert & """); "&chr(13)&chr(10)&_
		  "  return false; } " &chr(13)&chr(10)	   
	     End If	 
	   End If
       fields.moveNext()
	   loop
       set fields=nothing    
    form_script = form_script & "return true;	}"&chr(13)&chr(10)&_
    "function popupcal(elTarget){"&chr(13)&chr(10)&_
		"if (showModalDialog){"&chr(13)&chr(10)&_
			"var sRtn;"&chr(13)&chr(10)&_
			"sRtn = showModalDialog(""/netcom/calendar.asp"",elTarget.value,""center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;"");"&chr(13)&chr(10)&_
			"if (sRtn!='')"&chr(13)&chr(10)&_
			  "elTarget.value = sRtn;"&chr(13)&chr(10)&_
		 "}else"&chr(13)&chr(10)&_
		   "alert(""Internet Explorer 4.0 or later is required."");"&chr(13)&chr(10)&_
		"return false;"&chr(13)&chr(10)&_
		"window.document.focus;   "&chr(13)&chr(10)&_
"}"&chr(13)&chr(10)&_
	"//-->"&chr(13)&chr(10)&_
	"</SCRIPT> "&chr(13)&chr(10)
	form_content = form_content & form_script
	createFormContent = form_content
end function
%>
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=windows-1255">
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
<title><%=pertitle%></title>
</head>
<body style="margin:0px;SCROLLBAR-FACE-COLOR: #E6E6E6;SCROLLBAR-HIGHLIGHT-COLOR: #F7F7F7;SCROLLBAR-SHADOW-COLOR: #848484;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #808080;SCROLLBAR-TRACK-COLOR: #E6E6E6;SCROLLBAR-DARKSHADOW-COLOR: #ffffff"> 
<table width=620 border=0 cellpadding=0 cellspacing=0 align=center>
<tr><td width=100%>
<%If Len(pagesource) > 0 Then%>
<%=Replace(pagesource, "{$var_rcptname}", PeopleNAME)%>
<%End If%>
</td></tr></table>
<!--#include file="bottom_inc.asp"-->
</body>
</html>
