<!--#INCLUDE FILE="../reverse.asp"-->
<!--#include file="../connect.asp"-->
<%
Response.Expiresabsolute = Now() - 1 
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache" 
Langu = Request("Langu")
OrgID = trim(Request("O"))

strLocal = "http://" & Trim(Request.ServerVariables("SERVER_NAME"))
If Len(Trim(Request.ServerVariables("SERVER_PORT"))) > 0 Then
    strLocal = strLocal & ":" & Trim(Request.ServerVariables("SERVER_PORT"))
End If
strLocal = strLocal & Application("VirDir") & "/"
	

PEOPLE_ID = trim(Request("C"))
prodId = trim(Request("P"))
OrgID = trim(Request("O"))
reason = ""
show_sale = "True"	
'For each item In Request.Form
'	Response.Write item & " - " & Request.Form(item) & "<br>"
'Next
			

   if trim(prodId) = "" OR IsNumeric(prodId) = false then
		show_sale = "False"
		reason = "SALE NOT FOUND"		
	else
		sqlStr = "Select QUESTIONS_ID,PAGE_ID from products where product_id=" & prodId
		'Response.Write sqlStr
		'Response.End
		set rs_products = con.GetRecordSet(sqlStr)
			if not rs_products.eof then						
				pageId = trim(rs_products("Page_Id"))
				quest_id = rs_products("QUESTIONS_ID") '//sheloon number of the SALE																		
				If IsNull(quest_id) And trim(pageId) <> "" Then
					sqlstr = "Select Product_ID From Pages WHERE Page_ID = " & pageId
					set rs_p = con.getRecordSet(sqlstr)
					If not rs_p.eof Then
						quest_id = rs_p(0)						
					End if
					Set rs_p = Nothing	
				End If
				sqlstr = "Select Product_Name,PRODUCT_DESCRIPTION,PRODUCT_THANKS, RESPONSIBLE From Products WHERE Product_ID=" & quest_id				
				set rsd = con.getRecordSet(sqlstr)
				If not rsd.eof Then
				    Product_Name = rsd(0)
					PRODUCT_DESCRIPTION = rsd(1)
					PRODUCT_THANKS = rsd(2)
					If trim(PRODUCT_THANKS) = "" Then
						if Langu = "eng" then
							PRODUCT_THANKS = "The form has been received"
						else
							PRODUCT_THANKS = "הטופס נקלט במערכת"
						end if
					End If					
					RESPONSIBLE = rsd(3)
				End If
				set rsd = Nothing	
			end if
		set rs_products = Nothing			
End If

If isNumeric(PEOPLE_ID) And isNUmeric(prodId) And IsNumeric(OrgID) Then

							
if PEOPLE_ID = "0" then
	
	show_sale = "False" 
	reason = "CLIENT NOT AUTHORIZ"	
	
else
	sqlStr = "SELECT PEOPLE_ID, PEOPLE_NAME, CONTACT_ID, COMPANY_ID From PRODUCT_CLIENT"&_
	" Where PRODUCT_ID = "& prodId &" And PEOPLE_ID=" & PEOPLE_ID	
			'Response.Write sqlStr
			'Response.End
			set rs_client = con.GetRecordSet(sqlStr)
			if not rs_client.eof then		
			'Response.Write trim(OrgID) & " " & trim(rs_client("OrgID")) 				
					'if trim(OrgID) = trim(rs_client("ORGANIZATION_ID")) then
						show_sale = "True"												
						client_name = rs_client("PEOPLE_NAME")
						if trim(client_name) <> "" then
							client_name = "(" & client_name & ")"
						end if
						contactID = trim(rs_client("CONTACT_ID"))
						companyID = trim(rs_client("COMPANY_ID"))
														
						sqlStr = "Select  APPEAL_ID from APPEALS where PEOPLE_ID="& PEOPLE_ID &" and PRODUCT_ID=" & prodId	
						set rs_client = con.GetRecordSet(sqlStr)
						if not rs_client.eof then
							show_sale = "False" '//לא לאפשר הצבעה חוזרת ללקוח שכבר הצביע
							reason = "ANSWERED"							
						end if
						set rs_client = nothing	
																											
					'else
						'show_sale = "False"
						'reason = "CLIENT NOT AUTHORIZ"													
					'end if	
			else
				show_sale = "False"
				reason = "CLIENT NOT AUTHORIZ"											
			end if
			set rs_client = nothing
 end if			
		
		if show_sale = "True" Then
		rec = 0
	      
			if prodId <> nil  then
			    sqlstr = "Select Langu,QUESTIONS_ID, DATE_START, DATE_END from products where product_id=" & prodId
			  
				set prod = con.GetRecordSet(sqlstr)
				if not prod.eof then																			
					set rsd = con.getRecordSet("Select Langu, RESPONSIBLE From Products WHERE product_id = " & quest_id)
					if not rsd.eof Then
						Langu = trim(rsd(0))
						RESPONSIBLE = trim(rsd(1))
					end if
					set rsd = Nothing	
						
					if Langu = "eng" then
						dir_align = "ltr"
						td_align = "left"
						Langu = "eng"
					else
						dir_align = "rtl"
						td_align = "right"
						Langu = "heb"
					end if		
						
					DATE_START = trim(prod("DATE_START"))
					DATE_END = trim(prod("DATE_END"))
					if datediff("d",DATE_START,date()) < 0 or datediff("d",DATE_END,date()) >= 0 then
						if datediff("d",DATE_START,date()) < 0 then
								reason = "SALE NOT START"
								show_sale = "False"	
						elseif datediff("d",DATE_END,date()) >= 0 then
								reason = "SALE END"
								show_sale = "False"																																													
						end if
					end if
				end if
				set prod = nothing
			end if	
			
			If show_sale = "True" Then
			
			If Len(contactID) = 0 Or IsNull(contactID) Then
				contactID = "NULL"
			End If
			
			If Len(companyID) = 0 Or IsNull(companyID) Then
				companyID = "NULL"
			End If
			
			in_date = "getDate()"	
		    sqlstring=" SET NOCOUNT ON; INSERT INTO APPEALS (APPEAL_DATE,PEOPLE_ID,CONTACT_ID,COMPANY_ID,PRODUCT_ID,"&_
		    " QUESTIONS_ID,appeal_status,appeal_deleted,type_id,ORGANIZATION_ID) "&_
		    "VALUES (" & in_date &","& PEOPLE_ID & "," & contactID & "," & companyID & "," & prodId & "," & quest_id &_
		    ",'1',0,1,"&OrgID&"); SELECT @@IDENTITY AS NewID"
			'Response.Write(sqlstring)
			'Response.End 
			set rs_tmp = con.getRecordSet(sqlstring)
				appid = rs_tmp.Fields("NewID").value
			set rs_tmp = Nothing			
									
			sqlstr="SELECT Field_Id,FIELD_TYPE  FROM FORM_FIELD Where product_id=" & quest_id & " Order by Field_Order"
			'Response.Write sqlstr
			'Response.End
			set fields=con.GetRecordSet(sqlstr)
				do while not fields.EOF
					Field_Id = fields("Field_Id")				
					Field_Value = trim(Request("field" & Field_Id))		
					'Response.Write Field_Id & " " & Field_Value 
					'Response.End						
					sqlstr="INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (" & appid & "," & Field_Id & ",'"& sFix(Field_Value) &"')"
					'Response.Write(sqlstr) & "<br><br>"
					'Response.End
					con.ExecuteQuery sqlstr				
				fields.moveNext()
				loop
			set fields=nothing
		
		  
		in_date = " getDate()"	
	    con.ExecuteQuery("Update PRODUCT_CLIENT Set DATE_ANSWER="& in_date &" Where PRODUCT_ID=" & prodId & " and PEOPLE_ID=" & PEOPLE_ID)	 			
	    	    
	    If trim(appid) <> "" Then
        sqlstr = "EXECUTE get_feedbacks '','','','','" & OrgID & "','','','','','','','" & appID & "'"		
		'Response.Write("sqlstr=" & sqlstr)
		'Response.End
		set app = con.GetRecordSet(sqlstr)
		if not app.eof then
		
		appeal_date = Day(app("appeal_date"))& "/" & Month(app("appeal_date")) & "/" & Year(app("appeal_date"))		
		companyName = app("PEOPLE_COMPANY")
		contactName = app("PEOPLE_NAME")
		productName = app("product_Name")
		quest_id = app("questions_id")
		companyId = app("company_id")
		contactId = app("contact_id")
		appeal_status = app("appeal_status")
		Select Case LCase(trim(appeal_status))
			Case "1" : status_name  = "חדש" : status_num = 1
			Case "2" : status_name  = "בטיפול" : status_num = 2
			Case "3" : status_name  = "סגור"  : status_num = 3	
		End Select
		
		sqlstr =  "Select Langu, Product_Name, PRODUCT_DESCRIPTION, RESPONSIBLE From Products WHERE PRODUCT_ID = " & quest_id
		'Response.Write sqlstr
		'Response.End
		set rsq = con.getRecordSet(sqlstr)
		If not rsq.eof Then		
			Langu = trim(rsq(0))
			prodName = trim(rsq(1))
			PRODUCT_DESCRIPTION = trim(rsq(2))
			RESPONSIBLE = trim(rsq(3))
		End If
		set rsq = Nothing	
		if Langu = "eng" then
			td_align = "left"
			pr_language = "eng"
		else
			td_align = "right"
			pr_language = "heb"
		end if
		
		If isNumeric(RESPONSIBLE) And trim(RESPONSIBLE) <> "" Then
			sqlstr = "Select EMAIL From USERS Where USER_ID = " & RESPONSIBLE
			'Response.Write sqlstr
			'Response.End
			set rswrk = con.getRecordSet(sqlstr)
			If not rswrk.eof Then
			  toMail = trim(rswrk("EMAIL"))
			  fromEmail = trim(rswrk("EMAIL"))
			End If
			set rswrk = Nothing				
			
			xmlFilePath = "../../download/xml_forms/bizpower_forms.xml"		
			'----start adding to XML file ------
			set objDOM = Server.CreateObject("Microsoft.XMLDOM")
			objDom.async = false
		
			if objDOM.load(server.MapPath(xmlFilePath)) then
				Set objRoot = objDOM.documentElement
		
				Set objChield = objDOM.createElement("FORM")		
 				set objNewField=objRoot.appendChild(objChield) 		  			
		  
				Set objNewAttribute = objDOM.createAttribute("APPEAL_ID")
				objNewAttribute.text = appid		
 				objNewField.attributes.setNamedItem(objNewAttribute) 
				Set objNewAttribute=nothing
				
				Set objNewAttribute = objDOM.createAttribute("TYPE_ID")
				objNewAttribute.text = "1"
 				objNewField.attributes.setNamedItem(objNewAttribute) 
				Set objNewAttribute=nothing
				
				Set objNewAttribute = objDOM.createAttribute("PRODUCT_NAME")
				objNewAttribute.text = PRODUCT_NAME
 				objNewField.attributes.setNamedItem(objNewAttribute)  	
 				Set objNewAttribute=nothing
 				
 				Set objNewAttribute = objDOM.createAttribute("RESPONSIBLE")
				objNewAttribute.text = RESPONSIBLE
 				objNewField.attributes.setNamedItem(objNewAttribute)  	
 				Set objNewAttribute=nothing				
 		  
				objDom.save server.MapPath(xmlFilePath)
				    
				set objNewField=nothing
		Set objChield = nothing
	end if ' objDOM.load	
	set objDOM = nothing
	'----end adding to XML file ------		  
			
	If trim(Langu) = "heb" Then		
		
	strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & vbCrLf &_
	"<link href=" & strLocal & "netcom/IE4.css rel=STYLESHEET type=text/css></head><body>" & vbCrLf  &_			
	"<table border=0 width=450 cellspacing=0 cellpadding=0 align=center bgcolor=#e6e6e6><tr>" & vbCrLf  &_	
    "<td class=title_form width=100% style=""font-size:14pt;line-height:26px"" align=center>" & PRODUCT_NAME &_
    "</td></tr><tr>" & vbCrLf &_	
	"<td width=100% align=right valign=top  bgcolor=#C9C9C9 style=""border-bottom: 1px solid #808080;border-top: 1px solid #808080;padding-right:15px"">" & vbCrLf  &_
	"<TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=0><tr>" & vbCrLf &_
	"<td align=right width=100% ><span style=""width:70;"" class=""Form_R"" dir=""rtl"">" & appeal_date & "</span></td>" & vbCrLf &_
	"<td align=right width=130 nowrap class=""form_title"">תאריך הזנה</td></tr>" & vbCrLf &_
	"<tr><td align=right width=100% ><span class=status_num" & status_num &" style=""text-align:center"">" & status_name & "</span></td>" & vbCrLf &_
	"<td align=right width=130 nowrap class=""form_title"">סטטוס</td></tr>"
	If isNumeric(companyId) Then
		strBody = strBody & "<tr><td align=right width=100% >" & companyName & "</td>" & vbCrLf &_
		"<td align=right width=130 nowrap class=""form_title"">קישור ל" & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & "</td></tr>"	& vbCrLf	
	End If
	If isNumeric(contactId) Then
		strBody = strBody & "<tr><td align=right width=100% >" & contactName & "</td>" & vbCrLf &_
		"<td align=right width=130 nowrap class=""form_title"">קישור ל" & trim(Request.Cookies("bizpegasus")("ContactsOne")) & "</td></tr>" & vbCrLf
	End If
	If isNumeric(projectID) And trim(projectName) <> "" Then
		strBody = strBody & "<tr><td align=right width=100% >" & projectName & "</td>" & vbCrLf &_
		"<td align=right width=130 nowrap class=""form_title"">קישור ל" & trim(Request.Cookies("bizpegasus")("Projectone")) & "</td></tr>" & vbCrLf	
	End If
	strBody = strBody & "<tr><td height=10 colspan=2 nowrap></td></tr></table></td></tr>" & vbCrLf
	If trim(PRODUCT_DESCRIPTION) <> "" Then
	strBody = strBody & "<tr><td align=" & td_align & " width=100% bgcolor=#E6E6E6 colspan=2>" & vbCrLf &_
	"<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=3 width=100% ><TR>" & vbCrLf &_
	"<td class=""form_makdim"" dir=""rtl"" width=100% align=" & td_align & ">" & breaks(PRODUCT_DESCRIPTION) & "</td></tr></TABLE></td></tr>" & vbCrLf		
    End If			
    strBody = strBody & "<tr><td width=100% colspan=2 align=right><TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=3 bgcolor=White>" & vbCrLf
	sqlstr = "EXECUTE get_field_value '" & OrgID & "','','"& appId & "','" & quest_id & "',''"
	set fields=con.GetRecordSet(sqlstr)
	If Not fields.eof Then
		arr_fields = fields.getRows()
	End If
	set fields=nothing	
	If isArray(arr_fields) Then
		For ff=0 To Ubound(arr_fields,2)
			Field_Align = trim(arr_fields(0,ff))
			Field_Title = trim(breaks(arr_fields(5,ff)))
			Field_Type = trim(arr_fields(6,ff))
			Field_Value = trim(arr_fields(8,ff))		
				
			If pr_language = "eng" Then
				strBody = strBody & "<tr><td dir=rtl "
				If Field_Type <> "10" Then  'question
					strBody = strBody & " class=""form"""
				Else
					strBody = strBody & " class=""title_form"""
				End If
				strBody = strBody & " width=""100%"" align=left bgcolor=#DADADA valign=""top"" style=""padding-left:15px""><b> " & Field_Title & "</b></td></tr>"
			ElseIf pr_language = "heb" Then
				strBody = strBody & "<tr><td dir=rtl "
				If Field_Type <> "10" then  'question
					strBody = strBody & " class=""form"""
				Else
					strBody = strBody & " class=""title_form"""
				End If
				strBody = strBody & " width=""100%"" align=right bgcolor=#DADADA valign=""top"" style=""padding-right:15px""><b> " & Field_Title & "</b></td></tr>"
			End If
							
			If Field_Type <> "10" Then  'question
				strBody = strBody & "<tr><td width=""85%"" class=""form"" align=" & td_align & " bgcolor=#f0f0f0 style=""padding-right:15px"""
				If pr_language = "eng" Then
					strBody = strBody & " dir=ltr"
				Else
					strBody = strBody & " dir=rtl"
				End if
				strBody = strBody & " >&nbsp;"
				If Field_Type = "1" Or Field_Type = "2" Or Field_Type = "3" Or  Field_Type = "6" Or Field_Type = "7" Or Field_Type = "8" Or Field_Type = "9" Or Field_Type = "11" Or Field_Type = "12" Then
					strBody = strBody & breaks(Field_Value)				
				Elseif Field_Type = "5" Then  'check
					If Field_Value="on" And Field_Align="rtl" Then
						strBody = strBody & "כן"
					ElseIf Field_Value="on" And Field_Align="ltr" Then
						strBody = strBody & "yes"
					ElseIf Field_Align="rtl" Then
						strBody = strBody & "לא"
					Else
						strBody = strBody & "no"
					End If					
				End If
				strBody = strBody & "&nbsp;</td></tr>" & vbCrLf
			End If
		Next
		End If
		strBody = strBody & "</TABLE></td></tr>" & vbCrLf  &_
		"<tr><td align=center colspan=2 bgcolor=white><br><a href='"&strLocal&"netCom/members/appeals/feedback.asp?appid="&appid&"&UserID="&RESPONSIBLE&"' target=_blank style='FONT-SIZE:16px;COLOR:#0000FF'>לחץ כאן לפרטים</a><br></td></tr>" & vbCrLf  &_
		"</table></td></tr></table>" & vbCrLf
	Else
	    	strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & vbCrLf &_
			"<link href=" & strLocal & "netcom/IE4.css rel=STYLESHEET type=text/css></head><body>" & vbCrLf  &_			
			"<table border=0 width=380 cellspacing=0 cellpadding=0 align=center bgcolor=#e6e6e6 dir=rtl><tr>" & vbCrLf  &_	
			"<td class=title_form width=100% style=""font-size:14pt;line-height:26px"" align=center>" & PRODUCT_NAME &_
			"</td></tr><tr>" & vbCrLf &_	
			"<td width=100% align=left valign=top  bgcolor=#C9C9C9 style=""border-bottom: 1px solid #808080;border-top: 1px solid #808080;padding-left:15px"">" & vbCrLf  &_
			"<TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=0 dir=rtl><tr>" & vbCrLf &_
			"<td align=left width=100% ><span style=""width:70;"" class=""Form_R"" dir=""ltr"">" & appeal_date & "</span></td>" & vbCrLf &_
			"<td align=left width=130 nowrap class=""form_title"">Enter date</td></tr>" & vbCrLf &_
			"<tr><td align=left width=100% ><span class=status_num" & status_num &" style=""text-align:center"">" & status_name & "</span></td>" & vbCrLf &_
			"<td align=left width=130 nowrap class=""form_title"">Status</td></tr>"
			If isNumeric(companyId) Then
				strBody = strBody & "<tr><td align=left width=100% >" & companyName & "</td>" & vbCrLf &_
				"<td align=left width=130 nowrap class=""form_title"">Link to " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & "</td></tr>" & vbCrLf
			End If
			If isNumeric(contactId) Then
				strBody = strBody & "<tr><td align=left width=100% >" & contactName & "</td>" & vbCrLf &_
				"<td align=left width=130 nowrap class=""form_title"">Link to " & trim(Request.Cookies("bizpegasus")("ContactsOne")) & "</td></tr>" & vbCrLf
			End If
			If isNumeric(projectID) And trim(projectName) <> "" Then
				strBody = strBody & "<tr><td align=left width=100% >" & projectName & "</td>" & vbCrLf &_
				"<td align=left width=130 nowrap class=""form_title"">Link to " & trim(Request.Cookies("bizpegasus")("Projectone")) & "</td></tr>" & vbCrLf	
			End If
			strBody = strBody & "<tr><td height=10 colspan=2 nowrap></td></tr></table></td></tr>"
			If trim(PRODUCT_DESCRIPTION) <> "" Then
			strBody = strBody & "<tr><td align=" & td_align & " width=100% bgcolor=#E6E6E6 colspan=2>" & vbCrLf &_
			"<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=3 width=100% ><TR>" & vbCrLf &_
			"<td class=""form_makdim"" dir=""ltr"" width=100% align=" & td_align & ">" & breaks(PRODUCT_DESCRIPTION) & "</td></tr></TABLE></td></tr>" & vbCrLf		
			End If			
			strBody = strBody & "<tr><td width=100% colspan=2 align=left><TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=3 bgcolor=White>" & vbCrLf
			sqlstr = "EXECUTE get_field_value '" & OrgID & "','','"& appId & "','" & quest_id & "',''"
		    set fields=con.GetRecordSet(sqlstr)
			If Not fields.eof Then
				arr_fields = fields.getRows()
			End If
			set fields=nothing		    
			If isArray(arr_fields) Then
				For ff=0 To Ubound(arr_fields,2)
					Field_Align = trim(arr_fields(0,ff))
					Field_Title = trim(breaks(arr_fields(5,ff)))
					Field_Type = trim(arr_fields(6,ff))
					Field_Value = trim(arr_fields(8,ff))
						
					If pr_language = "eng" Then
						strBody = strBody & "<tr><td dir=ltr "
						If Field_Type <> "10" Then  'question
							strBody = strBody & " class=""form"""
						Else
							strBody = strBody & " class=""title_form"""
						End If
						strBody = strBody & " width=""100%"" align=left bgcolor=#DADADA valign=""top"" style=""padding-left:15px""><b> " & Field_Title & "</b></td></tr>"
					ElseIf pr_language = "heb" Then
						strBody = strBody & "<tr><td dir=ltr "
						If Field_Type <> "10" then  'question
							strBody = strBody & " class=""form"""
						Else
							strBody = strBody & " class=""title_form"""
						End If
						strBody = strBody & " width=""100%"" align=left bgcolor=#DADADA valign=""top"" style=""padding-left:15px""><b> " & Field_Title & "</b></td></tr>"
					End If
									
					If Field_Type <> "10" Then  'question
						strBody = strBody & "<tr><td width=""85%"" class=""form"" align=" & td_align & " bgcolor=#f0f0f0 style=""padding-left:15px"""
						strBody = strBody & " dir=ltr"
						strBody = strBody & " >"
						If Field_Type = "1" Or Field_Type = "2" Or Field_Type = "3" Or  Field_Type = "6" Or Field_Type = "7" Or Field_Type = "8" Or Field_Type = "9" Or Field_Type = "11" Or Field_Type = "12" Then
							strBody = strBody & breaks(Field_Value)				
						Elseif Field_Type = "5" Then  'check
							If Field_Value="on" Then
								strBody = strBody & "yes"						
							Else
								strBody = strBody & "no"
							End If	
						End If					
						strBody = strBody & "</td></tr>" & vbCrLf
					End If				
				
			Next
			End If
			strBody = strBody & "</TABLE></td></tr>" & vbCrLf  &_
			"<tr><td align=center colspan=2 bgcolor=white><br><a href='" & strLocal & "netCom/members/appeals/feedback.asp?appid="&appid&"&UserID="&RESPONSIBLE&"' target=_blank style='FONT-SIZE:16px;COLOR:#0000FF'><B>Click here for details</B></a><br></td></tr>" & vbCrLf  &_
			"</table></td></tr></table>" & vbCrLf
	    End If
	
		Dim Msg
		Set Msg = Server.CreateObject("CDO.Message")
			Msg.BodyPart.Charset = "windows-1255"
			Msg.From = "support@pegasusisrael.co.il"
			Msg.MimeFormatted = true
			Msg.Subject = PRODUCT_NAME & " - BIZPOWER"
			Msg.To = toMail		
			Msg.HTMLBody = strBody			
			Msg.HTMLBodyPart.ContentTransferEncoding = "quoted-printable"
			Msg.Send()						
		Set Msg = Nothing
	End If
		set app = nothing
		End If		
	End If	  	
	End If
End If
End If

sizeLogoPr = 0
If isNumeric(trim(quest_id)) Then
		sqlpg="SELECT PRODUCT_LOGO FROM PRODUCTS WHERE PRODUCT_ID="&quest_id&""
		'Response.Write sqlpg
		'Response.End
		set pg=con.getRecordSet(sqlpg)	
		sizeLogoPr=pg.Fields("PRODUCT_LOGO").ActualSize
		set pg = Nothing
End If

sizeLogoOrg = 0
If trim(OrgID) <> "" And sizeLogoPr = 0 Then
	sqlstring = "SELECT ORGANIZATION_LOGO FROM organizations WHERE ORGANIZATION_ID = " & OrgID
	'Response.Write sqlstring
	'Response.End
	set logo=con.GetRecordSet(sqlstring)
	if not logo.EOF  then				
		sizeLogoOrg = logo("ORGANIZATION_LOGO").ActualSize
	else
	    sizeLogoOrg = 0	
	end if
	set logo = Nothing	
End If			
%>
<html>
<head>
<title><%=Product_Name%></title>
<meta charset="windows-1255">
<link href="../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body marginwidth="0" marginheight="0" hspace="0" vspace="0" topmargin="0" leftmargin="0" rightmargin="0">
<table cellpadding="0" cellspacing="0" width="100%" ID="Table1">
<%If sizeLogoPr > 0 OR sizeLogoOrg > 0 Then%>
<tr>
      <td width="100%" valign="bottom" align=center>
      <table border="0" width="90%" cellspacing="1" cellpadding="1" align=center ID="Table2">
       <tr><td height="3"></td></tr>
        <tr>
         <td valign="bottom" align="center" width="100%">     
         <%If sizeLogoPr > 0 Then%>
         <img src="<%=strLocal%>/netcom/GetImage.asp?DB=PRODUCT&amp;FIELD=PRODUCT_LOGO&amp;ID=<%=quest_id%>">
         <%ElseIf sizeLogoOrg > 0 Then%>     
          <img src="<%=strLocal%>/netcom/GetImage.asp?DB=ORGANIZATION&amp;FIELD=ORGANIZATION_LOGO&amp;ID=<%=OrgId%>">          
         <%End If%> 
          </td>                
        </tr>
      </table>
      </td>
    </tr>  
<%End If%>      
<tr><td width="100%">
<tr><td nowrap height="10"></td></tr>
<table WIDTH="90%" ALIGN="center" BORDER="0" CELLPADDING="0" cellspacing="1" id="tb_1" name="tb_1">
<tr><td height=15></td></tr>
<tr>
		<td class="title_form" style="font-size:13pt" align=center width=100%  colspan=4><INPUT type="hidden" id=P name=P value="<%=prodId%>">
			<%if SALE_TYPE = "DISTRIBUTION" then%><%=client_name%> &nbsp; <%end if%><%=Product_Name%>   
		</td>
	</tr>
<%	


If show_sale = "False" And Len(trim(reason)) > 0 Then
If reason = "SALE NOT START" Then%>
<tr><td bgcolor="#FFFFFF" height=15 align=center dir="<%=dir_align%>" style="color:#FFFFFF;background-color:#F1BF3A;font-family:arial"><b>ההרשמה למבצע זה טרם החלה</b></td></tr>
<%ElseIf reason = "SALE END" Then%>
<tr><td bgcolor="#FFFFFF" height=15 align=center dir="<%=dir_align%>" style="color:#FFFFFF;background-color:#F1BF3A;font-family:arial"><b>ההרשמה למבצע זה נסגרה</b></td></tr>
<%ElseIf reason = "CLIENT NOT AUTHORIZ" Then%>
<tr><td bgcolor="#FFFFFF" height=15 align=center dir="<%=dir_align%>" style="color:#FFFFFF;background-color:#F1BF3A;font-family:arial"><b>אינך רשאי להרשם למבצע זה</b></td></tr>
<%ElseIf reason = "ANSWERED" Then%>
<%if Langu = "eng" then%>
<tr><td bgcolor="#FFFFFF" height=15 align=center dir=ltr style="color:#FFFFFF;background-color:#F1BF3A;font-family:arial"><b>You have already provided an answer to this survey.</b><br>Please note that it is not possible to reply to the same survey twice. </td></tr>
<%else%>
<tr><td bgcolor="#FFFFFF" height=15 align=center dir="<%=dir_align%>" style="color:#FFFFFF;background-color:#F1BF3A;font-family:arial"><b>לתשומת לבך, טופס זה כבר מולא בעבר על ידך</b><br>לא ניתן להשתתף פעמיים.</td></tr>
<%End If%>
<%End If
%>	
<tr><td bgcolor="#FFFFFF" height=15></td></tr>
<%Else%>
<tr>											
	<td width="100%" bgcolor="#FFFFFF" height="80">
	<table bgcolor=#E5E5E5  style="BORDER-RIGHT: #4e5960 1px solid;BORDER-TOP: #4e5960 1px solid;BORDER-LEFT: #4e5960 1px solid;BORDER-BOTTOM: #4e5960 1px solid" cellpadding="10" cellspacing="0" width="100%" height="80" ID="Table3">	
	<%if trim(PRODUCT_THANKS) <> "" then%>
	 <tr>
	 <td style="color:#000000;font-family:arial; padding-right:20px; padding-left:20px; font-size: 13pt" align="<%=td_align%>" dir="<%=dir_align%>"><%=breaks(PRODUCT_THANKS)%></td>
	 </tr>
	<%end if%> 
	 <tr><td nowrap height="10"></td></tr>		
</table></td></td></tr>
<%End If%>
<tr><td valign="bottom" width=100% align="right">
	<table height="200" border="0" cellpadding=0 cellspacing=0 width=100% id="tb_2" name="tb_2">
	<!--#include file="../members/products/bottom_inc.asp"-->     
    </table>
	</td></tr>   		
</table>
</table></td></tr>				
</body>
</html>

	
