<%@ Language=VBScript%>
<%Response.Buffer = False%>
<%ScriptTimeout=6000%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%
  UserID = trim(trim(Request.Cookies("bizpegasus")("UserID"))) 
  OrgID = trim(trim(Request.Cookies("bizpegasus")("OrgID"))) 
  query_groupID = trim(Request("send_groupId"))
  prodId = trim(Request("prodId")) 
  pageId = trim(Request("pageId")) 
  blocked_flag = "0"
  PathCalImage = strLocal & "/Netcom/"
  If isNumeric(prodId) = true And isNumeric(pageId) = false Then
  sqlstr =  "Select page_id FROM products where product_id = " & prodId
  set rspage = con.getREcordSet(sqlstr)
  If not rspage.eof Then
	pageId = trim(rspage(0))
  End If
  set rspage = Nothing	
  End If
  currEmailCount = trim(Request.Form("currEmailCount"))
  emailCount = trim(Request.Form("emailCount")) 
%>
<HTML>
<HEAD>
<!-- #include file="../../title_meta_inc.asp" -->
<SCRIPT LANGUAGE=javascript>
<!--	
	send_end = '<%=trim(Request.Form("send_end"))%>';
	if(send_end != "")
	{
		window.onunload = ref
	}
	function ref()
	{
		window.opener.document.location.href="products.asp";
	}	
//-->
</SCRIPT>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<style>
.normalB
{
    FONT-WEIGHT: bold;
    FONT-SIZE: 11pt;
    COLOR: #2e2380;
    FONT-FAMILY: Arial
}
</style>
</HEAD>
<BODY onload="window.focus();">
<%  
  If Request("send_groupId") <> nil Then
		sqlStr = "Select Count(DISTINCT PEOPLE_ID) "&_ 
		" FROM PEOPLES where ORGANIZATION_ID="& OrgID &_
		" AND GROUPID IN ("& Request("send_groupId") & ")" &_
		" And People_ID Not IN (Select People_ID From PRODUCT_CLIENT WHERE PRODUCT_ID = " & prodID & ")"
		If blocked_flag = "0" Then
			sqlStr = sqlStr & " AND PEOPLE_EMAIL NOT IN (SELECT blocked_email FROM blocked_emails where ORGANIZATION_ID=" & OrgID&")"		
		End If			
		'Response.Write sqlStr
		set rsc = con.getRecordSet(sqlStr)
		If not rsc.eof Then
			totalCount = trim(rsc(0))
		Else
			totalCount = 0	
		End If
		set rsc = Nothing	
  Else
		totalCount =  trim(Request.Form("totalCount"))
  End If
  
  If Request("send_groupId") <> nil Then
		sqlStr = "Select Count(DISTINCT PEOPLE_ID) "&_ 
		" FROM PEOPLES where ORGANIZATION_ID="& OrgID &_ 
		" AND GROUPID IN ("& Request("send_groupId") & ")" &_
		" And People_ID Not IN (Select People_ID From PRODUCT_CLIENT WHERE PRODUCT_ID = " & prodID & ")"
		If blocked_flag = "0" Then
			sqlStr = sqlStr & " AND PEOPLE_EMAIL NOT IN (SELECT blocked_email FROM blocked_emails where ORGANIZATION_ID=" & OrgID&")"		
		End If		
		'Response.Write sqlStr
		set rsc = con.getRecordSet(sqlStr)
		If not rsc.eof Then
			totalAllCount = trim(rsc(0))
		Else
			totalAllCount = 0	
		End If
		set rsc = Nothing	
  Else
		totalAllCount = trim(Request.Form("totalAllCount"))
  End If
  
  emailCount = 0
  emailLimit = 0	
  
  sql = "Select Email_Limit From Organizations WHERE ORGANIZATION_ID = " & OrgID
  set rs_org = con.getRecordSet(sql)	
  If not rs_org.eof Then
	emailLimit = trim(rs_org(0))
  End If
  set rs_org = Nothing	
	
  If isNumeric(emailLimit) Then
	emailLimit = cLng(emailLimit)
  End If	     	     
  
%>
<table cellpadding=0 cellspacing=0 width="100%">
<%If Request.Form("send_end") = nil Then ' end of sending%>
<tr><td height="20"></td></tr>	
<tr>
	<td align="center" nowrap class="normalB" dir=rtl><font color="red">תהליך שליחת מיילים מתבצע כעת, התהליך יארך מספר דקות</font></td>
</tr>		
<tr>
	<td  align="center" nowrap class="normalB" dir=rtl><font color="red">אין לסגור חלון זה ואין להפעיל ישומים אחרים במחשב בזמן התהליך.</font></td>
</tr>	
<%End If%>
<tr><td height="30" nowrap></td></tr>
<%If Request.Form("send_end") <> nil Then ' end of sending%>
<tr><td align="center" class="td_subj">שליחת המיילים הסתיימה</td></tr>
<%ElseIf currEmailCount >= emailLimit And Request.Form("send_end") <> nil Then%>			
<tr><td align="center" class="td_subj">כיוון שיתרתך לשליחת מיילים נגמרה שליחת מיילים הסתיימה</td></tr>
<%ElseIf Request.Form("send_end") = nil Then%>
<tr><td align="center" class="td_subj">שליחת המיילים החלה</td></tr>
<%End If%>
<tr><td height="10" nowrap></td></tr>
<%If isNumeric(trim(currEmailCount)) Then%>
<tr><td align="center" dir=rtl class="td_subj"> נשלחו &nbsp; <font color="#023296"><b><%=currEmailCount%></b></font> מיילים</td></tr>
<%End If%>
<tr><td align="center" class="td_subj" dir=rtl> סה"כ מיילים לשליחה <font color="#023296"><b><%=totalAllCount%></b></font></td></tr>
<tr><td height="20" nowrap></td></tr>
<tr><td nowrap align="center">
<INPUT class="but_menu" style="width:90"  type="button" onclick="javascript:window.close()" value="סגור חלון" id=button3 name=button3>
</td></tr>
</table>	
<%'Response.Write "emailCount=" & emailCount & " emailLimit = " & emailLimit & " currEmailCount = " & currEmailCount & " totalAllCount = "  & totalAllCount & " ORGANIZATION_ID = " & OrgID%>
<%'Response.End
If Request.Form("send_end") = nil Then 
  

Function RegExpTest(sEmail)
  RegExpTest = false
  Dim regEx, retVal
  Set regEx = New RegExp

  ' Create regular expression:
  regEx.Pattern ="^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$" 

  ' Set pattern:
  regEx.IgnoreCase = true

  ' Set case sensitivity.
  retVal = regEx.Test(sEmail)

  ' Execute the search test.
  If not retVal Then
    exit function
  End If

  RegExpTest = true
End Function


Function createForm(quest_id,prodId,pageId,PEOPLE_ID,OrgID)  

  poll_message = strLocal &"/netcom/seker/seker.asp?" & Encode("P=" & prodId & "&C=" & PEOPLE_ID & "&O=" & OrgID)
  
  If trim(quest_id) <> "" Then ' יש טופס בדף
	
	set prod = con.GetRecordSet("Select Langu from products where product_id=" & quest_id & " and ORGANIZATION_ID=" & OrgID )
	if not prod.eof then		
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
		If trim(PEOPLE_ID) = "" Then
			form_content = form_content & " href=""#"" OnClick=""return false;""> "
		Else
			form_content = form_content &  " href="""&vFix(poll_message)& """ target=""_blank"">"
		End If
		form_content = form_content & " <img src="""&strLocal&"netcom/GetImage.asp?DB=Page&FIELD=LINK_IMAGE&ID="&pageId&""" border=""0""></a>"
	 
	ElseIf trim(formLink) = "link" Then 'text
		'Response.Write image_name
		'Response.End
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
		If trim(PEOPLE_ID) = "" Then
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
				
		form_content = createFormContent(productName,pertype,perbgcolor,pertitlecolor,persize,percolor,persubjectcolor)	
	End If		
	set prod = nothing
	End If
	createForm = form_content
END Function
 
function createFormContent(productName,pertype,perbgcolor,titlecolor,persize,percolor,persubjectcolor)
		
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
		form_content = form_content & "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=3><TR>"&chr(13)&chr(10)
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
		
		form_content = form_content & PRODUCT_DESCRIPTION
		
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
		
	End If
		
	If Len(attachment_title_quest) > 0 Then
		form_content = form_content & "<TR><td class=""form_makdim"" dir=""rtl"" width=100% align=right>" &_
		"<a style=""color:black;font-weight:bolder;font-size:11pt"" href="&strLocal&"/download/products/"&attachment_quest&" target=_blank>"& attachment_title_quest&"</a></td></tr>"
	End If		
					
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
			form_content = form_content &"  style=""padding-left:10px"" align=left>"
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
			form_content = form_content & "</td></tr>"&chr(13)&chr(10)
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
			form_content = form_content & "</td></tr>"&chr(13)&chr(10)
			end if
		end if
		fields.moveNext()
		loop
	set fields=nothing 
	
	form_content = form_content & "</TABLE></td></tr>"&chr(13)&chr(10)	
	form_content = form_content &"<tr><td colspan=""4"" height=""10"" width=""550"">"&chr(13)&chr(10)
	form_content = form_content &"<input type=""text"" style=""width:0"" name=""C"" value="&PEOPLE_ID&">"&chr(13)&chr(10)
	form_content = form_content &"<input type=""text"" style=""width:0"" name=""P"" value="&prodId&">"&chr(13)&chr(10)
	form_content = form_content &"<input type=""text"" style=""width:0"" name=""O"" value="""&OrgID&""">"&chr(13)&chr(10)
	form_content = form_content &"<table width=""550"" align=""center"" cellpadding=""0"" cellspacing=""0"">"&chr(13)&chr(10)	
	form_content = form_content &"<tr><td width=""550"" bgcolor=""transparent"" height=""1"" colspan=""2""></td>"&chr(13)&chr(10)
	form_content = form_content &"</tr></table></td></tr>"&chr(13)&chr(10)
	if Langu = "eng" then
		form_content = form_content &"<tr><td colspan=""4"" align=center>"
		If trim(PEOPLE_ID) <> "" Then
			form_content = form_content & "<input type=""image"" SRC=""" & strLocal & "/netcom/images/SEND-ENG.gif"" border=""0"" name=""submit_button"" id=""submit_button"" hspace=""60""></td></tr>"&chr(13)&chr(10)
		Else
			form_content = form_content & "<IMG SRC=""" & strLocal & "/netcom/images/SEND-ENG.gif"" border=""0"" OnClick=""return false;"" name=""submit_button"" id=""submit_button"" hspace=""60""></td></tr>"&chr(13)&chr(10)
		End If 		
	Else
		form_content = form_content &"<tr><td colspan=""4"" align=center>"
		If trim(PEOPLE_ID) <> "" Then
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
    form_script = form_script & " return true;	}"&chr(13)&chr(10)&_
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

	sqlStr = "Select product_name,PRODUCT_TYPE,PRODUCT_DESCRIPTION,Langu,From_Mail,FROM_NAME,EMAIL_SUBJECT,"&_
	" Page_ID,QUESTIONS_ID,FILE_ATTACHMENT From PRODUCTS Where product_id=" & prodId &_
	" AND ORGANIZATION_ID=" & OrgID
	'Response.Write sqlStr
	'Response.End
	set rs_product = con.GetRecordSet(sqlStr)
		PRODUCT_DESCRIPTION = rs_product("PRODUCT_DESCRIPTION")
		product_name = rs_product("product_name")
		PRODUCT_TYPE = rs_product("PRODUCT_TYPE")
		fromEmail = trim(rs_product("From_Mail"))
		fromName = trim(rs_product("From_Name"))
		pageID = trim(rs_product("Page_Id"))
		quest_id = trim(rs_product("QUESTIONS_ID"))
		If trim(quest_id) <> "" Then ' יש טופס בדף	
			set prod = con.GetRecordSet("Select * from products where product_id=" & quest_id & " and ORGANIZATION_ID=" & OrgID )
			if not prod.eof then
				productName=prod("Product_Name")					
				PRODUCT_DESCRIPTION = prod("PRODUCT_DESCRIPTION")
				attachment_quest = prod("FILE_ATTACHMENT")
				attachment_title_quest = prod("ATTACHMENT_TITLE")
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
		End If
		set prod = nothing	
		subjectEmail = trim(rs_product("EMAIL_SUBJECT")) 
		FILE_ATTACHMENT = trim(rs_product("FILE_ATTACHMENT")) 		
		if rs_product("Langu") = "eng" then
			dir_align = "ltr"
			Langu = "eng"
		else
			dir_align = "rtl"
			Langu = "heb"			
		end if
	set rs_product=nothing
	
	If Len(trim(fromEmail)) = 0 Then	
	sql = "Select EMAIL From USERS  WHERE USER_ID = " & UserID
	set rs_org = con.getRecordSet(sql)	
	If not rs_org.eof Then
		fromEmail = trim(rs_org(0))
	End If
	set rs_org = Nothing	
End If	
	USER_EMAIL = fromEmail
	If isNumeric(pageId) And trim(pageId) <> "0" Then
		set pg=con.getRecordSet("SELECT Page_Title,Page_Source,Product_ID,FORM_LINK_IMAGE FROM Pages WHERE Page_Id="&pageId&" ")
		If not pg.eof Then
			pertitle=pg("Page_Title")	
			pagesource=trim(pg("Page_Source"))						
			formLink = trim(pg("FORM_LINK_IMAGE"))	
		End If 	 
		set pg = Nothing  
	End if  

	arr_types = Split(Request("send_groupId"),",")
	
	
    image_name = "id=""tofes"""	
	If Len(pagesource) > 0 Then
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
			'length_to_cut = ind_end - ind_start + 1	
			'string_to_replace = Mid(pagesource,ind_start,length_to_cut)
			first_slice = Mid(pagesource,1,ind_start-1)
			second_slice = Mid(pagesource,ind_end+1)
			'Response.Write form_content 
			'Response.End	
			'pagesource = Replace(pagesource,string_to_replace,form_content)
		End If	
	End If		
				
	FOR EACH send_groupId in arr_types				
		sqlStr = "Select Top 100 PEOPLE_ID, PEOPLE_NAME, PEOPLE_PHONE,"&_
		" PEOPLE_CELL, PEOPLE_FAX, PEOPLE_EMAIL, PEOPLE_COMPANY, PEOPLE_OFFICE,"&_
		" GROUPID, CONTACT_ID, COMPANY_ID  FROM PEOPLES WHERE ORGANIZATION_ID="& OrgID &_		
		" AND GROUPID IN ("& Request("send_groupId") &")"&_
		" AND PEOPLE_ID NOT IN "&_ 
		" (SELECT PEOPLE_ID FROM PRODUCT_CLIENT where ORGANIZATION_ID=" & OrgID &_
		" AND PRODUCT_ID ="& prodId &")"
		If blocked_flag = "0" Then
			sqlStr = sqlStr & " AND PEOPLE_EMAIL NOT IN (SELECT blocked_email FROM blocked_emails where ORGANIZATION_ID=" & OrgID&")"		
		End If
		'Response.Write sqlStr
		'Response.End	
		set people = con.GetRecordSet(sqlStr)
		Do while not people.eof
		    poll_message = ""
		    PEOPLE_ID = trim(people("PEOPLE_ID"))
		    peopleName=trim(people("PEOPLE_NAME"))				
			peoplePHONE=trim(people("PEOPLE_PHONE"))
			peopleFAX=trim(people("PEOPLE_FAX"))
			peopleCELL = trim(people("PEOPLE_CELL"))
			peopleEMAIL=trim(people("PEOPLE_EMAIL"))
			peopleCOMPANY=trim(people("PEOPLE_COMPANY"))
			peopleOFFICE=vFix(trim(people("PEOPLE_OFFICE")))			
			GROUPID = trim(people("GROUPID"))
			contactID = trim(people("CONTACT_ID"))
			companyID = trim(people("COMPANY_ID"))
			If Len(contactID) = 0 Or IsNull(contactID) Then
				contactID = "NULL"
			End If			
			If Len(companyID) = 0 Or IsNull(companyID) Then
				companyID = "NULL"
			End If			
			
			'//start of sending Email			
			url = strLocal & "/sale.asp?" & Encode("prodId=" & prodId & "&pageId=" & pageId & "&C=" & PEOPLE_ID & "&O=" & OrgID)
			Body_message = ""
			Body_message = Body_message &"<HTML><HEAD><META http-equiv=Content-Type content=""text/html; charset=windows-1255"">"&chr(13)&chr(10)									
			Body_message = Body_message &"<title>"&pertitle&"</title>"&chr(13)&chr(10)
			Body_message = Body_message &"<style>Body,P,TD {Font-Family:Arial; Font-Size:10pt}"&chr(13)&chr(10)
			Body_message = Body_message &"A.bottom { font-family:Arial (Hebrew); font-size:8pt; font-weight:500; color:#6A6A6A; vertical-align:top;}"&chr(13)&chr(10)
			Body_message = Body_message &"TD.bottom { font-family:Arial (Hebrew); font-size:8pt; font-weight:500; color:#6A6A6A; vertical-align:top;}</style></head>"&chr(13)&chr(10)			
			Body_message = Body_message &"<body topmargin=5 leftmargin=0 rightmargin=0>"&chr(13)&chr(10)
			Body_message = Body_message &"<center><P align=""center"" width=""100%"" style=""font-size:10pt;font-weight:600;font-family:Arial"">If you can't see this letter properly, please <A style=""color:blue;font-weight:600;text-decoration:underline;font-size=10pt"" href="""&url&""" target=""_blank"">click here</A> to see it on the web</P></center>"&chr(13)&chr(10)
			Body_message = Body_message &"<table cellpadding=0 cellspacing=0 width=620 align=center border=0><tr><td>"
            form_content = ""
            If trim(quest_id) <> "" Or IsNull(quest_id) = false Then
				form_content = createForm(quest_id,trim(Request("prodId")),pageId,PEOPLE_ID,OrgID)
			End if		
			
			If Len(first_slice) > 0 Or Len(form_content) > 0 Or Len(second_slice) > 0 Then
				pagesource = first_slice & form_content & second_slice			
			End If
			
			Body_message = Body_message & pagesource
			Body_message = Body_message & "</td></tr><tr><td align=""center"" bgcolor=""#FFFFFF""><table cellpadding=1 cellspacing=0 width=620 border=0 align=center>"&_
			"<tr><td height=5 nowrap></td></tr>"&chr(13)&chr(10)&_
			"<tr><td align=right class=""bottom"" width=100% dir=ltr><A class=""bottom"" href=""http://pegasus.bizpower.co.il"" target=""_blank"" dir=ltr><img src="""&strLocal&"/images/"&prodId&"_"&PEOPLE_ID&".aspx"" border=0 hspace=5></A></td>"&_
			"<td align=right valign=""top"" nowrap class=""bottom"" dir=rtl>דואר זה נשלח אליך מ"&trim(Request.Cookies("bizpegasus")("ORGNAME"))&_
			",&nbsp;במידה ואינך מעוניין לקבל דיוור מ"&trim(Request.Cookies("bizpegasus")("ORGNAME"))&_
			" <a class=""bottom"" href=""" & strLocal & "/remove_email.asp?" & Encode("OrgID="&OrgID&"&PEOPLE_ID="&PEOPLE_ID) &""" target=""_new"">לחץ כאן</a>&nbsp;</td></tr>"&chr(13)&chr(10)&_
			"<tr><td height=10 nowrap></td></tr></table>"&chr(13)&chr(10)&_
			"</td></tr></table>"
		
			USER_EMAIL = fromEmail
			'peopleEMAIL = "olga@eltam.com"
			'Response.Write Body_message
			'Response.End
						
			If RegExpTest(peopleEMAIL) Then		
			Dim Msg
			Set Msg = Server.CreateObject("CDO.Message")
				Msg.BodyPart.Charset = "windows-1255"
				If Len(fromName) > 0 Then
					Msg.From = trim(fromName) & " <" & trim(USER_EMAIL) & ">"
				Else
					Msg.From = USER_EMAIL
				End If				
				Msg.MimeFormatted = true
				Msg.Subject = subjectEmail
				Msg.To = peopleEMAIL				
				Msg.HTMLBody = Body_message								
				If Len(trim(FILE_ATTACHMENT)) > 0 Then				
				    Set fs = CreateObject("Scripting.FileSystemObject") 
					if fs.fileexists(Server.MapPath("../../../download/products/") & "/" & FILE_ATTACHMENT) then 
						Msg.AddAttachment Server.MapPath("../../../download/products/") & "/" & FILE_ATTACHMENT
					else 
					
					end if
					set fs = Nothing											
				End If
				Msg.HTMLBodyPart.ContentTransferEncoding = "quoted-printable"				
				Msg.Send()						
			Set Msg = Nothing								 
     
			sqlstr="INSERT INTO PRODUCT_CLIENT (PRODUCT_ID, PEOPLE_ID, PEOPLE_NAME, "&_
			" PEOPLE_PHONE, PEOPLE_FAX, PEOPLE_CELL, PEOPLE_EMAIL, PEOPLE_OFFICE, "&_
			" PEOPLE_COMPANY, GROUPID, User_ID, ORGANIZATION_ID, CONTACT_ID, COMPANY_ID, DATE_SEND) "&_
			" VALUES ("& prodId &","& PEOPLE_ID &",'"&sFix(peopleName) &"','"&_
			sFix(peoplePHONE)& "','" & sFix(peopleFAX) & "','"&_
			sFix(peopleCELL) & "','" & sFix(peopleEMAIL)&"','"&_
			sFix(peopleOFFICE)&"','"&sFix(peopleCOMPANY)&"',"&_ 
			GROUPID & "," &  UserID & "," & OrgID & "," & contactID & "," & companyID & ", getDate())"	
			con.ExecuteQuery(sqlstr)			
			emailCount = emailCount + 1	
		Else	
				sqlstr="INSERT INTO PRODUCT_CLIENT (PRODUCT_ID, PEOPLE_ID, PEOPLE_NAME, "&_
				" PEOPLE_PHONE, PEOPLE_FAX, PEOPLE_CELL, PEOPLE_EMAIL, PEOPLE_OFFICE, "&_
				" PEOPLE_COMPANY, GROUPID, User_ID, ORGANIZATION_ID, DATE_SEND) "&_
				" VALUES ("& prodId &","& PEOPLE_ID &",'"&sFix(peopleName) &"','"&_
				sFix(peoplePHONE)& "','" & sFix(peopleFAX) & "','"&_
				sFix(peopleCELL) & "','" & sFix(peopleEMAIL)&"','"&_
				sFix(peopleOFFICE)&"','"&sFix(peopleCOMPANY)&"',"&_ 
				GROUPID & "," &  UserID & "," & OrgID & ", NULL)"
				'Response.Write sqlstr
				'Response.End					
				con.ExecuteQuery(sqlstr)					
		End If		
								
		If isNumeric(emailLimit) Then
			If emailCount >= emailLimit Then							
				Exit Do
			End If
		End If	
		
		people.movenext
		loop
		set people = nothing
		
		sqlstr = "Select * From PRODUCT_GROUPS WHERE PRODUCT_ID = "& prodId &_
		" And groupID = " & send_groupId & " And ORGANIZATION_ID = " & OrgID
		set rs_check = con.getRecordSet(sqlstr)
		If rs_check.eof Then
			sqlStr = "INSERT INTO PRODUCT_GROUPS (PRODUCT_ID,groupID,USER_ID,Organization_ID,DATE_SEND) "&_ 
					 " VALUES ("& prodId &","& send_groupId &","& UserID & "," & OrgID &", getDate())"
			'Response.Write sqlStr
			'Response.End			
			con.ExecuteQuery(sqlstr)
		End If
		set rs_check = Nothing
			
		If isNumeric(emailLimit) Then
			If emailCount >= emailLimit Then							 			 
				Exit For
			End If
		End If
				
	NEXT
	
	If isNumeric(emailLimit) And isNumeric(emailCount) Then
	emailRemainder  = emailLimit - emailCount
	
	con.executeQuery("Update ORGANIZATIONS SET Email_Limit = " & emailRemainder & " WHERE ORGANIZATION_ID = " & OrgID)
	
	End If
	
	sqlStr = "Select Count(*) FROM PRODUCT_CLIENT WHERE ORGANIZATION_ID="& OrgID &_		
	" AND GROUPID IN ("& Request("send_groupId") &")"&_
	" AND PRODUCT_ID ="& prodId
	'Response.Write sqlStr
	'Response.End
	set rsc = con.getRecordSet(sqlStr)
	If not rsc.eof Then
		currEmailCount = trim(rsc(0))
	Else
		currEmailCount = 0	
	End If
	set rsc = Nothing
	
	If totalCount > 0 Then
	%>
	<form name="form1" id="form1" action="sendMail.asp" method="post">
	<input type="hidden" name="send_groupId" id="send_groupId" value="<%=Request("send_groupId")%>">
	<input type="hidden" name="prodId" id="prodId" value="<%=prodId%>">	
	<input type="hidden" name="currEmailCount" id="currEmailCount" value="<%=currEmailCount%>">
	<input type="hidden" name="emailCount" id="emailCount" value="<%=emailCount%>">	
	<input type="hidden" name="totalCount" id="totalCount" value="<%=totalCount%>">	
	<input type="hidden" name="totalAllCount" id="totalAllCount" value="<%=totalAllCount%>">	
	<input type="hidden" name="blocked_flag" id="blocked_flag" value="<%=blocked_flag%>">	
	</form>
	<%
		'Response.Redirect "sendMail.asp?send_groupId=" & Request("send_groupId") & "&prodId=" & prodId
	Else	
	%>
	 <form name="form1" id="form1" action="sendMail.asp" method="post">
	 <input type="hidden" name="prodId" id="prodId" value="<%=prodId%>">
	 <input type="hidden" name="currEmailCount" id="currEmailCount" value="<%=currEmailCount%>">
	 <input type="hidden" name="emailCount" id="emailCount" value="<%=emailCount%>">	
	 <input type="hidden" name="totalCount" id="totalCount" value="<%=totalCount%>">	
	 <input type="hidden" name="totalAllCount" id="totalAllCount" value="<%=totalAllCount%>">	
	 <input type="hidden" name="blocked_flag" id="blocked_flag" value="<%=blocked_flag%>">	
	 <input type="hidden" name="send_end" id="send_end" value="1">
	 </form>
	<%
		'Response.Redirect "sendMail.asp?prodId=" & prodId & "&send_end=1"
	End If
	If Request.Form("send_end") = nil Then%>	
	<script>
	function fun() {document.form1.submit();}
	setTimeout("fun()",2)
    </script>
	<%End If
	'//end of send email	
End If
%>	
</BODY>
</HTML>
<%Set con=Nothing%>