
<!--#include file="../../connect.asp"-->
<%Response.Buffer = False%>
<!--#include file="../../reverse.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</HEAD>
<BODY style="margin:0px;SCROLLBAR-FACE-COLOR: #E6E6E6;SCROLLBAR-HIGHLIGHT-COLOR: #F7F7F7;SCROLLBAR-SHADOW-COLOR: #848484;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #808080;SCROLLBAR-TRACK-COLOR: #E6E6E6;SCROLLBAR-DARKSHADOW-COLOR: #ffffff" onload='parent.document.all("CONTACT_NAME").focus()'>
<%	
    UserId = trim(Request.Cookies("bizpegasus")("UserId"))
    OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))
	lang_id	 = trim(Request.Cookies("bizpegasus")("LANGID"))
		
	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If
	If lang_id = 2 Then
		dir_var = "rtl"  :  	align_var = "left"  :  dir_obj_var = "ltr"
	Else 
		dir_var = "ltr"  :  	align_var = "right"   :   dir_obj_var = "rtl"
	End If	
	
	If Request("cont_name") <> nil Then
	
      sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 3 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(arr_title(0)) = arr_title(1)
			End If
		Next
	  End If
	  set rstitle = Nothing
    
    quest_id = trim(Request.QueryString("quest_id")) 
	if trim(Request("cont_name")) <> "" Or trim(Request.QueryString("cont_name")) <> "" then
		cont_name = trim(Request("cont_name"))		
		where_cont_name = " and UPPER(ltrim(rtrim(CONTACT_NAME))) LIKE '"& UCase(sFix(cont_name)) &"%'"
	else
		where_cont_name = ""
	end if	
	
	urlSort="contacts.asp?cont_name="&Server.URLEncode(cont_name)&"&quest_id=" & quest_id
    sort = Request.QueryString("sort")	
	if trim(sort)="" then  sort = 0  end if
	dim sortby(8)	
	sortby(0) = "rtrim(ltrim(CONTACT_NAME))"
	sortby(1) = "rtrim(ltrim(company_name))"
	sortby(2) = "rtrim(ltrim(company_name)) DESC"
	sortby(3) = "date_update"
	sortby(4) = "date_update DESC"
	sortby(5) = "CONTACT_NAME"
	sortby(6) = "CONTACT_NAME DESC"
	
	
%>
<table width="100%" align=center cellpadding=0 cellspacing=0 border=0 bgcolor=#E6E6E6>
<tr>
    <td align="left" width="100%" valign=top dir="<%=dir_var%>" >
    <table width="100%" cellspacing="1" cellpadding="1" bgcolor=#FFFFFF>       
    <tr> 	      
	    <td width="75"  nowrap class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<span id="word6" name=word6><!--טלפון ישיר--><%=arrTitles(6)%></span></td>
	    <td width="85"  nowrap align="<%=align_var%>" class="title_sort" dir="<%=dir_obj_var%>">&nbsp;<span id="word7" name=word7><!--תפקיד--><%=arrTitles(7)%></span>&nbsp;</td>
        <td width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="1" OR trim(sort)="2" then%>_act<%end if%>"><%if trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word17 title="ל<%=arrTitles(17)%>"><%elseif trim(sort)="2" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word18 title="<%=arrTitles(18)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=1" name=word18 title="<%=arrTitles(18)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>bot<%elseif trim(sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	     	      
        <td width="120"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="5" OR trim(sort)="6" then%>_act<%end if%>"><%if trim(sort)="5" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word17 title="<%=arrTitles(17)%>"><%elseif trim(sort)="6" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word18 title="<%=arrTitles(18)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=5" name=word18 title="<%=arrTitles(18)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="5" then%>bot<%elseif trim(sort)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	</tr>   	
	<%  sqlSelect="SELECT contacts.contact_id,contacts.contact_name,contacts.company_ID,company_name," &_
		" contacts.phone,contacts.cellular,contacts.messanger_name,contacts.email FROM contacts "&_
		" Inner Join companies On contacts.company_id = companies.company_id WHERE contacts.organization_id = " &_
		trim(OrgID) & where_cont_name & " order by "& sortby(sort)
		'Response.Write sqlSelect
		'Response.End
	   	set contactList = con.getRecordSet(sqlSelect)
	   	If not contactList.eof Then
		While not contactList.eof
			contactID=contactList(0)
			contactName=rtrim(ltrim(contactList(1)))
			companyID=contactList(2)
			companyName=rtrim(ltrim(contactList(3)))
			phone=contactList(4)
			cellular=contactList(5)
			messanger_name=contactList(6)
			email=contactList(7)	
	%>
       <tr>    	      	     
	      <td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" href="appeal.asp?quest_id=<%=quest_id%>&companyId=<%=companyId%>&contactId=<%=contactId%>" target=_parent style="font-size:11px">&nbsp;<%=phone%>&nbsp;</a></td>
	      <td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" href="appeal.asp?quest_id=<%=quest_id%>&companyId=<%=companyId%>&contactId=<%=contactId%>" target=_parent>&nbsp;<%=messanger_name%>&nbsp;</a></td>
          <td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" href="appeal.asp?quest_id=<%=quest_id%>&companyId=<%=companyId%>&contactId=<%=contactId%>" style="line-height:120%;padding-top:3px;padding-bottom:3px" target=_parent>&nbsp;<%=companyName%>&nbsp;</a></td>	        
          <td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" href="appeal.asp?quest_id=<%=quest_id%>&companyId=<%=companyId%>&contactId=<%=contactId%>" target=_parent>&nbsp;<%=contactName%>&nbsp;</a></td>
	   </tr> 
	<%
		contactList.moveNext
		Wend
		set contactList = Nothing
		Else%>
		<tr>
	   <td colspan="8" align=center class="title_sort1" dir="rtl"><span id="word9" name=word9><!--לא נמצאו--><%=arrTitles(9)%></span>&nbsp;</td>
	   </tr>
	<%	End If	%>	
</td></tr></table>
<%	End If	%>
<%set con =nothing%>
</BODY>
</HTML>
