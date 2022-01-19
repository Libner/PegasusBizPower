
<!--#include file="../../connect.asp"--><%Response.Buffer = False%>
<!--#include file="../../reverse.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</HEAD>
<BODY style="margin:0px;SCROLLBAR-FACE-COLOR: #E6E6E6;SCROLLBAR-HIGHLIGHT-COLOR: #F7F7F7;SCROLLBAR-SHADOW-COLOR: #848484;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #808080;SCROLLBAR-TRACK-COLOR: #E6E6E6;SCROLLBAR-DARKSHADOW-COLOR: #ffffff" onunload='parent.document.all("company_name").focus()'>
<%	
    UserId = trim(Request.Cookies("bizpegasus")("UserId"))
    OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))
	lang_id	 = trim(Request.Cookies("bizpegasus")("LANGID"))
		
	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If
	If lang_id = 2 Then
		dir_var = "rtl"
		align_var = "left"
		dir_obj_var = "ltr"
	Else
		dir_var = "ltr"
		align_var = "right"
		dir_obj_var = "rtl"
	End If	
	
	If Request("comp_name") <> nil Then
	
      sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 4 Order By word_id"				
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
	if trim(Request("comp_name")) <> "" Or trim(Request.QueryString("comp_name")) <> "" then
		comp_name = trim(Request("comp_name"))		
		where_comp_name = " and UPPER(LTRIM(RTRIM(companies.company_name))) LIKE '"& UCase(sFix(comp_name)) &"%'"
	else
		where_comp_name = ""
	end if	
	
	urlSort="companies.asp?comp_name="&Server.URLEncode(comp_name)&"&quest_id=" & quest_id
    sort = Request.QueryString("sort")	
	if trim(sort)="" then  sort = 0  end if
	dim sortby(8)	
	sortby(0) = "rtrim(ltrim(company_name))"
	sortby(1) = "rtrim(ltrim(company_name))"
	sortby(2) = "rtrim(ltrim(company_name)) DESC"
	sortby(3) = "rtrim(ltrim(city_Name))"
	sortby(4) = "rtrim(ltrim(city_Name)) DESC"
	sortby(5) = "CONTACT_NAME"
	sortby(6) = "CONTACT_NAME DESC"
	sortby(7) = "type_name"
	sortby(8) = "type_name DESC"
	
	If Request.QueryString("companyId") <> nil Then
	    companyId = trim(Request.QueryString("companyId"))
		sqlstr = "Select COMPANY_ID, COMPANY_NAME, city_Name,address, "&_		
		" phone1,phone2,fax1,email,url,private_flag,company_desc "&_
		" From companies Where ORGANIZATION_ID = " & trim(OrgID) &_
		" AND Company_ID = " & 	companyId
		'Response.Write sqlstr
		'Response.End
		set rs_comp=con.GetRecordSet(sqlstr)    
		If not rs_comp.eof Then 				
			company_id = trim(rs_comp(0))
			If IsNull(rs_comp(1)) Then
				company_name = ""
			Else
				company_name = trim(rs_comp(1))					
			End If	
			If IsNull(rs_comp(2)) Then
				cityName = ""
			Else			
				cityName = trim(rs_comp(2))	
			End If
			If IsNull(rs_comp(3)) Then
				address = ""
			Else				
				address = trim(rs_comp(3))			
			End If
			If IsNull(rs_comp(4)) Then
				phone1 = ""
			Else				
				phone1 = trim(rs_comp(4)) 				
			End If
			If IsNull(rs_comp(5)) Then
				phone2 = ""
			Else				
				phone2 = trim(rs_comp(5)) 			
			End If
			If IsNull(rs_comp(6)) Then
				fax1 = ""
			Else				
				fax1 = trim(rs_comp(6)) 
			End If
			If IsNull(rs_comp(7)) Then
				email = ""
			Else				
				email = trim(rs_comp(7)) 
			End If
			If IsNull(rs_comp(8)) Then
				url = ""
			Else				
				url = trim(rs_comp(8)) 
			End If
			If IsNull(rs_comp(9)) Then
				private_flag = ""
			Else				
				private_flag = trim(rs_comp(9))
			End If
			If IsNull(rs_comp(10)) Then
				company_desc = ""
			Else				
				company_desc = trim(rs_comp(10))
			End If	
  		End If
  		set rs_comp = Nothing    	
  	%>
  	<INPUT type="hidden" id="COMPANY_ID" name="COMPANY_ID" value="<%=trim(vFix(COMPANY_ID))%>">		
	<INPUT type="hidden" id="COMPANY_NAME" name="COMPANY_NAME" value="<%=vFix(COMPANY_NAME)%>">		
	<INPUT type="hidden" id="types" name="types" value="<%=vFix(types)%>">		
	<INPUT type="hidden" id="cityName" name="cityName" value="<%=vFix(cityName)%>">
	<INPUT type="hidden" id="address" name="address" value="<%=vFix(address)%>">				
	<INPUT type="hidden" id="prefix_phone1" name="prefix_phone1" value="<%=vFix(prefix_phone1)%>">
	<INPUT type="hidden" id="phone1" name="phone1" value="<%=vFix(phone1)%>">
	<INPUT type="hidden" id="prefix_phone2" name="prefix_phone2" value="<%=vFix(prefix_phone2)%>">
	<INPUT type="hidden" id="phone2" name="phone2" value="<%=vFix(phone2)%>">
	<INPUT type="hidden" id="prefix_fax1" name="prefix_fax1" value="<%=vFix(prefix_fax1)%>">
	<INPUT type="hidden" id="fax1" name="fax1" value="<%=vFix(fax1)%>">
	<INPUT type="hidden" id="email" name="email" value="<%=vFix(email)%>">
	<INPUT type="hidden" id="url" name="url" value="<%=vFix(url)%>">
	<INPUT type="hidden" id="company_desc" name="company_desc" value="<%=vFix(company_desc)%>">
	<INPUT type="hidden" id="private_f" name="private_f" value="<%=vFix(private_flag)%>">		

	<SCRIPT LANGUAGE=javascript>
	<!--
			
		private_flag = window.document.all("private_f").value;
		if(window.parent.document.all("companyId") != null)
		{
			window.parent.document.all("companyId").value = window.document.all("company_id").value;		
		}	
			
		if(window.parent.document.all("COMPANY_NAME") != null)	
		{
			window.parent.document.all("COMPANY_NAME").value = window.document.all("company_name").value;
			window.parent.document.all("COMPANY_NAME").readOnly = true;
			window.parent.document.all("COMPANY_NAME").className = "Form_R";				
			
			window.parent.document.all("delete_company").style.display = "inline";
			window.parent.document.all("delete_company").style.visible = true;		
					
		}	
			
		if(window.parent.document.all("cityName") != null)
		{		
			window.parent.document.all("cityName").value = window.document.all("cityName").value;		
			if(private_flag == "1")
			{
				window.parent.document.all("cityName").readOnly = true;
				window.parent.document.all("cityName").className = "Form_R";
			}
			else
			{
				window.parent.document.all("cityName").readOnly = false;
				window.parent.document.all("cityName").className = "Form";
			}	
		}	
			
		if(window.parent.document.all("address") != null)
		{
			window.parent.document.all("address").value = window.document.all("address").value;	
			if(private_flag == "1")
			{		
				window.parent.document.all("address").readOnly = true;
				window.parent.document.all("address").className = "Form_R";
			}
			else
			{
				window.parent.document.all("address").readOnly = false;
				window.parent.document.all("address").className = "Form";
			}
			window.parent.document.all("address").focus();	
		}			


		if(window.parent.document.all("phone1") != null)
		{
			window.parent.document.all("phone1").value = window.document.all("phone1").value;			
			if(private_flag == "1")
			{				
				window.parent.document.all("phone1").readOnly = true;
				window.parent.document.all("phone1").className = "Form_R";
			}
			else
			{
				window.parent.document.all("phone1").readOnly = false;
				window.parent.document.all("phone1").className = "Form";
			}	
		}			

			
		if(window.parent.document.all("phone2") != null)
		{
			window.parent.document.all("phone2").value = window.document.all("phone2").value;				
			if(private_flag == "1")
			{						
				window.parent.document.all("phone2").readOnly = true;
				window.parent.document.all("phone2").className = "Form_R";
			}
			else
			{
				window.parent.document.all("phone2").readOnly = false;
				window.parent.document.all("phone2").className = "Form";
			}	
		}
			
		if(window.parent.document.all("fax1") != null)
		{
			window.parent.document.all("fax1").value = window.document.all("fax1").value;
			if(private_flag == "1")
			{											
				window.parent.document.all("fax1").readOnly = true;
				window.parent.document.all("fax1").className = "Form_R";
			}
			else
			{
				window.parent.document.all("fax1").readOnly = false;
				window.parent.document.all("fax1").className = "Form";
			
			}	
		}		
		
		if(window.parent.document.all("company_email") != null)	
		{
			window.parent.document.all("company_email").value= window.document.all("email").value;
			if(private_flag == "1")
			{													
				window.parent.document.all("company_email").readOnly = true;
				window.parent.document.all("company_email").className = "Form_R";
			}
			else
			{
				window.parent.document.all("company_email").readOnly = false;
				window.parent.document.all("company_email").className = "Form";
			}			
		}	
		
			
		if(window.parent.document.all("url") != null)	
		{
			window.parent.document.all("url").value = window.document.all("url").value;
			if(private_flag == "1")
			{															
				window.parent.document.all("url").readOnly = true;
				window.parent.document.all("url").className = "Form_R";
			}
			else
			{
				window.parent.document.all("url").readOnly = false;
				window.parent.document.all("url").className = "Form";		
			}	
		}		

		
		if(window.parent.document.all("company_desc") != null)	
		{
			window.parent.document.all("company_desc").value= window.document.all("company_desc").value;
			if(private_flag == "1")
			{															
				window.parent.document.all("company_desc").readOnly = true;
				window.parent.document.all("company_desc").className = "Form_R";
			}
			else
			{
				window.parent.document.all("company_desc").readOnly = false;
				window.parent.document.all("company_desc").className = "Form";		
			}	
		}
		
		if(window.parent.document.all("cities_list") != null)	
		{
			if(private_flag == "1")
			{															
				window.parent.document.all("cities_list").disabled = true;
				window.parent.document.all("cities_list").style.display = "none";
				window.parent.document.all("cities_list").style.visible = false;
			}
			else
			{
				window.parent.document.all("cities_list").disabled = false;
				window.parent.document.all("cities_list").style.display = "inline";
				window.parent.document.all("cities_list").style.visible = true;		
			}	
		}		
		
	//-->
	</SCRIPT>  	
  	<%	
	End If
%>
<table width="100%" align=center cellpadding=0 cellspacing=0 border=0 bgcolor=#E6E6E6>
<tr>
    <td align="left" width="100%" valign=top dir="<%=dir_var%>" >
    <table width="100%" cellspacing="1" cellpadding="1" bgcolor=#FFFFFF>       
    <tr> 	      
	    <td width="150" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;<span id="word4" name=word4><%=arrTitles(4)%></span>&nbsp;</td>	    
  	    <td width="75"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;<span id="word6" name=word6><%=arrTitles(6)%></span></td>	    
  	    <td width="90"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="3" OR trim(sort)="4" then%>_act<%end if%>"><%if trim(sort)="3" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word18 title="למיון בסדר יורד"><%elseif trim(sort)="4" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word19 title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=3" name=word19 title="למיון בסדר עולה"><%end if%><span id="word7" name=word7><%=arrTitles(7)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort)="3" then%>bot<%elseif trim(sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
        <td width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="1" OR trim(sort)="2" then%>_act<%end if%>"><%if trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word19 title="למיון בסדר יורד"><%elseif trim(sort)="2" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word20 title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=1" name=word20 title="למיון בסדר עולה"><%end if%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>bot<%elseif trim(sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	     	              
	</tr>   	
	<%  sqlSelect="SELECT company_ID,company_name, phone1, fax1, email, city_Name FROM companies  "&_
		" WHERE ORGANIZATION_ID = " & trim(OrgID) & " AND isNULL(private_flag,0) = 0 "&_	
	    where_comp_name & " order by "& sortby(sort)
	   	set companyList = con.getRecordSet(sqlSelect)
	   	If not companyList.eof Then
		While not companyList.eof			
			companyID=companyList(0)
			companyName=rtrim(ltrim(companyList(1)))    
			phone=companyList(2)	
			fax=companyList(3)	
			email=companyList(4)
			cityName=companyList(5)			
	%>
	<tr>    	      	     
	      <td class="card" valign=top dir="<%=dir_obj_var%>" align="<%=align_var%>"><a class="link_categ" href="mailto:<%=email%>"><%=email%>&nbsp;</a></td>
	      <td class="card" valign=top dir="<%=dir_obj_var%>" align="center"><a class="link_categ" href="<%=urlSort%>&companyId=<%=companyId%>" target=_self style="font-size:11px"><%=phone%></a></td>	     
	      <td class="card" valign=top dir="<%=dir_obj_var%>" align="<%=align_var%>"><a class="link_categ" href="<%=urlSort%>&companyId=<%=companyId%>" target=_self><%=cityName%>&nbsp;</a></td>
          <td class="card" valign=top dir="<%=dir_obj_var%>" align="<%=align_var%>" dir=rtl><a class="link_categ" href="<%=urlSort%>&companyId=<%=companyId%>" target=_self style="line-height:120%;padding-top:3px;padding-bottom:3px"><%=companyName%></a></td>	                     
    </tr> 	
	<%
		companyList.moveNext
		Wend
		set companyList = Nothing
		Else%>
	<tr><td colspan=5 class=card dir="<%=dir_obj_var%>" align="center"><%=arrTitles(10)%></td></tr>	

	<%	End If	%>	
</td></tr></table>
<%	End If	%>	
<%set con =nothing%>
</BODY>
</HTML>
