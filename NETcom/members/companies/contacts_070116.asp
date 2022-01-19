<!--  רשימת אנשי הקשר הכוללת אפשרות חיפוש, מיון ודיפדוף -->
<%@ LANGUAGE=VBScript %>
<%Response.Buffer = True%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%search_contact = trim(Request.Form("search_contact"))
	 if trim(Request.QueryString("search_contact")) <> "" then
		search_contact = trim(Request.QueryString("search_contact"))
	 end if					
	 
	search_contact_name = trim(Request.Form("search_contact_name"))
	if trim(Request.QueryString("search_contact_name")) <> "" then
	search_contact_name = trim(Request.QueryString("search_contact_name"))
	end if	 
	If Len(search_contact_name) > 0 Then
		search_contact = trim(search_contact_name)
	End If
	
	search_company = trim(Request.Form("search_company"))
	if trim(Request.QueryString("search_company")) <> "" then
	search_company = trim(Request.QueryString("search_company"))
	end if
	 	
	search_messanger_name = trim(Request.Form("search_messanger_name"))
	if trim(Request.QueryString("search_messanger_name")) <> "" then
	search_messanger_name = trim(Request.QueryString("search_messanger_name"))
	end if
	
	search_city_Name = trim(Request.Form("search_city_Name"))
	if trim(Request.QueryString("search_city_Name")) <> "" then
	search_city_Name = trim(Request.QueryString("search_city_Name"))
	end if		
		 
	search_phone = trim(Request.Form("search_phone"))
	if trim(Request.QueryString("search_phone")) <> "" then
	search_phone = trim(Request.QueryString("search_phone"))
	end if
	
	'search_cellular = trim(Request.Form("search_cellular"))
	'if trim(Request.QueryString("search_cellular")) <> "" then
	'search_cellular = trim(Request.QueryString("search_cellular"))
	'end if			
	search_cellular=""
	search_email = trim(Request.Form("search_email"))
	if trim(Request.QueryString("search_email")) <> "" then
	search_email = trim(Request.QueryString("search_email"))
	end if	
	
	show_only_double = trim(Request.Form("show_only_double"))
	if trim(Request.QueryString("show_only_double")) <> "" then
	show_only_double = trim(Request.QueryString("show_only_double"))
	end if
	'Response.Write show_only_double
	If trim(show_only_double) = "true" Or trim(show_only_double) = "1" Then
		show_only_double = 1
	Else 
		show_only_double = 0
	End If
		
	OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
	lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))
	RowsInList = Request.Cookies("bizpegasus")("RowsInList")
	is_companies = trim(Request.Cookies("bizpegasus")("ISCOMPANIES"))
	
  	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If
	If lang_id = 2 Then
		dir_var = "rtl"  : align_var = "left"  : dir_obj_var = "ltr"
		arr_status = Array("","new","active","close","appeal")
	Else
		dir_var = "ltr"  :  align_var = "right"  :  dir_obj_var = "rtl"
		arr_status = Array("","עתידי","פעיל","סגור","פונה")
	End If		

	if trim(Request.QueryString("contact_type"))<>"" then
		contact_type = trim(Request.QueryString("contact_type"))
	else
		contact_type = 0	
	end if 
		 
	if trim(Request.QueryString("page"))<>"" then
		page=Request.QueryString("page")
	else
		Page=1
	end if  
	 
	if trim(Request.QueryString("numOfRow"))<>"" then
		numOfRow=Request.QueryString("numOfRow")
	else
		numOfRow = 1
	end if  
 
    sort = Request.QueryString("sort")	
    if trim(sort)="" then  sort = 0  end if

	urlSort="contacts.asp?search_contact=" & Server.URLEncode(search_contact) & "&contact_type=" & contact_type  & _
	"&amp;search_email=" & Server.URLEncode(search_email) & _
	"&amp;search_company=" & Server.URLEncode(search_company) & _
	"&amp;search_messanger_name=" & Server.URLEncode(search_messanger_name) & _
	"&amp;search_city_Name=" & Server.URLEncode(search_city_Name) & _
	"&amp;search_phone=" & search_phone&"&amp;search_cellular=" & Server.URLEncode(search_cellular)	&_
	"&amp;show_only_double=" & show_only_double
	    
	If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
		PageSize = RowsInList
	Else	
     	PageSize = 10
	End If	
     
	dim sortby(8)	
	sortby(0) = 0'"rtrim(ltrim(CONTACT_NAME))"
	sortby(1) = 1'"rtrim(ltrim(company_name))"
	sortby(2) = 2'"rtrim(ltrim(company_name)) DESC"
	sortby(3) = 3'"rtrim(ltrim(messanger_name))"
	sortby(4) = 4'"rtrim(ltrim(messanger_name)) DESC"
	sortby(5) = 5'"rtrim(ltrim(CONTACT_NAME))"
	sortby(6) = 6'"rtrim(ltrim(CONTACT_NAME)) DESC"
	sortby(7) = 7'"rtrim(ltrim(contact_city_Name))"
	sortby(8) = 8'"rtrim(ltrim(contact_city_Name)) DESC"     
    
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 3 Order By word_id"				
	set rstitle = con.getRecordSet(sqlstr)		
	If not rstitle.eof Then
	arrTitlesD = rstitle.getRows()		
	redim arrTitles(Ubound(arrTitlesD,2)+1)
	For i=0 To Ubound(arrTitlesD,2)		
		arrTitles(arrTitlesD(0,i)) = arrTitlesD(1,i)		
	Next
	End If
	set rstitle = Nothing%>
<html>
<head>
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--
		var oPopup = window.createPopup();	
		function contact_typeDropDown(obj)
		{
			oPopup.document.body.innerHTML = contact_type_Popup.innerHTML; 
			oPopup.document.charset="windows-1255";
			oPopup.show(0-120+obj.offsetWidth, 17, 120, 148, obj);    
		}
		
		function doSearch()
		{
			var searchCellular = new String(document.getElementById("search_phone").value);
			if(document.getElementById("search_company"))
			{
				searchCompany = document.getElementById("search_company").value;
			}
			else
			{
				searchCompany = "";
			}			
			document.location.href = 'contacts.asp?search_cellular=' + searchCellular + 			
			'&search_phone=' + document.getElementById("search_phone").value +
			'&search_city_Name=' + document.getElementById("search_city_Name").value +
			'&search_contact_name=' + document.getElementById("search_contact_name").value +
			'&search_company=' + searchCompany +
			'&search_messanger_name=' + document.getElementById("search_messanger_name").value + 
			'&search_email=' + document.getElementById("search_email").value +
			'&show_only_double=' + document.getElementById("show_only_double").checked;		
			return true;
		}	
		
		function entsubsearchg(e) { 
			if( typeof( e ) == "undefined" && typeof( window.event ) != "undefined" )
				e = window.event;
    
			if (e && e.keyCode == 13)
				return doSearch();		
			else
				return true; }	
				
			function DoCallback(url, params)
			{
				var pageUrl = url + "?callback=true&" + params;
				var xmlRequest = new ActiveXObject("Microsoft.XMLHTTP");
				xmlRequest.open("GET", pageUrl, false);
				xmlRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
				xmlRequest.send(null);
			}				
				
			function chkDouble(isDouble, ContactId)
			{
				//window.alert(isDouble)
				if(isDouble == true)
					isDouble =  1;
				else
					isDouble = 0;	
				DoCallback("change_double.asp", "ContactId=" + ContactId + "&isDouble=" + isDouble);
				
				if (isDouble == 1)
				{
					document.getElementById("tr_" + ContactId).style.backgroundColor = "#F2E1E1";
					document.getElementById("tr_" + ContactId).onmouseout = function() { this.style.backgroundColor='#F2E1E1'; };
				}
				else
				{
					document.getElementById("tr_" + ContactId).style.backgroundColor = "#E6E6E6";
					document.getElementById("tr_" + ContactId).onmouseout = function() { this.style.backgroundColor='#E6E6E6'; };
				}
				return true;
			}	
						
//-->
</script>
</head>
<body scroll=no onload="resizeTo(document.body.scrollWidth,document.body.scrollHeight)" style="margin:0px;SCROLLBAR-FACE-COLOR: #E6E6E6;SCROLLBAR-HIGHLIGHT-COLOR: #F7F7F7;SCROLLBAR-SHADOW-COLOR: #848484;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #808080;SCROLLBAR-TRACK-COLOR: #E6E6E6;SCROLLBAR-DARKSHADOW-COLOR: #ffffff">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td class="card" align="left" width="100%" valign=top dir="<%=dir_var%>">
    <table width="100%" cellspacing="1" cellpadding="1" border=0 bgcolor="#ffffff" >
		<tr>
  			<td class="title_sort" align="right" dir="ltr"><input type="button" 
			onclick="javascript:document.location.href='contacts.asp';" 
			id="btnShowAll" value="הראה הכל" style="width: 64px" class="button" NAME="btnShowAll">&nbsp;&nbsp;
			<input type="button" onclick="return doSearch();" id="btnSearch" value="חפש" style="width: 44px" 
			class="button" NAME="btnSearch">	  			
  		</td>	      	    
		<td width="150" nowrap class="title_sort" align="right" dir="rtl"><input type="text" class="search" dir="rtl" style="width:148px;" value="<%=vFix(search_email)%>" name="search_email" ID="search_email" onkeypress="return entsubsearchg(event);"></td>  		
  		<td width="150" nowrap class="title_sort" align="center" dir="rtl" colspan=3><input type="text" class="search" dir="rtl" style="width:73px;" value="<%=vFix(search_phone)%>" name="search_phone" ID="search_phone" onkeypress="return entsubsearchg(event);"></td>
		<td width="100" nowrap class="title_sort" align="center" dir="rtl"><input type="text" class="search" dir="rtl" style="width:98px;" value="<%=vFix(search_city_Name)%>" name="search_city_Name" ID="search_city_Name" onkeypress="return entsubsearchg(event);"></td>
		<td width="85" nowrap class="title_sort" align="center" dir="rtl"><input type="text" class="search" dir="rtl" style="width:83;" value="<%=vFix(search_messanger_name)%>" name="search_messanger_name" ID="search_messanger_name" onkeypress="return entsubsearchg(event);" ></td>
		<%If trim(is_companies) = "1" Then%>
		<td width="100%" nowrap class="title_sort" align="center" dir="rtl"><input type="text" class="search" dir="rtl" style="width:98%;" value="<%=vFix(search_company)%>" name="search_company" ID="search_company" onkeypress="return entsubsearchg(event);" ></td>
		<%End If%>
		<td width="100" nowrap class="title_sort" align="center" dir="rtl"><input type="text" class="search" dir="rtl" style="width:98;" value="<%=vFix(search_contact_name)%>" name="search_contact_name" ID="search_contact_name" onkeypress="return entsubsearchg(event);" ></td>
		<td nowrap class="title_sort">&nbsp;</td>
		<td nowrap class="title_sort"><input type="checkbox" id="show_only_double" <%If show_only_double = 1 Then%> checked <%End If%> value="1" title="הצג רק לקוחות כפולים" ></td>
	</tr>	    
    <tr> 	      
	    <td width="120" id=td_group name=td_group nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;<!--קבוצה--><%=arrTitles(3)%>&nbsp;<IMG id="choose1" name=word16 style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(16)%>" align=absmiddle onclick="return false" onmousedown="contact_typeDropDown(td_group)"></td>
	    <td width="150" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;<!--אימייל--><%=arrTitles(4)%>&nbsp;</td>
	    <td width="72"  nowrap class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<!--טלפון נייד--><%=arrTitles(5)%></td>
  	    <td width="75"  nowrap class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<!--טלפון ישיר-->טלפון נוסף</td>
  	
  	    <td width="75"  nowrap class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<!--טלפון ישיר--><%=arrTitles(6)%></td>
  	    <td width="100" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="7" OR trim(sort)="8" then%>_act<%end if%>"><%if trim(sort)="7" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word17 title="<%=arrTitles(17)%>"><%elseif trim(sort)="8" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word18 title="<%=arrTitles(18)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=7" name=word18 title="<%=arrTitles(18)%>"><%end if%>&nbsp;<!--עיר--><%=arrTitles(22)%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="7" then%>bot<%elseif trim(sort)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	    <td width="85"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="3" OR trim(sort)="4" then%>_act<%end if%>"><%if trim(sort)="3" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word17 title="<%=arrTitles(17)%>"><%elseif trim(sort)="4" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word18 title="<%=arrTitles(18)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=3" name=word18 title="<%=arrTitles(18)%>"><%end if%>&nbsp;<!--תפקיד--><%=arrTitles(7)%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="3" then%>bot<%elseif trim(sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	    <%If trim(is_companies) = "1" Then%>
        <td width=100% align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="1" OR trim(sort)="2" then%>_act<%end if%>"><%if trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word17 title="ל<%=arrTitles(17)%>"><%elseif trim(sort)="2" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word18 title="<%=arrTitles(18)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=1" name=word18 title="<%=arrTitles(18)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>bot<%elseif trim(sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	     	      
        <%End If%>
        <td width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="5" OR trim(sort)="6" then%>_act<%end if%>"><%if trim(sort)="5" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word17 title="<%=arrTitles(17)%>"><%elseif trim(sort)="6" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word18 title="<%=arrTitles(18)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=5" name=word18 title="<%=arrTitles(18)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="5" then%>bot<%elseif trim(sort)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
        <td nowrap class="title_sort">&nbsp;</td>
        <td nowrap class="title_sort">&nbsp;</td>
	</tr>   
<%
if sFix(search_contact)<>"" or sFix(search_email)<>"" or sFix(search_company)<>"" or sFix(search_city_Name)<>"" or sFix(search_messanger_name) <>"" or  sFix(search_phone)<>"" or sFix(search_cellular)<>"" or  contact_type>0 then
sqlSelect = "Exec dbo.get_company_contacts @Page=" & Page & ", @RecsPerPage=" & PageSize & ", @OrgID=" & OrgID & _
	", @contact_type=" & contact_type & ", @contact_name='" & sFix(search_contact)  & "', @search_email='" & sFix(search_email) & _
	"', @company_name='" & sFix(search_company) & "', @city_name='" & sFix(search_city_Name) & _
	"', @messanger_mame = '" & sFix(search_messanger_name) & "', @phone = '" & sFix(search_phone) & _
	"', @cellular = '" & sFix(search_cellular) & "', @sort='" & sortby(sort) & "', @show_only_double=" & show_only_double

	'Response.Write (sqlSelect) 
	'Response.End
	set contactList=con.GetRecordSet(sqlSelect) 
	If not contactList.EOF then	
		recCount = contactList("CountRecords")
		arr_conts = contactList.getRows()
	End If	
	Set contactList = Nothing
	
	If isArray(arr_conts) Then
    For cc=0 To Ubound(arr_conts,2)
		If cc Mod 2 = 0 Then
			tr_bgcolor = "#E6E6E6"
		Else
			tr_bgcolor = "#C9C9C9"
		End If	    
		contactId=trim(arr_conts(1,cc))
		contact=trim(arr_conts(2,cc))
		companyID=trim(arr_conts(3,cc))
		companyName=rtrim(ltrim(arr_conts(4,cc)))      
		phone=trim(arr_conts(5,cc))
		cellular=trim(arr_conts(6,cc))
		duty=trim(arr_conts(7,cc))
		email=trim(arr_conts(8,cc))
		contact_city_name=rtrim(ltrim(arr_conts(9,cc)))  
		types=rtrim(ltrim(arr_conts(10,cc)))	
		phone_additional=rtrim(ltrim(arr_conts(13,cc)))	
		If Len(types) > 20 Then
			short_types = Left(types, 17) & ".."
		Else short_types=types
		End If      		
		If isNull(arr_conts(11,cc)) Then
			is_double = 0
		ElseIf cBool(arr_conts(11,cc)) Then			
			is_double = 1	: tr_bgcolor = "#F2E1E1"
		Else
			is_double = 0
		End If		%>
		<tr id="tr_<%=contactId%>" bgcolor="<%=tr_bgcolor%>" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='<%=tr_bgcolor%>';">
	      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" title="<%=vFix(types)%>"><a class="link_categ" href="contact.asp?companyId=<%=companyId%>&contactId=<%=contactId%>" target=_parent>&nbsp;<%=short_types%>&nbsp;</a></td>
	      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top">&nbsp;<%If Len(email) > 0 Then%><a class="file_link" style="font-size:11px" href="mailto:<%=email%>"><%=email%></a><%End If%>&nbsp;</td>
	      <td align="center" dir="<%=dir_obj_var%>" valign="top"><a class="link_categ" href="contact.asp?companyId=<%=companyId%>&contactId=<%=contactId%>" target=_parent style="font-size:11px"><%=cellular%></a></td>
	      <td align="center" dir="<%=dir_obj_var%>" valign="top"><a class="link_categ" href="contact.asp?companyId=<%=companyId%>&contactId=<%=contactId%>" target=_parent style="font-size:11px"><%=phone_additional%></a></td>
	      <td align="center" dir="<%=dir_obj_var%>" valign="top"><a class="link_categ" href="contact.asp?companyId=<%=companyId%>&contactId=<%=contactId%>" target=_parent style="font-size:11px"><%=phone%></a></td>
	      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top"><a class="link_categ" href="contact.asp?companyId=<%=companyId%>&contactId=<%=contactId%>" target=_parent>&nbsp;<%=contact_city_name%>&nbsp;</a></td>
	      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top"><a class="link_categ" href="contact.asp?companyId=<%=companyId%>&contactId=<%=contactId%>" target=_parent>&nbsp;<%=duty%>&nbsp;</a></td>
	      <%If trim(is_companies) = "1" Then%>
          <td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top"><a class="link_categ" href="contact.asp?companyId=<%=companyId%>&contactId=<%=contactId%>" target=_parent style="line-height:120%;padding-top:3px;padding-bottom:3px">&nbsp;<%=companyName%>&nbsp;</a></td>	        
          <%End If%>
          <td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top"><a class="link_categ" href="contact.asp?companyId=<%=companyId%>&contactId=<%=contactId%>" target=_parent>&nbsp;<%=contact%>&nbsp;</a></td>
		  <td align="center" valign="top"><img src="../../images/forms_icon.gif" border="0" style="cursor: pointer;"	
		  alt="היסטוריית טפסים של הלקוח" onclick="window.open('../appeals/contact_appeals.asp?contactID=<%=contactId%>','winCA','top=20, left=10, width=950, height=500, scrollbars=1');"  ></td>
          <td align="center" valign="top"><input type="checkbox" name="chk_double" value="<%=is_double%>" title="לקוח כפול"
          onclick="return chkDouble(this.checked, '<%=contactId%>');" <%If is_double = 1 Then%> checked <%End If%>    ></td>
	   </tr><%
	  Next
	  NumberOfPages = Fix((recCount / PageSize)+0.99)
	  
	  If NumberOfPages > 1 Then
	  urlSort = urlSort & "&sort=" & sort
	  %>
	  <tr>
		<td width="100%" align=middle colspan="10" nowrap bgcolor="#e6e6e6" dir=ltr>
			<table border="0" cellspacing="0" cellpadding="2">               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If  %>	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter title="<%=arrTitles(19)%>" href="<%=urlSort%>&page=<%=10*(numOfRow-1)-9%>&amp;numOfRow=<%=numOfRow-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	                  <%else%>
	                     <td  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter title="<%=arrTitles(20)%>" href="<%=urlSort%>&page=<%=10*(numOfRow) + 1%>&amp;numOfRow=<%=numOfRow+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>			
	<%	End if	%>
	<tr>
	   <td colspan="10" height=18 bgcolor="#e6e6e6" align=center class="card" dir="ltr" style="color:#6E6DA6;font-weight:600"><!--נמצאו--><%=arrTitles(8)%>&nbsp;<%=recCount%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsMulti"))%></td>
	</tr>
	<% 
	Else %>
	<tr>
	   <td colspan="10" align=center class="title_sort1" dir="rtl"><!--לא נמצאו--><%=arrTitles(9)%>&nbsp;<%=arrTitles(23)%></td>
	</tr>
<% End If%><%end if 'if search%>
</table></td></tr>
</table>
<DIV ID="contact_type_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="position:absolute;overflow:auto; top:0; left:0; width:120; height:148; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;SCROLLBAR-FACE-COLOR:#D3D3D3;SCROLLBAR-HIGHLIGHT-COLOR: #D3D3D3; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e5e5e5;" >
<%set contact_typeList=con.GetRecordSet("SELECT type_ID,type_Name FROM contact_type WHERE Organization_Id = " & trim(OrgID) & " ORDER BY type_Name")
  If not contact_typeList.EOF Then
	arr_cont_types = contact_typeList.getRows()
  End If
  set contact_typeList=Nothing
  
  If isArray(arr_cont_types) Then
  For cc=0 To Ubound(arr_cont_types,2)  
	prcontact_typeID = trim(arr_cont_types(0,cc))
	prcontact_typeName = trim(arr_cont_types(1,cc))%>
   <DIV onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#d3d3d3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#d3d3d3; border:1px solid black; padding-right:3px;  padding-left:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='contacts.asp?search_contact=<%=Server.URLEncode(search_contact)%>&sort=<%=sort%>&contact_type=<%=prcontact_typeID%>'">
    <%=prcontact_typeName%></DIV>              
 <%Next
 End If%>
    <DIV onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#d3d3d3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#d3d3d3; border:1px solid black; padding-right:3px;  padding-left:3px;cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='contacts.asp?search_contact=<%=Server.URLEncode(search_contact)%>&sort=<%=sort%>'">
    <!--הכל--><%=arrTitles(14)%></DIV>
</DIV> 
</body>
</html>
<%set con=Nothing%>