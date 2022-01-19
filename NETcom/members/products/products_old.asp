<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">
<!--
function CheckDelProd() {
  return (confirm("?האם ברצונך למחוק את ההפצה"))    
}
function openPreview(pageId,prodId)
{
	result = window.open("../Pages/result.asp?prodId="+prodId+"&pageId="+pageId,"Result","toolbar=0, menubar=0, resizable=1, scrollbars=1, fullscreen = 0, width=780, height=540, left=2, top=2");
	return false;
}
function closeProduct(quest_id,prod_Id,show_insite)
{
	if(quest_id != null)
	{
		if(show_insite == "True")
			msg = "<%=Space(46)%>,שים לב\n\n.מרגע זה טופס משוב ייפתח למענים\n\n<%=Space(22)%>? האם ברצונך להמשיך"
		else
		    msg = "<%=Space(46)%>,שים לב\n\n.מרגע זה טופס משוב ייחסם למענים\n\n<%=Space(22)%>? האם ברצונך להמשיך"
		
		if(window.confirm(msg))
			document.location.href = "products.asp?prodId="+prod_Id+"&show="+show_insite;
	}
	return false;

}
function checkEmailLimit(emailLimit, IsOkToSend)
{	
	if(isNaN(emailLimit) == false)
	{
		emailLimit = parseInt(emailLimit);
		if(emailLimit <= 0)
		{
			window.alert("לא ניתן לבצע הפצה משום שיתרתך לשליחת מיילים נגמרה \n\n 04 - לקנית יתרה נוספת ניתן להתקשר ל 8770282");
			return false;
		}
	}
	IsOkToSend = chkIsOkToSend();
	if(IsOkToSend == '0')	
	{
		window.alert("ברגעים אלו מתבצעות מספר רב של הפצות במערכת במקביל \n\n על מנת שהפצה שלך תבוצע באופן מהיר ואמין \n\n אנא נסה לבצע את ההפצה בעוד מספר דקות");
		return false;
	}
	return true;
}

function chkIsOkToSend()
{
	var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");
    xmlHTTP.open("POST", 'chkIsOkToSend.asp', false);
    xmlHTTP.send('<?xml version="1.0" encoding="UTF-8"?>');
    
    //window.alert(xmlHTTP.ResponseText)
    return xmlHTTP.ResponseText;		 
}
//-->
</script>  
</head>

<%
	reqprodid = Request.QueryString("prodId")
	if request("delProd")="1" and reqprodid <> nil then
		
		set pr=con.GetRecordSet("select FILE_ATTACHMENT from products where product_id=" & reqprodid)
		if not pr.eof then
			fileName1=pr("FILE_ATTACHMENT")			
			If trim(fileName1) <> "" Then
			sqlstr = "Select TOP 1 product_id FROM products WHERE FILE_ATTACHMENT LIKE '" & sFix(fileName1) & "' AND product_id <> " & reqprodid
			set rs_check = con.getRecordSet(sqlstr)
			if rs_check.eof then
				set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!   			
   				file_path="../../../download/products/" & fileName1
   				'Response.Write fs.FileExists(server.mappath(file_path))
				'Response.End
				if fs.FileExists(server.mappath(file_path)) then
					set a = fs.GetFile(server.mappath(file_path))
					a.delete			
				end if	
			end if
			set rs_check = nothing	
			End If				
		end if
		con.ExecuteQuery("Delete from PRODUCT_GROUPS where product_id="&reqprodid & " and ORGANIZATION_ID=" & trim(OrgID) )
		con.ExecuteQuery("Delete from PRODUCT_CLIENT where product_id="&reqprodid & " and ORGANIZATION_ID=" & trim(OrgID) )
		con.ExecuteQuery("Delete from Form_Value where appeal_id in (Select appeal_id from appeals where product_id="&reqprodid&")" )
		con.ExecuteQuery("Delete from appeals where product_id="&reqprodid)
		con.ExecuteQuery("Delete from products where product_id="&reqprodid & " and ORGANIZATION_ID=" & trim(OrgID) )		
	end if
	
	if Request.QueryString("confirm") = "True" and reqprodid <> nil then
		con.ExecuteQuery("Update products Set confirm=1 Where product_id="&reqprodid & " and ORGANIZATION_ID=" & trim(OrgID) )
	elseif Request.QueryString("confirm") = "False" and reqprodid <> nil then
		con.ExecuteQuery("Update products Set confirm=0 Where product_id="&reqprodid & " and ORGANIZATION_ID=" & trim(OrgID) )
	end if
	
	if Request.QueryString("show") = "True" and reqprodid <> nil then
		con.ExecuteQuery("Update products Set show_insite=1, date_end = null Where product_id="&reqprodid & " and ORGANIZATION_ID=" & trim(OrgID) )
	elseif Request.QueryString("show") = "False" and reqprodid <> nil then
		con.ExecuteQuery("Update products Set show_insite=0, date_end = getdate() Where product_id="&reqprodid & " and ORGANIZATION_ID=" & trim(OrgID) )
	end if
	
	Function FormatMediumDate(DateValue)
		Dim strYYYY
		Dim strYY
		Dim strMM
		Dim strDD

			strYYYY = CStr(DatePart("yyyy", DateValue))
			If Len(strYYYY) > 2 Then
				strYY  = Right(strYYYY,2)
			Else
				strYY = strYYYY
			End If

			strMM = CStr(DatePart("m", DateValue))
			If Len(strMM) = 1 Then strMM = "0" & strMM

			strDD = CStr(DatePart("d", DateValue))
			If Len(strDD) = 1 Then strDD = "0" & strDD

			FormatMediumDate = strDD & "/" & strMM & "/" & strYY
	End Function 
	
	'------------------------------------ Check organization email limit  ---------------------		
	sql = "Select Email_Limit,Email_Limit_Month,date_add From ORGANIZATIONS WHERE ORGANIZATION_ID = " & OrgID
	set rs_org = con.getRecordSet(sql)	
	If not rs_org.eof Then
		emailLimit = trim(rs_org(0))
		EmailLimitMonth = trim(rs_org(1)) 
		date_add = trim(rs_org(2)) 
		number_of_month = DateDiff("m",date_add,Date())
		Date_curr_add = DateAdd("m",number_of_month + 1,date_add)		 
	End If
	set rs_org = Nothing	
	
	If isNumeric(emailLimit) Then
		emailLimit = cLng(emailLimit)
	End If
	'------------------------------------ End of Check organization email limit  ---------------------		
	
	'------------------------------------ Check if there is not current mail sending ---------------------	
	sqlstr = "SET DATEFORMAT DMY; Select Max(DATE_SEND_END) From PRODUCTS Where " & _
	" (PRODUCT_NUMBER Like '111') And (DateDiff(mi, DATE_SEND_END, GetDate()) <= 1)"
	Set rs_tmp = con.getRecordSet(sqlstr)
	If Not rs_tmp.Eof Then
		max_date_send_end = trim(rs_tmp(0))
	Else
		max_date_send_end = null
	End If
	Set rs_tmp = Nothing
	
	IsOkToSend = 1
	If IsDate(max_date_send_end) Then
		IsOkToSend = 0
	End If
	'------------------------------------ End if Check if there is not current mail sending ---------------------	
	
	dim sortby(16)	
	sortby(1) = 1'" CAST(p1.DATE_START as float), p1.product_id DESC"
	sortby(2) = 2'" CAST(p1.DATE_START as float) DESC, p1.product_id DESC"
	sortby(3) = 3'" p2.product_name, CAST(p1.DATE_START as float) DESC"
	sortby(4) = 4'" p2.product_name DESC, CAST(p1.DATE_START as float) DESC"
	sortby(5) = 5'" Page_Title, CAST(p1.DATE_START as float) DESC"
	sortby(6) = 6'" Page_Title DESC, CAST(p1.DATE_START as float) DESC"
	sortby(7) = 7'" count_clients"
	sortby(8) = 8'" count_clients DESC"
	sortby(9) = 9'" count_send"
	sortby(10) = 10'" count_send DESC"
	sortby(11) = 11'" count_open"
	sortby(12) = 12'" count_open DESC"
	sortby(13) = 13'" count_not_open"
	sortby(14) = 14'" count_not_open DESC"
	
	sort = Request("sort")	
	if sort = nil then
		sort = 2
	end if
	urlSort = "products.asp?1=1"	
%>
<%	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 75 Order By word_id"				
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
	  set rstitle = Nothing%>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 2%>
<%numOfLink = 2%>
<!--#include file="../../top_in.asp"-->
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
<tr><td class="page_title" dir="<%=dir_var%>">&nbsp;<%=arrTitles(1)%>&nbsp;-&nbsp;<span class="td_subj" style="color:#6E6DA6"><%=arrTitles(2)%>&nbsp;<font color=red><%=emailLimit%></font></font>&nbsp;
-&nbsp;<span class="td_subj" style="color:#6E6DA6"><!--תאריך חידוש יתרה--><%=arrTitles(3)%>&nbsp;<font color=red><%=Date_curr_add%></font></font>&nbsp;
-&nbsp;<span class="td_subj" style="color:#6E6DA6"><!--כמות חודשית--><%=arrTitles(4)%>&nbsp;<font color=red><%=EmailLimitMonth%></font></font>&nbsp;
</td></tr>
<tr><td align=center valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0>
<tr>
<td width=100% valign=top>
	<table width="100%" cellspacing="1" cellpadding="1" align=center border="0" bgcolor="#ffffff" dir="<%=dir_var%>">
	<tr>
		<td valign=top class="title_sort" align=center><!--מחיקה--><%=arrTitles(18)%></td>
		<td valign=top nowrap class="title_sort" align=center><!--משובים--><%=arrTitles(17)%></td>
		<td valign=top class="title_sort" align=center><!--קובץ מצורף--><%=arrTitles(16)%></td>
		<td valign=top nowrap align=center class="title_sort<%if sort=1 or sort=2 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=2 then%>1<%elseif sort=1 then%>2<%else%>2<%end if%>" target="_self"><!--תאריך הפצה--><%=arrTitles(15)%><img src="../../images/arrow_<%if trim(sort)="1" then%>down<%elseif trim(sort)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top">&nbsp;&nbsp;&nbsp;&nbsp;</a></td>		
		<td width="50" valign=top nowrap align=center class="title_sort<%if sort=13 or sort=14 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=14 then%>13<%elseif sort=13 then%>14<%else%>14<%end if%>" target="_self"><!--לא נפתחו--><%=arrTitles(14)%><img src="../../images/arrow_<%if trim(sort)="13" then%>down<%elseif trim(sort)="14" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="50" valign=top nowrap align=center class="title_sort<%if sort=11 or sort=12 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=12 then%>11<%elseif sort=11 then%>12<%else%>12<%end if%>" target="_self"><!--נפתחו--><%=arrTitles(13)%><img src="../../images/arrow_<%if trim(sort)="11" then%>down<%elseif trim(sort)="12" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="50" valign=top nowrap align=center class="title_sort<%if sort=9 or sort=10 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=10 then%>9<%elseif sort=9 then%>10<%else%>10<%end if%>" target="_self"><!--נשלחו--><%=arrTitles(12)%><img src="../../images/arrow_<%if trim(sort)="9" then%>down<%elseif trim(sort)="10" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	
		<td width="50" valign=top nowrap align=center class="title_sort<%if sort=7 or sort=8 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=7 then%>8<%elseif sort=8 then%>7<%else%>7<%end if%>" target="_self"><!--סה"כ--><%=arrTitles(10)%><img src="../../images/arrow_<%if trim(sort)="7" then%>down<%elseif trim(sort)="8" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"><br><!--נמענים--><%=arrTitles(11)%>&nbsp;&nbsp;</a></td>	
		<td width="36" valign=top nowrap class="title_sort" align=center><!--פעיל--><%=arrTitles(9)%></td>
		<!--td width="140" valign=top nowrap align="<%=align_var%>" class="title_sort<%if sort=3 or sort=4 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=4 then%>3<%elseif sort=3 then%>4<%else%>4<%end if%>" target="_self">טופס מופץ&nbsp;<img src="../../images/arrow_<%if trim(sort)="3" then%>down<%elseif trim(sort)="4" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td-->
		<td width="150" valign=top nowrap class="title_sort" align="<%=align_var%>">&nbsp;<!--נושא--><%=arrTitles(8)%>&nbsp;</td>
		<td valign=top class="title_sort" align=center><!--הפצה חוזרת--><%=arrTitles(7)%></td>		
		<td width="100%" align="<%=align_var%>" dir="<%=dir_var%>" valign=top class="title_sort<%if sort=5 or sort=6 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=6 then%>5<%elseif sort=5 then%>6<%else%>6<%end if%>" target="_self"><!--דף מופץ--><%=arrTitles(6)%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="5" then%>down<%elseif trim(sort)="6" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>		
		<td valign=top class="title_sort" align=center><!--סוג--><%=arrTitles(5)%></td>
	</tr>
<%	
    if Request("Page")<>"" then
		Page=request("Page")
	else
		Page=1
	end if
	sqlstr = "Exec dbo.GetProductsSend '" & trim(OrgID) & "','" & sortby(sort) & "'"
	'Response.Write sqlStr
	set prod = con.GetRecordSet(sqlStr)
    if not prod.eof then
		If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
			PageSize = RowsInList
		Else	
     		PageSize = 10
		End If	   
	    prod.PageSize = PageSize
		prod.AbsolutePage=Page
		recCount=prod.RecordCount 	
		prod_count=prod.RecordCount	
		NumberOfPages = prod.PageCount
		i=1
		j=0
		do while (not prod.eof and j<prod.PageSize)
    	prod_Id = trim(prod("product_id"))
    	FILE_ATTACHMENT = trim(prod("FILE_ATTACHMENT"))
    	If Len(FILE_ATTACHMENT) > 0 Then
    		attachment_path = "../../../download/products/" & FILE_ATTACHMENT
    	Else
    		attachment_path = ""	
    	End If
    	page_id = trim(prod("page_id"))
    	pageTitle = trim(prod("page_title"))    	
    	If trim(pageTitle)="" Or isNull(pageTitle) Then
    		pageTitle = "אין"
    	End if    	
    	questions_id = prod("questions_id")
    	product_name = prod("product_name")
    	If trim(product_name) = "" Or IsNull(product_name) Then
    		product_name = "אין"
    	End if
    	If IsDate(prod("date_start")) Then
    		date_start=FormatMediumDate(prod("date_start"))
    	Else
    		date_start=""
    	End If	
    	if prod("show_insite") = "1" then 
    		show_insite = true
    	else
    		show_insite = false
    	end if      	
    	EMAIL_SUBJECT = trim(prod("EMAIL_SUBJECT"))  
    	product_type = trim(prod("product_type")) 
    	If trim(product_type) = "3" Then
    		If trim(lang_id) = "1" Then
    			product_type_name = "דואר"
    		Else
    			product_type_name = "Postal mail"
    		End If    			
    	Else
    		If trim(lang_id) = "1" Then
    			product_type_name = "מייל"
    		Else
    			product_type_name = "Email"
    		End If    		
    	End If 
    	count_send = trim(prod("count_send"))
    	If IsNull(count_send) = false And IsNumeric(count_send) = true Then
    		If cInt(count_send) > 0 Then
    		issend = true
			send = cInt(count_send)
			Else
			issend = false
			send = 0	
			End If
		else
			issend = false
			send = 0
		end if
		count_clients = trim(prod("count_clients"))
		If IsNull(count_clients) = false And IsNumeric(count_clients) = true Then
    		If cInt(count_clients) > 0 Then
    		isclients = true
			clients = cInt(count_clients)
			Else
			isclients = false
			clients = 0
			End If
		else
			isclients = false
			clients = 0
		end if
		
		count_answered = trim(prod("count_answered"))
		If IsNull(count_answered) = false And IsNumeric(count_answered) = true Then
			If cInt(count_answered) > 0 Then
    		isanswered = true
			answers = cInt(count_answered)
			Else
			isanswered = false
			answers = 0
			End If
		else
			isanswered = false
			answers = 0
		end if		
		
		count_open = trim(prod("count_open"))
		If IsNull(count_open) = false And IsNumeric(count_open) = true Then
			If cInt(count_open) > 0 Then
    		isopen = true
			opened = cInt(count_open)
			Else
			isopen = false
			opened = 0
			End If
		else
			isopen = false
			opened = 0
		end if	
		
		count_not_open = trim(prod("count_not_open"))
		If IsNull(count_not_open) = false And IsNumeric(count_not_open) = true Then
			If cInt(count_not_open) > 0 Then
			not_opened = cInt(count_not_open)
			Else
			not_opened = 0
			End If
		else
			not_opened = 0
		end if			
%>
<tr>
	<td align="center" class="card" nowrap>&nbsp;<a href="products.asp?prodId=<%=prod_Id%>&delProd=1" ONCLICK="return CheckDelProd()"><IMG SRC="../../images/delete_icon.gif" BORDER=0 alt="מחיקת הפצה"></a>&nbsp;</td>
	<td align="center" class="card" nowrap>&nbsp;<%If isanswered Then%><a href="../appeals/feedbacks.asp?sekerId=<%=prod_Id%>&prodId=<%=questions_id%>" title="<%=answers%> - מספר מענים"><IMG SRC="../../images/check_icon.gif" BORDER=0></a>&nbsp;<%End If%></td>
	<td align="center" class="card" nowrap><%If Len(attachment_path) > 0 Then%><a href="<%=attachment_path%>" target=_blank class=link1><img src="../../images/bull2.gif" border=0 hspace=0 vspace=0></a><%Else%>&nbsp;<%End If%></td>
	<td align="center" class="card" nowrap>&nbsp;<%=date_start%>&nbsp;</td>
	<td align="center" class="card" nowrap><%If Not trim(product_type) = "3" Then%><a href="javascript:void(0);" OnClick="javascript:window.open('not_opened_list.asp?prodID=<%=prod_Id%>&pageID=<%=page_id%>','NotOpenedList','left=50,top=50,scrollbars=1,width=650,height=500');" class="link1"><%=not_opened%></a><%End If%></td>
	<td align="center" class="card" nowrap><%If Not trim(product_type) = "3" Then%><a href="javascript:void(0);" OnClick="javascript:window.open('opened_list.asp?prodID=<%=prod_Id%>&pageID=<%=page_id%>','OpenedList','left=50,top=50,scrollbars=1,width=650,height=500');" class="link1"><%=opened%></a><%End If%></td>
	<td align="center" class="card" nowrap><%=send%></td>
	<td align="center" class="card" nowrap><%=clients%></td>	
	<td align="center" class="card" nowrap><%if show_insite = false then%><IMG SRC="../../images/lamp_off.gif" BORDER=0 alt="לא מופיע באתר" class=hand onClick="return closeProduct('<%=questions_id%>','<%=prod_Id%>','<%=not show_insite%>')"><%else%><IMG SRC="../../images/lamp_on.gif" BORDER=0 alt="מופיע באתר" class=hand onClick="return closeProduct('<%=questions_id%>','<%=prod_Id%>','<%=not show_insite%>')"><%end if%></td>
	<!--td align="<%=align_var%>"  class="card" nowrap><%If IsNumeric(questions_id) Then%><a href="#" onclick="window.open('check_form.asp?prodId=<%=questions_id%>','Preview','left=20,top=20,tollbar=0,menubar=0,scrollbars=1,resizable=0,width=660');" class="link_categ"><%=product_name%>&nbsp;</a><%Else%><%=product_name%>&nbsp;<%End If%></td-->
	<td align="<%=align_var%>" class="card" dir="rtl">&nbsp;<%=EMAIL_SUBJECT%>&nbsp;</td>
	<td align="center" class="card" nowrap><%If Not trim(product_type) = "3" Then%><!--a href="default.asp?prodId=<%=prod_Id%>"><IMG SRC="../../images/list_icon.gif" BORDER=0 alt="הפצה חוזרת"></a--><%End If%>&nbsp;</td>
	<td align="<%=align_var%>"  class="card" dir="rtl"><a href="#" onclick="return openPreview('<%=page_id%>','<%=prod_Id%>')" class="link_categ"><%=pageTitle%>&nbsp;</a></td>
	<td align="<%=align_var%>" class="card" dir="rtl">&nbsp;<%=product_type_name%>&nbsp;</td>
</tr>
<%		prod.MoveNext
        j=j+1
		loop
	end if
	set prod=nothing%>
<% if NumberOfPages > 1 then%>
		<tr class="card">
		<td width="100%" align=center nowrap class="card" colspan="13">
			<table border="0" cellspacing="0" cellpadding="2">               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRow") <> nil Then
	               numOfRow = Request.QueryString("numOfRow")
	           Else numOfRow = 1
	           End If
	         %>
	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow-1)-9%>&numOfRow=<%=numOfRow-1%>" title="לדפים הקודמים"><b><<</b></a></td>			                
			  <%end if%>
	             <td><font size="2" color="#060165">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="center"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	                  <%else%>
	                     <td align="center"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=i+10*(numOfRow-1)%>&numOfRow=<%=numOfRow%>"><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td><font size="2" color="#060165">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow) + 1%>&numOfRow=<%=numOfRow+1%>" title="לדפים הבאים"><b>>></b></a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>				
	<%End If%>		
	<tr><td align=center height=20px class="card" colspan="13"><font style="color:#6E6DA6;font-weight:600"><%=prod_count%> : <!--סה"כ הפצות--><%=arrTitles(10)%>&nbsp;<%=arrTitles(1)%></font></td></tr>								
</table></td>
<td width=90 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100%>
<tr><td align="<%=align_var%>" colspan=2 height="30" nowrap></td></tr>
<!--tr><td colspan=2 align="center"><a class="button_edit_1" style="width:84;line-height:11pt;padding:2pt;" href='addprod.asp' onclick="return checkEmailLimit('<%=emailLimit%>','<%=IsOkToSend%>')"><!--הפצת מייל--><%'arrTitles(19)%><!--/a></td></tr-->
<tr><td align="<%=align_var%>" colspan=2 height="5" nowrap></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:84;line-height:11pt;padding:2pt;" href='addmail.asp'><!--הפצת דואר--><%=arrTitles(20)%></a></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr>
</table></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr></table>
</body>
</html>
<%set con=nothing%>