<%@ Language=VBScript%>

<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%numOfTab = 1%>
<%numOfLink = 4%>
<%
 urlSort = "GeneralEmailByGroup.asp?clname=" & Server.URLEncode(clname) & "&orgname=" & Server.URLEncode(orgname) & _
  "&usname=" & Server.URLEncode(usname) & "&orderOwner=" & Server.URLEncode(orderOwner) 


dim sortby(2)	
	sortby(1) = " response_date"
	sortby(2) = "response_date DESC"
	sort = Request("sort")	
	if sort = nil then
		sort = 2
	end if
	if (Request("prodId") <> nil) then 'and Request.QueryString("delProd") = nil) then
		prodId = Request("prodId")
		else
		prodId=0
		end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../../IE4.css" rel="stylesheet" type="text/css">
</head>
<body topmargin=0 bottommargin=0 leftmargin=0 rightmargin=0 bgcolor="#E5E5E5">
<!--#include file="../../logo_top.asp"-->
<table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>" ID="Table1">
			<FORM action="GeneralEmailByGroup.asp?sort=<%=sort%>" method=POST id="form1" name="form1" target="_self">
		
<tr><td>
	<table border="0" width="100%" cellspacing="0" cellpadding="0" bgcolor="#060165" ID="Table2">
		<tr><td width=100% align="<%=align_var%>"><!--#include file="../../top_in.asp"--></td></tr>
	</table>
</td></tr>
<tr><td>
		<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" ID="Table3">
			<tr> <td class="page_title" style="border-left:none" width="100%">
 <SELECT dir="<%=dir_obj_var%>" size=1 ID="prodId" name="prodId" class="sel" style="width:310px;font-size:10pt" onChange="document.location.href='GeneralEmailByGroup.asp?sort=<%=sort%>&prodID='+this.value">	
	<OPTION value="">כל הטפסים </OPTION>
<%
	If is_groups = 0 Then
		sqlstr = "SELECT product_id, product_name FROM dbo.Products WHERE (product_number = '0') " & _
		" AND (ORGANIZATION_ID=" & OrgID & ") ORDER BY product_name"
		' משתמש אשר שייך לקבוצה אבל אינו אחראי באף קבוצה
		Else
		sqlstr = "exec dbo.get_products_list '" & OrgID & "','" & UserID & "'"
	End If
	'Response.Write sqlstr
	'Response.End
	set rs_products = con.GetRecordSet(sqlstr)
	if not rs_products.eof then 
		ResProductsList = rs_products.getRows()		
	end if
	set rs_products=nothing				
	If IsArray(ResProductsList) Then
	For i=0 To Ubound(ResProductsList,2)
		prod_Id = ResProductsList(0,i)   	
		product_name = ResProductsList(1,i)
%>
	<OPTION value="<%=prod_Id%>" <%If trim(prodId) = trim(prod_Id) Then%> selected <%End If%>> <%=product_name%> </OPTION>
<%	Next	
 End If	
%>
</SELECT>&nbsp;&nbsp;&nbsp;</td></tr>
			<tr><td height=0 nowrap></td></tr>
			<tr><td width="100%" valign="top" align="center">
			
			<!--start code-->
			
			<!-- start code --> 
<%if Request("Page")<>"" then
		Page=request("Page")
	else
		Page=1
	end if
	If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
		 PageSize = RowsInList
	Else	
    	 PageSize = 10
	End If
	if trim(Request("textcont")) <> "" then
		textcont=trim(Request("textcont"))
	else
		textcont=""
	end if
	if trim(Request("textTitle")) <> "" then
		textTitle=trim(Request("textTitle"))
	else
	textTitle=""
	end if
	if trim(Request("productName")) <> "" then
		productName=trim(Request("productName"))
	else
	productName=""
	end if
	if trim(Request("usname")) <> "" then
		usname=trim(Request("usname"))
	else
	usname=""
	end if
	if trim(Request("UserN")) <> "" then
		UserN=trim(Request("UserN"))
	else
	UserN=""
	end if
		if trim(Request("UserEmail")) <> "" then
		UserEmail=trim(Request("UserEmail"))
	else
	UserEmail=""
	end if
	
	
	
		
	%>
	
					<table width="100%" cellspacing="1" cellpadding="3" align=center border="0" bgcolor="#ffffff" ID="Table4">
					<tr style="background-color: #8A8A8A">
					<td dir="<%=dir_obj_var%>" align="left"  style="padding-left: 5px;">
	 <input type="submit" value="חפש" class="but_menu" style="width: 50px" ID="btnSubmit" NAME="btnSubmit">&nbsp;&nbsp;<input type="button" 
		value="הצג הכל" class="but_menu" style="width: 60px" onclick="document.location.href='GeneralEmailByGroup.asp'" ID="btnShowAll" NAME="btnShowAll"</td>
		<td><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:90px;" value="<%=vFix(UserEmail)%>" name="UserEmail" ID="UserEmail" maxlength="50"></td>		
		<td><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:90px;" value="<%=vFix(UserN)%>" name="UserN" ID="UserN" maxlength="50"></td>
							<td align=right><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:90px;" value="<%=vFix(usname)%>" name="usname" ID="usname" maxlength="50"></td>
					<td  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:250px;" value="<%=vFix(ContactName)%>" name="ContactName" ID="ContactName" maxlength="50"></td>
		
					<td nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:250px;" value="<%=vFix(textCont)%>" name="textCont" ID="textCont" maxlength="50"></td>
					<td nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:250px;" value="<%=vFix(textTitle)%>" name="textTitle" ID="Text1" maxlength="50"></td>

					<td align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width: 250px;" value="<%=vFix(productName)%>" name="productName" ID="productName" maxlength="50"></td>
<td>&nbsp;</td>
<td>&nbsp;</td>
</tr>

							<tr>
					<td align=right  colspan=2 class="title_sort">אימייל שולח</td>
							<td align=right  class="title_sort" >מאת</td>
		
											<td width="255" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="right">עובד</td>
		<td width="15%" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="right">איש קשר</td>
							<td nowrap class="title_sort" dir="<%=dir_obj_var%>" align="right">תוכן אימייל</td>
								<td nowrap class="title_sort" dir="<%=dir_obj_var%>" align="right">נושא</td>
	
		<td width="55" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="right">שם טופס</td>
	<td width="65" nowrap align="right" dir="rtl" style="padding: 0px 3px;" class="title_sort<%if sort=1 or sort=2 then%>_act<%end if%>"><a class="title_sort" title="מיון" HREF="<%=urlSort%>&sort=<%if sort=2 then%>1<%elseif sort=1 then%>2<%else%>2<%end if%>" target="_self">תאריך<img src="../../images/arrow_<%if trim(sort)="1" then%>down<%elseif trim(sort)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td  class="title_sort">&nbsp;</td>													
	</tR>
	
	<%
  sqlstr = "exec dbo.get_Email_All_padding @Page=" & Page & ", @RecsPerPage=" & PageSize &",@prodId="& prodId & _
 ", @sender_Name='" & sFix(usname) & _
  "', @sort='" & sortby(sort) &"',@textcont='"& textcont &"',@textTitle='" &textTitle &"',@productName='" &productName &"',@UserEmail='" & UserEmail &"',@UserN='"& UserN &"'" 
'Response.Write sqlStr
'Response.End
	set app=con.GetRecordSet(sqlStr)	%>
	<%	if not app.eof then %>
<%
	
recCount = app("CountRecords")
 aa = 0
	 do while not app.eof
		If aa Mod 2 = 0 Then
			tr_bgcolor = "#E6E6E6"
		Else
			tr_bgcolor = "#C9C9C9"
		End If	
	
		
		%>
	<tr bgcolor="<%=tr_bgcolor%>" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='<%=tr_bgcolor%>';">
	<td colspan=2 align=right style="padding: 0px 3px;" valign=top dir=rtl><%if app("response_sendGroup")="0" then%><%=app("email")%><%else%><%=app("FromEmail")%><%end if%></td>
	<td align=right style="padding: 0px 3px;" valign=top dir=rtl><%if app("response_sendGroup")="0" then%><%=app("SENDER_NAME")%><%else%><%=app("FromName")%><%end if%></td>
	<td align=right style="padding: 0px 3px;" valign=top dir=rtl><%=app("SENDER_NAME")%></td>
	<td align=right style="padding: 0px 3px;" valign=top dir=rtl><%=app("contactName")%></td>
	<td align=right style="padding: 0px 3px;" valign=top dir=rtl width=100%><a class="title_sort" href="appeal_card.asp?quest_id=<%=app("QuestionId")%>&appid=<%=app("appeal_id")%>&allTasks=0"><%if len(app("response_content"))>250 then%><%=left(app("response_content"),250)%>...<%else%><%=app("response_content")%><%end if%></a><br></td>
	<td align=right style="padding: 0px 3px;" valign=top><a class="title_sort" href="appeal_card.asp?quest_id=<%=app("QuestionId")%>&appid=<%=app("appeal_id")%>&allTasks=0"><%=app("response_subject")%></a></td>
	<td align=right style="padding: 0px 3px;" valign=top dir=rtl><a class="title_sort" href="GeneralEmailByGroup.asp?Prodid=<%=app("QuestionId")%>"><%=app("ProductName")%></a></td>
	<td align=right style="padding: 0px 3px;"><%=day(app("response_date"))%>/<%=month(app("response_date"))%>/<%=mid(year(app("response_date")),3,2)%></td>
<td ><%if app("response_sendGroup")="1" then%><img src="../../images/users.gif" alt="אימייל קבוצתי"><%else%><img src="../../images/user_icon.gif"><%end if%>

</td>
	</tr>
		<%	app.movenext
		aa = aa + 1
		
		loop%>
						<%	NumberOfPages = Fix((recCount / PageSize)+0.9)	%>
								
										<% if NumberOfPages > 1 then%>
										<tr class="card">
											<td width="100%" align="center" nowrap class="card" dir="ltr" colspan=10>
												<table border="0" cellspacing="0" cellpadding="2" ID="Table6">
													<% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRow") <> nil Then
	               numOfRow = Request.QueryString("numOfRow")
	           Else numOfRow = 1
	           End If %>
													<tr>
														<%if numOfRow <> 1 then%>
														<td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow-1)-9%>&numOfRow=<%=numOfRow-1%>" name=22 title="לדפים הקודמים"><b><<</b></a></td>
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
														<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow) + 1%>&numOfRow=<%=numOfRow+1%>" name=23 title="לדפים הבאים"><b>>></b></a></td>
														<%end if%>
													</tr>
												</table>
		</table></td></tr>
<%end if%><%end if%>
 <% set app = nothing%>

					 </table>
			 </td></tr>
		</table>
</td></tr>
</form>
</table>

</body>
</html>