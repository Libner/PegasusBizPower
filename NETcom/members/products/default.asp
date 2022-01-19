<%@ Language=VBScript%>
<%ScriptTimeout=6000%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<% query_groupId = trim(Request.QueryString("groupId"))    
   prodId = Request.QueryString("prodId")
    
   If trim(query_groupId) = "" And trim(prodID) <> "" Then
		sqlStr = "Select TOP 1 GROUPID FROM PRODUCT_GROUPS WHERE PRODUCT_ID = " &_
		prodID & " AND ORGANIZATION_ID= "& OrgID & " Order BY GROUPID"
		set rs_gr = con.getRecordSet(sqlStr)
		If not rs_gr.eof Then
			query_groupId = trim(rs_gr(0))
		End If
		set rs_gr = Nothing		
	
	
	set rstmp = con.GetRecordSet("Select Count(DISTINCT PEOPLE_ID) from PRODUCT_CLIENT where GROUPID IN (" &_
	"Select GROUPID FROM PRODUCT_GROUPS WHERE PRODUCT_ID = " & prodID & ")")
		if not rstmp.eof then				
			clients = trim(rstmp(0))
		else				
			clients = 0
		end if
	set rstmp = nothing	
	
	set tmp = con.GetRecordSet("Select Count(PEOPLE_ID) from PRODUCT_CLIENT where product_id="& prodID &" and DATE_SEND IS NOT NULL and ORGANIZATION_ID=" & OrgID)
		if not tmp.eof then
			issend = true
			send = trim(tmp(0))
		else
			issend = false
			send = 0
		end if
	set tmp = nothing 
	End If 
	
	if Request("Page")<>"" then
		Page=Request("Page")
	else
		Page=1
	end if	
	
	if trim(Request.QueryString("numOfRow"))<>"" then
		numOfRow=Request.QueryString("numOfRow")
	else
		numOfRow = 1
	end if  
	found_product = false
	If isNumeric(prodId) Then
		sqlStr = "select EMAIL_SUBJECT from PRODUCTS where product_id=" & prodId & " AND ORGANIZATION_ID=" & OrgID	
	    set rs_product = con.GetRecordSet(sqlStr)	
	    If not rs_product.eof Then	
			product_name = trim(rs_product(0))
			found_product = true
		Else
			found_product = false
		End If
		set rs_product=nothing
	Else
		found_product = false	
	End if
	
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
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
var winEmail;
function check_fields(formObj) 
{
	if (checkbox_groups()==false){
	window.alert("עליך לבחור לפחות קבוצה אחת להפצה");
	return false;}
	emailLimit = "<%=emailLimit%>";
	if(isNaN(emailLimit) == false)
	{
		emailLimit = parseInt(emailLimit);
		if(emailLimit <= 0)
		{
			window.alert("לא ניתן לבצע הפצה משום שיתרתך לשליחת מיילים נגמרה \n\n 04 - לקנית יתרה נוספת ניתן להתקשר ל 8770282");
			return false;
		}
	}	
	if (window.confirm("?האם ברצונך לשלוח הפצה לנמענים שנבחרו"))
	{
		save_win = window.open("","save_win","scrollbars=1,toolbar=0,top=100,left=100,width=450,height=250,align=center,resizable=1")
		window.document.myForm.target = "save_win";
		window.document.myForm.submit();
	}
	return false;
}

function popup_Email()
{  
	winEmail = window.open("", "winEmail", "scrollbars=1,toolbar=0,top=100,left=100,width=450,height=180,align=center,resizable=1");
	//winEmail = window.showModelessDialog("",window,"dialogHeight:180px;dialogWidth:450px;dialogLeft:100px;dialogTop:100px;help:0;resizable:1")
	//winEmail.name = "winEmail";
	winEmail.focus();
	return false;
 }

function checkbox_groups()
{
	// set var checkbox_choices to zero
	var checkbox_choices = 0;
	if(myForm.groups_id.length)
	{
		// Loop from zero to the one minus the number of checkbox button selections
		for (counter = 0;  counter < myForm.groups_id.length; counter++)
		{
			// If a checkbox has been selected it will return true
			// (If not it will return false)
			if (myForm.groups_id[counter].checked)
			{ 
				checkbox_choices = checkbox_choices + 1; 
			}
		}
	}
	else
	{
		if (myForm.groups_id.checked)
		{ 
			checkbox_choices = checkbox_choices + 1; 
		}
	}


	if (checkbox_choices < 1 )
	{
		// If there were less than one selections made display an alert box 
		return (false);
	}

	if (checkbox_choices > 8 )
	{
		// If there were more than 8 selections made display an alert box 
		return (checkbox_choices);
	}

	return (true);
}


function goFormsPEOPLE(PEOPLE_ID,groupId)
{
	mywindCust = window.open("../groups_clients/addpeople.asp?PEOPLE_ID="+ PEOPLE_ID+"&groupId="+groupId,'addCONTACT',"alwaysRaised,left=100,top=100,height=350,width=450,status=no,toolbar=no,menubar=no,location=no");
	return false;
}

function openPreview(pageId,prodId)
{
	result = window.open("../Pages/result.asp?prodId="+prodId+"&pageId="+pageId,"Result","toolbar=0, menubar=0, resizable=1, scrollbars=1, fullscreen = 0, width=780, height=540, left=2, top=2");
	return false;
}
//-->
</SCRIPT>
</head>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 2%>
<%numOfLink = 2%>
<%topLevel2 = 20 'current bar ID in top submenu - added 03/10/2019%>
<!--#include file="../../top_in.asp"-->
<%If found_product Then%>
<div align="center"><center>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
<tr><td class="page_title" dir="rtl">הפצת דיוור פרסומי&nbsp;<font color="#6E6DA6"><%=product_name%></font>&nbsp;
-&nbsp;<span class="td_subj" style="color:#6E6DA6">יתרת מיילים לשליחה&nbsp;<font color=red><%=emailLimit%></font></font>&nbsp;
-&nbsp;<span class="td_subj" style="color:#6E6DA6">תאריך חידוש יתרה&nbsp;<font color=red><%=Date_curr_add%></font></font>&nbsp;
-&nbsp;<span class="td_subj" style="color:#6E6DA6">כמות חודשית&nbsp;<font color=red><%=EmailLimitMonth%></font></font>&nbsp;

</td></tr>
<tr>
<td align=center width=100% valign=top>
<table cellpadding=0 cellspacing=0 width=100% ID="Table1">
<tr><td width=100% valign=top>
<table cellspacing="0" cellpadding="0" align=center border="0" width=100%>
<FORM name="myForm" ACTION="save_product.asp" METHOD="post" enctype="multipart/form-data">								
 <tr><td width="60%" nowrap valign=top><input type=hidden name="prodID" id="prodId" value="<%=prodId%>">
 <%if trim(query_groupId) <> "" and IsNumeric(query_groupId)="True" and delGroup <> "True" then%>
	<table border=0 width=100% cellpadding=0 cellspacing=0>				
		<tr>
		<td class="title" height=22px align=right>נמענים&nbsp;</td>
		</tr>						
		<tr><td align=right>
		<table border=0 width=100% cellpadding=1 cellspacing=1>				
		<tr>							
			<td width=65 align=center class="title_sort">תאריך מענה</td>														
			<td width=65 align=center class="title_sort">תאריך הפצה</td>							
			<td width=100 align=right class="title_sort">חברה</td>														
			<td  align=right class="title_sort">Email</td>
		</tr>				
		<%					
			sqlStr = "Select PEOPLE_ID,PEOPLE_EMAIL,PEOPLE_COMPANY from PRODUCT_CLIENT Where ORGANIZATION_ID=" & OrgID &_
			" And GROUPID = " & query_groupId & " And Product_ID = " & prodID & " order by 2"	
			''Response.Write sqlStr
			set CONTACT = con.GetRecordSet(sqlStr)
			if not CONTACT.eof then
					CONTACT.PageSize = 25
					CONTACT.AbsolutePage=Page					
					recCount=CONTACT.RecordCount 
					NumberOfPages=CONTACT.PageCount
					i=1
					j=0 
					do while (not CONTACT.eof and j < CONTACT.PageSize)
					PEOPLE_ID=CONTACT("PEOPLE_ID") 
					EMAIL = CONTACT("PEOPLE_EMAIL")
					COMPANY_NAME = CONTACT("PEOPLE_COMPANY")					
				
					sqlStr = "select DATE_SEND, DATE_ANSWER from PRODUCT_CLIENT where ORGANIZATION_ID = "& OrgID &" and PRODUCT_ID = "& prodId &" and PEOPLE_ID=" & PEOPLE_ID
					set rs_PEOPLE_product = con.GetRecordSet(sqlStr)
					if not rs_PEOPLE_product.eof then
						DATE_SEND = rs_PEOPLE_product("DATE_SEND")
						DATE_ANSWER = rs_PEOPLE_product("DATE_ANSWER")
					else
						DATE_SEND = ""
						DATE_ANSWER = ""							
					end if
					set rs_PEOPLE_product = nothing
					
				%>
				<tr>							
					<td width=70 nowrap align=center class="card"><%=DATE_ANSWER%></td>							
					<td width=70 nowrap align=center class="card"><%=DATE_SEND%></td>														
					<td width=100 nowrap class="card" align=right style="padding-right:2px;<%if trim(groupId) = trim(query_groupId) then%>background:#DFDFDF<%end if%>"><a href="" ONCLICK="return goFormsPEOPLE('<%=PEOPLE_ID%>','<%=query_groupId%>');" class="linkFaq"><%=COMPANY_NAME%></a></td>							
					<td width=100% class="card" align=right style="padding-right:2px;<%if trim(groupId) = trim(query_groupId) then%>background:#DFDFDF<%end if%>"><a href="" ONCLICK="return goFormsPEOPLE('<%=PEOPLE_ID%>','<%=query_groupId%>');" class="linkFaq"><%=EMAIL%></a></td>
				</tr>					
				<%
					CONTACT.movenext
					j=j+1
					loop
				%>
				<%
					url = "default.asp?prodId=" & prodId & "&groupId=" & query_groupId
				%>
	  <%if NumberOfPages > 1 then%>
	  <tr>							
	 	 <td align=center colspan="4" class="card">
		 <table border="0" cellspacing="0" cellpadding="2">
		 <% If NumberOfPages > 10 Then 
				num = 10 : numOfRows = cInt(NumberOfPages / 10)
			Else num = NumberOfPages : numOfRows = 1    	                      
			End If
		%>
	    <tr>
		<%if numOfRow <> 1 then%> 
		<td valign="middle" align="right">
		<A class=pageCounter title="לדפים הקודמים" href="<%=url%>&page=<%=10*(numOfRow-1)-9%>&amp;numOfRow=<%=numOfRow-1%>" >&lt;&lt;</a></td>			                
		<%end if%>
		<td><font size="2" color="#001c5e">[</font></td>    
		 <%for i=1 to num
	        If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	        if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		        <td bgcolor="#e6e6e6" align="right"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	        <%else%>
	            <td bgcolor="#e6e6e6" align="right"><A class=pageCounter href="<%=url%>&page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>" ><%=i+10*(numOfRow-1)%></a></td>
	        <%end if
	        end if
	    next%>
		<td bgcolor="#e6e6e6" align="right"><font size="2" color="#001c5e">]</font></td>
		<%if NumberOfPages > cint(num * numOfRow) then%>  
		<td bgcolor="#e6e6e6" align="right"><A class=pageCounter title="לדפים הבאים" href="<%=url%>&page=<%=10*(numOfRow) + 1%>&amp;numOfRow=<%=numOfRow+1%>" >&gt;&gt;</a></td>
		<%end if%>		               
		</tr>
		<tr><td align=center colspan="11" class="card"><font style="color:#6E6DA6;font-weight:600"><%=recCount%> סה"כ נמענים</font></td></tr>
		</table></td>
	    <%end if 'recCount > res.PageSize%>				
		<%	
		  	 set CONTACT = nothing				
			 end if				
		%>
		</table></td></tr>										
		</table>
		<%end if%>				
		</td>		
		<td valign=top align="right" width=40% nowrap>
		<%'//start of groups%>
		<table border=0 width=100% nowrap cellpadding=0 cellspacing=0 align="right">				
		<tr>
		<td class="title" height=22px align=right>קבוצות&nbsp;</td>
		</tr>
		<tr><td align=right width=100%>
		<table border=0 width=100% cellpadding=1 cellspacing=1>
		<tr>
			<td width=20 nowrap align=center class="title_sort">&nbsp;</td>
			<td width=72 nowrap align=right class="title_sort">תאריך הפצה</td>							
			<td width=100% align=right class="title_sort">שם קבוצה</td>
		</tr>							
		<%				
			sqlStr = "Select GROUPID, GROUPNAME from GROUPS WHERE ORGANIZATION_ID= " & OrgID &_
			" And GROUPID IN (Select GROUPID FROM PRODUCT_GROUPS WHERE PRODUCT_ID = " & prodID &_
			" AND ORGANIZATION_ID= "& OrgID & ") Order BY GROUPID"
			'Response.Write sqlStr
			set rs_types = con.GetRecordSet(sqlStr)
			deff_all = 0
			do while not rs_types.eof
				groupId = rs_types("groupId")
				groupName = rs_types("groupName")	
				''DATE_SEND = rs_types("DATE_SEND")				
				str_groupId	= str_groupId & groupId & "; "				
				sqlStr = "select DATE_SEND from PRODUCT_GROUPS where ORGANIZATION_ID= "& OrgID &" and PRODUCT_ID=" & prodId & " and GROUPID=" & groupId
				'Response.Write sqlStr
				'Response.End
				set rs_PRODUCT_GROUP = con.GetRecordSet(sqlStr)
				if not rs_PRODUCT_GROUP.eof then
					DATE_SEND = rs_PRODUCT_GROUP("DATE_SEND")
				else
					DATE_SEND = ""
				end if
				set rs_PRODUCT_GROUP = nothing
				
				sqlstr = "Select Count(*) from PRODUCT_CLIENT  where ORGANIZATION_ID= "& OrgID &" and PRODUCT_ID=" & prodId &_
				" AND groupId="& groupId & " And DATE_SEND IS NOT NULL"
				set rsc = con.getRecordSet(sqlstr)
				If not rsc.eof Then
					count_send = trim(rsc(0))
				End If
				set rsc = Nothing
				
				sqlstr = "SELECT Count(*) FROM PRODUCT_CLIENT WHERE GROUPID ="& groupId &" and PRODUCT_ID=" & prodId & " And ORGANIZATION_ID = " & OrgID
				set rsc = con.getRecordSet(sqlstr)
				If not rsc.eof Then
					count_all = trim(rsc(0))
				End If
				set rsc = Nothing
				
				If isNumeric(count_send) Then
					 count_send = cLng(count_send)
				Else count_send = 0
				End If
				
				If isNumeric(count_all) Then
					 count_all = cLng(count_all)
				Else count_all = 0
				End If
				
				dim deff
				
				deff = count_all - count_send
				deff_all = deff_all + deff	
				
			   %>
				<tr>
					<td width=20 class="card" align=center><input type="checkbox" name="groups_id" ID="groups_id" value="<%=groupId%>" <%if deff = 0 then%>disabled<%end if%>></td>
					<td width=65 class="card" align=right style="padding-right:2px;<%if trim(groupId) = trim(query_groupId) then%>background:#9D9D9D<%end if%>"><%=DATE_SEND%></td>							
					<td class="card" align=right dir=rtl style="padding-right:2px;<%if trim(groupId) = trim(query_groupId) then%>background:#9D9D9D<%end if%>"><a href="default.asp?prodId=<%=prodId%>&groupId=<%=groupId%>" class="linkFaq"><%=groupName%></a></td>
				</tr>					
				<%
				rs_types.movenext
				loop
				set rs_types = nothing
				%>
				</table></td></tr>			
			</table>
			<%'//end of groups%>			
		</td>	
	</tr>
</Form>
</table>
</td>
<td width=100 nowrap align=right valign=top>
<table cellpadding=0 cellspacing=0 width=100%>
<tr><td align="right" colspan=2 height="23" nowrap></td></tr>
<tr><td bgcolor=#C9C9C9 align="right" colspan=2 height="15" nowrap></td></tr>
<tr><td bgcolor=#C9C9C9 nowrap colspan=2 align="center"><a class="button_edit_1" style="width:93;" href="#" OnClick="return openPreview('<%=pageId%>','<%=prodId%>')">תצוגה מקדימה</a></td></tr>
<%If deff_all > 0 Then%>
<tr><td bgcolor=#C9C9C9 nowrap colspan=2 align="center"><a class="button_edit_1" style="width:93;" href="#" OnClick="return check_fields(window.document.forms(0));">שלח הפצה</a></td></tr>
<%End if%>
<tr><td bgcolor=#C9C9C9 align="right" colspan=2 height="15" nowrap></td></tr>
</table></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table>
<%End If%>
</body>
<%set con=nothing%>
</html>
