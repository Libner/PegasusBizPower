<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css" />
<script LANGUAGE="JavaScript">
<!--
function chkDeleteSMS() {
  return (confirm("?האם ברצונך למחוק את ההפצה"))    
}

function chkSmsLimit(smsLimit, IsOkToSend)
{	
	if(isNaN(smsLimit) == false)
	{
		smsLimit = parseInt(smsLimit);
		if(smsLimit <= 0)
		{
			window.alert("לא ניתן לבצע הפצה משום שיתרתך לשליחת הודעות נגמרה \n\n 04 - לקנית יתרה נוספת ניתן להתקשר ל 8770282");
			return false;
		}
	}
	return true;
}
//-->
</script>  
</head>
<%delSmsId = trim(Request.QueryString("delSmsId"))
	If IsNumeric(delSmsId) And trim(delSmsId) <> "" Then		
		con.ExecuteQuery("DELETE FROM sms_to_people WHERE sms_id="&delSmsId & " AND ORGANIZATION_ID=" & trim(OrgID) )
		con.ExecuteQuery("DELETE FROM sms_sends WHERE sms_id="&delSmsId & " AND ORGANIZATION_ID=" & trim(OrgID) )		
	End If
	
	'------------------------------------ Check organization sms limit  ---------------------		
	sql = "SELECT Sms_Limit FROM ORGANIZATIONS WHERE ORGANIZATION_ID = " & OrgID
	set rs_org = con.getRecordSet(sql)	
	If not rs_org.eof Then
		smsLimit = trim(rs_org(0))
	End If
	set rs_org = Nothing	
	
	If isNumeric(smsLimit) Then
		smsLimit = cLng(smsLimit)
	End If
	'------------------------------------ End of Check organization sms limit  ---------------------			
	
	dim sortby(6)	
	sortby(0) = "sms_id DESC"
	sortby(1) = "sms_date"
	sortby(2) = "sms_date DESC"
	sortby(3) = "count_peoples"
	sortby(4) = "count_peoples DESC"
	sortby(5) = "sms_desc"
	sortby(6) = "sms_desc DESC"	
	
	sort = Request("sort")	
	if sort = nil then
		sort = 0
	end if
	urlSort = "default.asp?1=1"	%>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 46%>
<%numOfLink = 1%>
<%topLevel2 = 49 'current bar ID in top submenu - added 03/10/2019%>
<!--#include file="../../top_in.asp"-->
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
<tr><td class="page_title" dir="<%=dir_obj_var%>">&nbsp;הפצות SMS&nbsp;-&nbsp;<span class="td_subj" style="color:#6E6DA6">יתרת sms לשליחה&nbsp;<font color=red><%=smsLimit%></font></font>&nbsp;</td></tr>
<tr><td align=center valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0>
<tr>
<td width=100% valign=top>
	<table width="100%" cellspacing="1" cellpadding="1" align=center border="0" bgcolor="#ffffff" dir="<%=dir_var%>">
	<tr>
		<td valign=top class="title_sort" align=center>מחיקה</td>
		<td valign=top nowrap align=center class="title_sort<%if sort=1 or sort=2 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=2 then%>1<%elseif sort=1 then%>2<%else%>2<%end if%>" target="_self">תאריך הפצה<img src="../../images/arrow_<%if trim(sort)="1" then%>down<%elseif trim(sort)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>		
		<td width="60" nowrap valign=top align=center class="title_sort<%if sort=3 or sort=4 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=4 then%>3<%elseif sort=3 then%>4<%else%>4<%end if%>" target="_self">נשלחו<img src="../../images/arrow_<%if trim(sort)="3" then%>down<%elseif trim(sort)="4" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	
		<td width="380" valign=top nowrap class="title_sort" align="<%=align_var%>">&nbsp;תוכן sms&nbsp;</td>
		<td width="100%" align="<%=align_var%>" dir="<%=dir_var%>" valign=top class="title_sort<%if sort=5 or sort=6 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=6 then%>5<%elseif sort=5 then%>6<%else%>6<%end if%>" target="_self">שם הפצה&nbsp;<img src="../../images/arrow_<%if trim(sort)="5" then%>down<%elseif trim(sort)="6" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>		
		<td width="80" nowrap valign=top class="title_sort" align=center>סטטוס</td>
	</tr>
<%
	Dim arrStatus(4) 
	arrStatus(0) = "טיוטה" : arrStatus(1) = "מתוזמן לשליחה" : arrStatus(2) = "בתהליך שליחה" : arrStatus(3) =  "שליחה הסתיימה"
	
    If IsNumeric(Request("Page")) And trim(Request("Page")) <> "" Then
		Page = cLng(Request("Page"))
	ELse
		Page=1
	End If
	sms_count = 0
	sqlstr = "SELECT sms_id, sms_desc, sms_text, sms_date, error, status, (SELECT COUNT(people_id) FROM sms_to_people STP " & _
	" WHERE (S.sms_id = STP.sms_id) AND (date_send is Not NULL)) as count_peoples " & _
	" FROM sms_sends S WHERE organization_id = " & trim(OrgID) & " ORDER BY " & sortby(sort)
	'Response.Write sqlStr
	set rs_sms = con.getRecordSet(sqlstr)
    if not rs_sms.eof then
		If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
			PageSize = RowsInList
		Else	
     		PageSize = 10
		End If	   
	    rs_sms.PageSize = PageSize
		rs_sms.AbsolutePage = Page
		recCount = rs_sms.RecordCount 	
		sms_count = rs_sms.RecordCount	
		NumberOfPages = rs_sms.PageCount
		i=1
		j=0
		do while (not rs_sms.eof and j<rs_sms.PageSize)
    	smsId = trim(rs_sms("sms_id"))
    	sms_desc = rs_sms("sms_desc")
    	If trim(sms_desc) = "" Or IsNull(sms_desc) Then
    		sms_desc = "אין"
    	End if
    	If IsDate(rs_sms("sms_date")) Then
    		sms_date= Day(rs_sms("sms_date")) & "/" & Month(rs_sms("sms_date")) & "/" & Year(rs_sms("sms_date"))
    	Else
    		sms_date=""
    	End If	
    	sms_text = trim(rs_sms("sms_text"))  
		count_peoples = trim(rs_sms("count_peoples"))
		If IsNull(count_peoples) = false And IsNumeric(count_peoples) = true Then
    		If cLng(count_peoples) > 0 Then
    		isclients = true
			peoples = cLng(count_peoples)
			Else
			isclients = false
			peoples = 0
			End If
		else
			isclients = false
			peoples = 0
		end if		
		error = trim(rs_sms("error"))
		status = trim(rs_sms("status"))
		sms_status = ""
		If Len(error) > 0 Then
			 sms_status = error
		Else
			  sms_status = arrStatus(status)		
		End If		%>
<tr>
	<td align="center" class="card" nowrap>&nbsp;<a href="default.asp?delSmsId=<%=smsId%>" onclick="return chkDeleteSMS();"><IMG SRC="../../images/delete_icon.gif" BORDER=0 alt="מחיקת הפצה"></a>&nbsp;</td>
	<td align="center" class="card" nowrap>&nbsp;<%=sms_date%>&nbsp;</td>
	<td align="center" class="card" nowrap><a href="javascript:void(0);" OnClick="javascript:window.open('peoples.asp?smsId=<%=smsId%>','peoplesList','left=50,top=50,scrollbars=1,width=650,height=500');" class="link1"><%=peoples%></a></td>
	<td align="<%=align_var%>" class="card" dir="rtl">&nbsp;<%=sms_text%>&nbsp;</td>
	<td align="<%=align_var%>" class="card" dir="rtl">&nbsp;<%=sms_desc%>&nbsp;</td>
	<td align="center" class="card" dir="rtl"><%=sms_status%></td>
</tr>
<%		rs_sms.MoveNext
        j=j+1
		loop
	end if
	set rs_sms=nothing%>
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
	<tr><td align=center height=20px class="card" colspan="13"><font style="color:#6E6DA6;font-weight:600"><%=sms_count%> : סה"כ הפצות</font></td></tr>								
</table></td>
<td width=90 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100%>
<tr><td align="<%=align_var%>" colspan="2" height="15" nowrap></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:84;line-height:11pt;padding:2pt;" href='addsms.asp' onclick="return chkSmsLimit('<%=smsLimit%>','<%=IsOkToSend%>')">sms שלח</a></td></tr>
<tr><td align="<%=align_var%>" colspan="2" height="10" nowrap></td></tr>
</table></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr></table>
</body>
</html>
<%set con=nothing%>