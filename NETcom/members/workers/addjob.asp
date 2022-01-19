<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<%
    If Request.QueryString("add") <> nil Then
		If trim(Request.Form("job_id")) = "" Then ' add type
		'	sqlstr = "Insert into jobs (Organization_ID,job_name,hour_pay) values (" &_
		'	trim(Request.Cookies("bizpegasus")("OrgID")) & ",'" &_
		'	sFix(Request.Form("job_name")) & "','" & sFix(Request.Form("hour_pay")) & "')"
			
					sqlstr = "SET NOCOUNT ON;Insert into jobs (Organization_ID,job_name) values (" &_
			trim(Request.Cookies("bizpegasus")("OrgID")) & ",'" &_
			sFix(Request.Form("job_name"))  & "'); SELECT @@IDENTITY AS NewID"
	'	response.Write sqlstr
	'	response.end 
	set rs_tmp = con.getRecordSet(sqlStr)
			'con.executeQuery(sqlstr) 
				job_id = rs_tmp.Fields("NewID").value
					set rs_tmp = Nothing	
				
		'	response.Write job_id
			%>		
		
	<%	Else ' update type
			job_id = trim(Request.Form("job_id"))
			sqlstr="Update jobs set job_name = '" & sFix(Request.Form("job_name")) & "' Where job_id = " & job_id
			con.executeQuery(sqlstr) %>
		
	<%	End If
'	response.Write "job_id="&job_id
	
		sqlstr = "DELETE FROM bar_jobs WHERE organization_id = " & OrgID & " AND job_id = " & job_id
		'Response.Write sqlstr
		'Response.End
		con.executeQuery(sqlstr)
				
		sqlstr="SELECT bar_id FROM bars ORDER BY PARENT_ID, bar_Order"
		SET rs_bars = con.getRecordSet(sqlstr)
		WHILE NOT rs_bars.eof
			barID = trim(rs_bars(0))
			barVisible = Request.Form("is_visible"&barID)
			If trim(barVisible) = "on" Then
				barVisible = 1
			Else
				barVisible = 0
			End If	
		
			sqlstr = "INSERT INTO bar_jobs  ([bar_id],[organization_id],[job_id],[is_visible]) VALUES (" & _
			barID & "," & OrgID & "," & job_id & ",'" & barVisible & "')"
		'	response.Write sqlstr
		'	response.end
			con.executeQuery(sqlstr)
		rs_bars.moveNext
		WEND
		SET rs_bars = Nothing   %>
		<script language="javascript" type="text/javascript">
		<!--
			opener.focus();
			opener.window.location.reload();
			self.close();
		//-->
		</script>			
		<%	
	End If

	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 37 Order By word_id"				
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
	  
	sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
	set rsbuttons = con.getRecordSet(sqlstr)
	If not rsbuttons.eof Then
	arrButtons = ";" & rsbuttons.getString(,,",",";")		
	arrButtons = Split(arrButtons, ";")
	End If
	set rsbuttons=nothing %>
<%
	If Request.QueryString("job_id") <> nil Then
		job_id = trim(Request.QueryString("job_id"))
		If Len(job_id) > 0 Then
			sqlstr="Select job_name, hour_pay From jobs Where job_id = " & job_id
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				job_name = trim(rssub("job_name"))	
				hour_pay = trim(rssub("hour_pay"))	
			End If
			set rssub=Nothing			
		End If
	End If	
%>	  
<script language="javascript" type="text/javascript">
<!--
		function checkForm()
		{
			if(window.document.form1.job_name.value == "")
			{
				<%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס סוג עובד"
				Else
					str_alert = "Please insert the employee type !!"
				End If   
				%>			
				window.alert("<%=str_alert%>");
				window.document.form1.job_name.focus();
				return false;
			}
			if(window.document.form1.hour_pay.value == "")
			{
				<%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס יעד הזמנות מינימלי"
				Else
					str_alert = "Please insert the hour cost !!"
				End If   
				%>					
				window.alert("<%=str_alert%>");
				window.document.form1.hour_pay.focus();
				return false;
			}
				
			return true;
			   
		}
		
		function GetNumbers ()
		{
			var ch=event.keyCode;
			event.returnValue =(ch >= 48 && ch <= 57) || ch == 46 ;
		} 
	
		function check_all_bars(objChk,parentID)
		{
			input_arr = document.getElementsByTagName("input");	
			for(i=0;i<input_arr.length;i++)	{
				
				if(input_arr(i).type == "checkbox")
				{
					currparentId = "";
					objValue = new String(input_arr(i).id);			
					value_arr = objValue.split("!");
					currparentId =  value_arr[1];
					if(currparentId == parentID)
					{
						//input_arr(i).disabled = objChk.checked;
						input_arr(i).checked = objChk.checked;
					}	
				}	
			}
			return true;
		}	
//-->
</script>
</head>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">
<table border="0" width="480" cellspacing="0" cellpadding="0" align="center" dir="<%=dir_var%>">
<tr >
   <td  align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="1" cellspacing="0">
	   <tr>		 
		 <td class="page_title" align="<%=align_var%>" valign="middle" width="100%"><%If Len(type_id) > 0 Then%><!--עדכון--><%=arrTitles(1)%><%Else%><!--הוספת--><%=arrTitles(2)%><%End If%>&nbsp;<!--סוג עובד--><%=arrTitles(4)%>&nbsp;</td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width="100%">
<form name="form1" id="form1" action="addjob.asp?add=1" target="_self" method="post">
<input type=hidden name=job_id id=job_id value="<%=job_id%>">
<table width="480" cellspacing="1" cellpadding="2" align="center" border="0">
<tr>
	<td align="<%=align_var%>" width="380" nowrap>
	<input type="text" name="job_name" id="job_name" value="<%=vFix(job_name)%>" dir="<%=dir_obj_var%>" size=40 maxLength=50>	
	</td>
	<td width="100" nowrap align="<%=align_var%>">&nbsp;<b><!--סוג עובד--><%=arrTitles(3)%></b>&nbsp;</td>	
</tr>
<tr><td height="5" nowrap colspan="2"></td></tr>
<tr><td class="title" align="<%=align_var%>" colspan="2">&nbsp;&nbsp;הרשאות&nbsp;&nbsp;</td></tr>
<tr><td colspan="2" align="center">
<table cellpadding="0" cellspacing="0" width="96%" border="0" align="center">
<% sql_obj="SELECT bar_id, bar_title, parent_id, parent_title FROM dbo.bar_organizations_table('"&OrgID&"') "&_
	" WHERE (PARENT_ID IS NOT NULL) AND (IS_VISIBLE = '1') ORDER BY Parent_Order, bar_Order"
	set rs_obj=con.getRecordSet(sql_obj)
	While not rs_obj.eof
	barID = trim(rs_obj(0))
	barTitle = trim(rs_obj(1))
	barParent = trim(rs_obj(2))
	parentTitle = trim(rs_obj(3))
	If trim(barID) = "43" Then
		barTitle = barTitle & " <font color=red>(הרשאות טפסים)</font>"
	End If 	

	If Len(job_id) > 0 then	   	
		sql_obj="SELECT is_visible FROM dbo.bar_jobs WHERE (organization_id = '" & OrgID & "') AND (job_id = '" & job_id & "') " & _
		" AND (bar_id = " & barID & ")"
		'Response.Write sql_obj
		set rs_visible=con.getRecordSet(sql_obj)
		If not rs_visible.eof Then		
			barVisible = rs_visible("is_visible")
		Else
			barVisible = 0
		End If
		set rs_visible = Nothing
	Else
		barVisible = 1
	End If

    If isNumeric(barParent) Then
	If old_parent <> parentTitle Then
	If Len(job_id) > 0 Then
		sql_obj="SELECT is_visible FROM dbo.bar_jobs WHERE (organization_id = '" & OrgID & "') AND (job_id = '" & job_id & "') " & _
		" AND (bar_id = " & barParent & ")"
		'Response.Write sql_obj
		'Response.End
		set rs_parent=con.getRecordSet(sql_obj)
		If not rs_parent.eof Then					
			parentVisible = trim(rs_parent(0))			
		End If
		set rs_parent = Nothing		
	Else
		parentVisible = 1	
	End If	%>
<tr>
	<th class="form_title" align="<%=align_var%>"><input type="checkbox" dir="ltr" name="is_visible<%=barParent%>" <%If trim(parentVisible) = "1" Then%> checked <%End If%> ID="is_visible<%=barParent%>" onclick="return check_all_bars(this,'<%=barParent%>')"></th>
	<th align="<%=align_var%>" class="form_title" nowrap width="200" dir="<%=dir_obj_var%>">&nbsp;<%=parentTitle%>&nbsp;</th>
</tr>
<%End If	
	End If%>
<tr>
	<td align="<%=align_var%>"><input type="checkbox" dir="ltr" name="is_visible<%=barID%>" id="<%=barID%>!<%=barParent%>"  <%If trim(barVisible) = "1" Then%> checked <%End If%>></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="200" dir="<%=dir_obj_var%>"><%=barTitle%>&nbsp;</th>
</tr>
<% old_parent = parentTitle
   rs_obj.moveNext
   Wend
   set  rs_obj = Nothing%>
</table></td></tr>
<tr><td height="5" nowrap colspan="2"></td></tr>
<tr><td height="1" nowrap colspan="2" bgcolor="#808080" ></td></tr>
<tr><td height="10" nowrap colspan="2"></td></tr>
<tr><td align="center" colspan="2">
<input type="button" value="<%=arrButtons(2)%>" class="but_menu" style="width:90" onclick="window.close();" id=Button2 name=Button2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="submit" value="<%=arrButtons(1)%>" class="but_menu" style="width:90" onclick="return checkForm()" id=Button1 name=Button1></td></tr>
</table>
</form>
</td></tr></table>
</BODY>
</HTML>
<%Set con = Nothing%>