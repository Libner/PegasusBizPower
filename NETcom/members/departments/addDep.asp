<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<% 
    If Request.QueryString("add") <> nil Then
		If trim(Request.Form("DepartmentId")) = "" Then ' add type
			sqlstr = "Insert into Departments (DepartmentName,PriorityLevel) values ('"  & sFix(Request.Form("DepartmentName")) & "',"& Fix(Request.Form("PriorityLevel")) &") "
		'   response.Write sqlstr
		'   response.end
		   con.executeQuery(sqlstr)
		   sql="SELECT top 1 DepartmentId  from Departments  order by DepartmentId desc"
		set rs_tmp = con.getRecordSet(sql)
		if not rs_tmp.eof then
				DepartmentId = rs_tmp("DepartmentId")
			end if
			
		set rs_tmp = Nothing
		
		
				
			 %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				opener.focus();
				opener.window.location.reload();
				self.close();
			//-->
			</SCRIPT>
		
	<%	Else ' update type
			DepartmentId = trim(Request.Form("DepartmentId"))
			sqlstr="Update Departments set DepartmentName = '" & sFix(Request.Form("DepartmentName")) & "' Where DepartmentId = " & DepartmentId
			con.executeQuery(sqlstr)
		
			
	
				 %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				opener.focus();
				opener.window.location.reload();
				self.close();
			//-->
			</SCRIPT>	
	<%	End If
	End If
	
%>

<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
		function checkForm()
		{
		
			if(window.document.form1.DepartmentName.value == "")
			{
			
				<%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס מחלקה "
				Else
					str_alert = "Please insert the group name !!"
				End If   
				%>
				window.alert("<%=str_alert%>");
				window.document.form1.DepartmentName.focus();
				return false;
			}
		//	alert(window.document.form1.task_types.length)
	
//-->
</SCRIPT>
</HEAD>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<form name=form1 id=form1 action="addDep.asp?add=1" target="_self" method="post">

<%
	If Request.QueryString("DepartmentId") <> nil Then
		DepartmentId = trim(Request.QueryString("DepartmentId"))
		If Len(DepartmentId) > 0 Then
			sqlstr="Select DepartmentName,PriorityLevel From Departments Where DepartmentId = " & DepartmentId
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				DepartmentName = trim(rssub("DepartmentName"))	
				PriorityLevel =	trim(rssub("PriorityLevel"))			
			End If
			set rssub=Nothing
		End If
	else
			sqlstr = "select top 1 PriorityLevel   from Departments order by PriorityLevel desc "
            set rssub = con.getRecordSet(sqlstr)
            if not rssub.eof Then
            PriorityLevel=	trim(rssub("PriorityLevel"))+1
            else
           PriorityLevel=1
            end if
	End If	
%>
<table dir="<%=dir_var%>" border="0" width="100%" cellspacing="0" cellpadding="0" align="center" ID="Table1">
<tr >
   <td  align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="1" cellspacing="0" ID="Table2">
	   <tr>	
	   <td class="page_title" align=left>
<input type=button value="ביטול" class="but_menu" style="width:90" onclick="window.close();" id=Button2 name=Button2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=submit value="אישור" class="but_menu" style="width:90" onclick="return checkForm()" id=Button1 name=Button1></td>	 
		 <td class="page_title" align="<%=align_var%>" valign="middle" width="100%"><%If Len(DepartmentId) > 0 Then%><span id=word1 name=word1>עדכון מחלקה</span><%Else%><span id="word2" name=word2>הוספת מחלקה</span><%End If%>&nbsp;</td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width=100%>
<table width="480" cellspacing="1" cellpadding="2" align=center border="0" ID="Table3">
<input type=hidden name=DepartmentId id=DepartmentId value="<%=DepartmentId%>">
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" class="texts" name="DepartmentName" id="DepartmentName" value="<%=vFix(DepartmentName)%>" dir="<%=dir_obj_var%>" size=100 maxLength=100>	
	</td>
	<td width="150" nowrap align="center">&nbsp;<b><span id=word3 name=word3>מחלקה</span></b>&nbsp;</td>	
</tr>
<tr><td align="<%=align_var%>" width=330 nowrap><input type="text" name="PriorityLevel" id="PriorityLevel" value="<%=vFix(PriorityLevel)%>" dir="<%=dir_obj_var%>" size=2 maxLength=3></td>		

<td  width="150" nowrap align="center">&nbsp;<b><span id="Span1" name=word3> רמת עדיפות</span></b>&nbsp;</td>
</tr>

</table>
								</td>
							
							</tr>
							<tr>
								<td bgcolor="#e6e6e6" height="10"></td>
							</tr>
<!--site country-->
<tr><td height=5 colspan="2" nowrap></td></tr>
</form>
</table>
</td></tr> </table>
</BODY>
</HTML>
