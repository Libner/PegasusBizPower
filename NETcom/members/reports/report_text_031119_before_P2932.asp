<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
	quest_id = Request.Form("quest_id")
	fId = Request.Form("FieldId")
	start_date = trim(Request("dateStart"))		
    end_date = trim(Request("dateEnd"))		
    start_date_ = Month(start_date) & "/" & Day(start_date) & "/" & Year(start_date)   
    end_date_ = Month(end_date) & "/" & Day(end_date) & "/" & Year(end_date) %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--
	function showcard(Id,gID)
	{
		strOpen="../groups_clients/addpeople.asp?people_Id="+Id+"&groupId="+gID;
		oClient = window.open(strOpen, "Client", "scrollbars=1,toolbar=0,top=50,left=150,width=450,height=350,align=center,resizable=0");
		oClient.focus();
		return false;
	}
//-->
</script>
</head>
<body style="margin: 0px" >
<div align="center">
<table border="0" width="780" cellspacing="0" cellpadding="0" align=center ID="Table1">
<tr>
<td width="780" align="center">
<!--#INCLUDE FILE="logo_top.asp"-->
</td></tr> 
</table>
<table border="0" width="780" cellspacing="0" cellpadding="0" align=center ID="Table2">
<!-- start code --> 
	<tr bgcolor="#FF9900">
		<td align="center" height="20" style="color:#000000; font-size:13pt" align=right valign=bottom dir="rtl"><b>			
         <%If trim(lang_id) = "1" Then%>דוח תשובות לשאלה חופשית<%Else%>Open text - Answer report<%End If%></b></td>
	</tr>				

  <tr>
    <td width="100%" align="center">
	<table WIDTH="100%" BGCOLOR="#999999" ALIGN="center" BORDER="0" CELLPADDING="1" cellspacing="0">	
			<tr>
			<td align="right" width="100%" valign="top">
<%				set prod = con.GetRecordSet("Select Product_Name, Langu, Is_Archive FROM PRODUCTS WHERE PRODUCT_ID=" & quest_id & " AND ORGANIZATION_ID=" & orgID)
					if not prod.eof then
					productName = trim(prod("Product_Name"))
					IsArchive = trim(prod("Is_Archive"))
					if prod("Langu") = "eng" then
						td_align = "left"
						pr_language = "eng"
					else
						td_align = "right"
						pr_language = "heb"
					end if
				end if
				set prod = nothing%> 			 
			<table border="0" cellpadding="0" cellspacing="0" width="100%" align="right" bgcolor="#F0F0F0">
			<tr><td class="form" align=center width=100% bgcolor=#999999 colspan=4><INPUT type="hidden" id=product_id name=product_id value="<%=prod_id%>">
					<table width="100%" cellpadding="0" cellspacing="0">
						<tr>
							<td align="center">
							<font color=#FFFFFF size=3 dir="rtl"><b><%=productName%></b></font>
							</td>
						</tr>
						<tr>
						   <td align="center"><font color=#FFFFFF size=3 dir="rtl"><b><%=end_date & " - " & start_date%></b></font></td>
						</tr>
					</table>		
				</td>
			</tr>
		<tr><td width="100%" height=10></td></tr>	
	<%
	sqlStr = "SET DATEFORMAT MDY; SELECT  PC.PEOPLE_ID, PC.PEOPLE_EMAIL, PC.GROUPID, dbo.FORM_VALUE.FIELD_ID, dbo.FORM_VALUE.FIELD_VALUE, " &_
	" dbo.FORM_FIELD.FIELD_TYPE, dbo.FORM_FIELD.FIELD_SIZE, dbo.FORM_FIELD.FIELD_ALIGN, dbo.FORM_FIELD.FIELD_TITLE, dbo.APPEALS.APPEAL_DATE "&_
	" FROM dbo.APPEALS INNER JOIN "
	If IsArchive = "1" Then
			sqlStr = sqlStr & " dbo.PRODUCT_CLIENT PC "
	Else
		    sqlStr = sqlStr & " dbo.PRODUCT_CLIENT_ARCH PC "
	End If
	 sqlStr = sqlStr & " ON dbo.APPEALS.PRODUCT_ID = PC.PRODUCT_ID AND " &_
	" dbo.APPEALS.PEOPLE_ID = PC.PEOPLE_ID INNER JOIN " &_
	" dbo.FORM_VALUE ON dbo.APPEALS.APPEAL_ID = dbo.FORM_VALUE.APPEAL_ID INNER JOIN " &_
	" dbo.FORM_FIELD ON dbo.FORM_VALUE.FIELD_ID = dbo.FORM_FIELD.FIELD_ID " &_
	" AND dbo.FORM_VALUE.FIELD_ID = " & fId & " AND dbo.APPEALS.QUESTIONS_ID = " & quest_id &_
	" AND dbo.APPEALS.ORGANIZATION_ID = " & OrgID
	If isDate(start_date) Then
		sql = sql & " And DateDiff(d,APPEAL_DATE,'"&start_date_&"') <= 0"
	End If	
    If isDate(end_date) Then
	    sql = sql & " And DateDiff(d,APPEAL_DATE,'"&end_date_&"') >= 0"			
    End If 
	'Response.Write sqlStr
	'Response.End	 	    
	set fld = con.GetRecordSet(sqlStr)
	if not fld.eof then %>
	<tr>
		<td>
		<TABLE align=center BORDER=0 CELLSPACING=0 CELLPADDING=0 width=100%>
		<tr height=20>
			<tD valign="middle" width="100%" style="font-size:10pt;font-weight:bold" align=<%=td_align%> valign="top"  bgcolor=#DADADA dir="rtl">&nbsp;<%=fld("Field_Title")%>&nbsp;</tD>
		</tr>	
		</table>
		</td>
	</tr>
	 <% 
	 Do while not fld.eof
		Field_Value=fld("Field_Value")
		FIELD_TYPE=fld("FIELD_TYPE")
		if FIELD_TYPE = "5" then
			if pr_language = "heb"  then 
				Field_Value = "כן"
			else
				Field_Value = "Yes"			
			end if	
		end if
		Field_Align=fld("Field_Align")
		CLIENT_EMAIL=fld("PEOPLE_EMAIL")
		Client_Id=fld("PEOPLE_ID")	  
		GROUPID=fld("GROUPID")	  
	  
		If trim(Field_Value) <> "" Then	  
	%>
		<tr>
		  <td width="100%">
		  <TABLE align=center BORDER=0 CELLSPACING=1 CELLPADDING=1 width=100%>
			<tr>
			<%if pr_language = "heb" then%>
			    <td class="form" width="20%" align=left valign="top" nowrap><a class=linkFaq href="" onclick="return showcard(<%=client_Id%>,<%=GROUPID%>)"><%=CLIENT_EMAIL%></a></td>
				<td class="form" width="80%" align=right valign="top" dir="rtl"><%=Field_Value%></td>
			<%else%>	
				<td class="form" width="80%" align=left valign="top" dir="rtl"><%=Field_Value%></td>
				<td class="form" width="20%" align=center valign="top" nowrap><a class=linkFaq href="" onclick="return showcard(<%=client_Id%>,<%=GROUPID%>)" target=_self><%=CLIENT_EMAIL%></a></td>
			<%end if%>
	         </tr>			
	      </table>      
				
		</tr>
		<tr>
			   <td width="100%" height="1" bgcolor="#DADADA"></td>
		</tr>		
		<tr>
			   <td width="100%" height="5" nowrap></td>
		</tr>
	<%
		end if
	fld.MoveNext
	loop
	end if 'not fld.eof
	set fld = nothing%>
<!-- end code --> 

</table>
</td></tr>
</table>
</td></tr>

	<tr><td width="100%" height=5></td></tr>		
	
	<!--#INCLUDE FILE="bottom_inc.asp"-->
</table>

</div>
</body>
<%  set con = nothing %>
</html>
