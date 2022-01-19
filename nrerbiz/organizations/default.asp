<!--#include file="../../netcom/connect.asp"-->
<!--#include file="../../netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
</head>
<script LANGUAGE="JavaScript">
<!--
function CheckDel(str) {
  return (confirm("? האם ברצונך למחוק את הארגון"))    
}
//-->
</script>
<%
    Session.LCID = "1037"
    Response.CharSet = "windows-1255"

	sort = Request.QueryString("sort")	
	if trim(sort)="" then  sort = 0  end if
	
	dim sortby(5)	
	sortby(0) = " ORGANIZATION_ID"
	sortby(1) = " rtrim(ltrim(ORGANIZATION_NAME))"
	sortby(2) = " rtrim(ltrim(ORGANIZATION_NAME)) DESC"	
	sortby(3) = " rtrim(ltrim(IN_WORK)), rtrim(ltrim(ORGANIZATION_NAME))"
	sortby(4) = " rtrim(ltrim(IN_WORK)) DESC, rtrim(ltrim(ORGANIZATION_NAME))"

    if Request.QueryString("delORGANIZATION_ID")<>nil then	
        str_dir = "../../"
    
        delOrgID = trim(Request.QueryString("delORGANIZATION_ID"))
        'start delete all organization files
		sqlstr = "SELECT attachment, attachment_closing FROM tasks WHERE ORGANIZATION_ID = " & delOrgID & " And ( Len(IsNULL(attachment, '')) > 0 OR Len(IsNULL(attachment_closing, '')) > 0 ) "
		set rs_check = con.getRecordSet(sqlstr)
		While Not rs_check.eof 
		    attachment = trim(rs_check(0))
		    attachment_closing = trim(rs_check(1)) 
		    If Len(attachment) > 0 Then
   				DeleteFile(str_dir & "download/tasks_attachments/" & attachment)
   			End If
   			If Len(attachment_closing) > 0 Then	
   				DeleteFile(str_dir & "download/tasks_attachments/" & attachment_closing)
   			End If	
			rs_check.moveNext
		Wend
		set rs_check = nothing
		
		sqlstr = "SELECT document_file FROM company_documents_view WHERE ORGANIZATION_ID = " & delOrgID & " And Len(IsNULL(document_file, '')) > 0 "
		set rs_check = con.getRecordSet(sqlstr)
		While Not rs_check.eof 
		    document_file = trim(rs_check(0))
   			DeleteFile(str_dir & "download/documents/" & document_file)
			rs_check.moveNext
		Wend
		set rs_check = nothing
		
		sqlstr = "SELECT FILE_ATTACHMENT FROM products WHERE ORGANIZATION_ID = " & delOrgID & " And Len(IsNULL(FILE_ATTACHMENT, '')) > 0"
		set rs_check = con.getRecordSet(sqlstr)
		While NOT rs_check.eof 
		    file_attachment = trim(rs_check(0))
   			DeleteFile(str_dir & "download/products/" & file_attachment)
			rs_check.moveNext
		Wend
		set rs_check = nothing		
		
		con.ExecuteQuery "Exec dbo.delete_organization @delOrgId=" & trim(delOrgID)
		
		Response.Redirect "default.asp?sort=" & sort
end if

Sub DeleteFile(file_path)
		Set fs=Server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object! 
   		'Response.Write fs.FileExists(server.mappath(file_path))
		'Response.End
		If fs.FileExists(Server.MapPath(file_path)) Then
			Set a = fs.GetFile(Server.MapPath(file_path))
			a.delete		
		Else		
			Response.Write Server.MapPath(file_path) & "<br>"
		End If	
		Set fs=Nothing
End Sub

If Request.QueryString("update_work") <> nil Then
	OrgId = trim(Request.QueryString("OrgID"))
	sqlstr = "Update ORGANIZATIONS Set in_work = 1 - IsNull(IN_WORK,0) WHERE ORGANIZATION_ID = " & OrgID
	con.executeQuery(sqlstr)
	Response.Redirect "default.asp?sort=" & sort
End If
%>
<body bgcolor="#FFFFFF">
<div align="right">
<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="4" class="page_title" width=100%>דף ניהול ארגונים</td>
  </tr>
  <tr>
     <td width="5%" align="center"><a class="button_admin_1" href="addOrg.asp">הוספת ארגון חדש</a></td>    
     <td width="5%" align="center"><a class="button_admin_1" href="../../nrerbiz/choose.asp">חזרה לדף ניהול ראשי</a></td>  
     <td width="5%" align="center"></td> 
     <td width="*%" align="center"></td>      
  </tr>
</table>
<table width="100%" align="center" border="0" cellpadding="1" cellspacing="1">
<tr bgcolor="#B4B4B4" height="22">
   <td class="11normalB" nowrap align="center">&nbsp;מחיקה&nbsp;</td>   
   <td class="11normalB" nowrap align="center">&nbsp;פעיל&nbsp;</td>
   <td class="11normalB" nowrap align="center">&nbsp;קבצים&nbsp;</td>
   <td class="11normalB" nowrap align="center">&nbsp;הרשאות&nbsp;</td>
   <td class="11normalB" nowrap align="center">&nbsp;עובדים&nbsp;</td>   
   <td class="11normalB" nowrap align="center">&nbsp;תפקידים&nbsp;</td>
   <td class="11normalB" nowrap align="center">&nbsp;קבוצות&nbsp;</td>
   <td class="11normalB" nowrap align="center">&nbsp;מיילים&nbsp;חסומים&nbsp;</td>
   <td class="11normalB" nowrap align="center" valign=middle>
	<%if trim(sort)="3" then
        Myclass = "adm"
        Mystr = "default.asp?sort=4"
        Mypicture = "../images/up_sort.gif"
        elseif trim(sort)="4" then
        Myclass = "adm"
        Mystr = "default.asp?sort=3"
        Mypicture = "../images/down_sort.gif"
        else
        Myclass =  "adm"
        Mystr = "default.asp?sort=3"
        Mypicture = ""
        end if           
    %>&nbsp;<a class="<%=Myclass%>" href="<%=Mystr%>" valign=middle>בעבודה</a>&nbsp;
   </td>
   <td width="100%" class="11normalB" align="right">
    <%if trim(sort)="1" then
        Myclass = "adm"
        Mystr = "default.asp?sort=2"
        Mypicture = "../images/up_sort.gif"
        elseif trim(sort)="2" then
        Myclass = "adm"
        Mystr = "default.asp?sort=1"
        Mypicture = "../images/down_sort.gif"
        else
        Myclass =  "adm"
        Mystr = "default.asp?sort=1"
        Mypicture = ""
        end if           
    %><a class="<%=Myclass%>" href="<%=Mystr%>">שם ארגון</a>&nbsp;
   </td>
</tr>
<%
sqlStr = "Select ORGANIZATION_ID, rtrim(ltrim(ORGANIZATION_NAME)), ACTIVE, IsNull(IN_WORK,0) from ORGANIZATIONS Order BY " & sortby(sort) 
'Response.Write sqlStr
set rs_orgs = con.GetRecordSet(sqlStr)
If not rs_orgs.EOF Then
    numOfOrgs = rs_orgs.RecordCount
	arrOrgs = rs_orgs.getRows()
End If	
set rs_orgs = nothing

If IsArray(arrOrgs) Then
For count=0 To numOfOrgs-1
	ORGANIZATION_ID = arrOrgs(0,count)
	ORGANIZATION_NAME = trim(arrOrgs(1,count))
	perSite = arrOrgs(2,count)
	inWork = arrOrgs(3,count)
%>
	<tr>
		<td align="center" valign="top" bgcolor="#DDDDDD">
		<%If trim(inWork) <> "1" Then%>
		<a href="default.asp?delORGANIZATION_ID=<%=ORGANIZATION_ID%>&sort=<%=sort%>" ONCLICK="return CheckDel()"><img src="../images/delete_icon.gif" border="0" alt="מחיקה"></a>
		<%End If%>
		</td>		
	    <td align="center" valign="top" bgcolor="#DDDDDD" nowrap><a href="vsbPress.asp?idsite=<%=ORGANIZATION_ID%>"><%if perSite = "0" then%><img src="../images/lamp_off.gif" alt="לא מופיע באתר" border="0" WIDTH="13" HEIGHT="18"><%else%><img src="../images/lamp_on.gif" alt="מופיע באתר" border="0" WIDTH="13" HEIGHT="18"><%end if%></a></td>
	    <td align="center" valign="top" bgcolor="#DDDDDD" nowrap>
	    <INPUT type=image src="../../netcom/images/bull2.gif" border=0  onclick="window.open('sum_files.asp?ORGANIZATION_ID=<%=ORGANIZATION_ID%>','sum_window','top=50,left=50,width=660,height=500,scrollbars=1');return false;" title="רשימת קבצים מצורפים">
	    </td>
    	<td align="center" bgcolor="#dddddd" nowrap>&nbsp;<a href="permitions.asp?ORGANIZATION_ID=<%=ORGANIZATION_ID%>"><IMG SRC="../../netcom/images/signIn_icon.gif" BORDER=0 alt="הרשאות"></a>&nbsp;</td>		
	    <td align="center" valign="top" bgcolor="#DDDDDD"><a href="workers.asp?ORGANIZATION_ID=<%=ORGANIZATION_ID%>"><img src="../images/users.gif" border="0" alt="רשימת עובדים"></a></td>
	    <td align="center" valign="top" bgcolor="#DDDDDD"><a href="jobs.asp?ORGANIZATION_ID=<%=ORGANIZATION_ID%>"><img src="<%=Application("VirDir")%>/netcom/images/user_icon.gif" border="0" alt="רשימת תפקידים"></a></td>
	    <td align="center" valign="top" bgcolor="#DDDDDD"><a href="groups.asp?ORGANIZATION_ID=<%=ORGANIZATION_ID%>"><img src="<%=Application("VirDir")%>/netcom/images/write.gif" border="0" alt="רשימת קבוצות"></a></td>
	    <td align="center" valign="top" bgcolor="#DDDDDD" nowrap>
	    <INPUT type=image src="../../netcom/images/maatafa.gif" border=0  onclick="window.open('blocked_mails.asp?OrgID=<%=ORGANIZATION_ID%>','blocked_window','top=50,left=50,width=660,height=500,scrollbars=1');return false;" title="רשימת מיילים חסומים">
	    </td>
		<td align="center" valign="middle" bgcolor="#DDDDDD" nowrap><a href="default.asp?update_work=1&OrgID=<%=ORGANIZATION_ID%>&sort=<%=sort%>"><%if inWork = "0" then%><img src="../images/send_radio_off.gif" alt="הארגון לא נמצא בעבודה" border="0"><%else%><img src="../images/send_radio_on.gif" alt="הארגון נמצא בעבודה" border="0"><%end if%></a></td>
		<td align="right"  valign="top" bgcolor="#DDDDDD" nowrap><a href="addOrg.asp?ORGANIZATION_ID=<%=ORGANIZATION_ID%>" class="admin"><%=ORGANIZATION_NAME%>&nbsp;</a></td>
	</tr>
<%
	Next
%>
	<tr><td align=center height=20px class="normalB" bgcolor="#C7C7C7" colspan=11>נמצאו&nbsp;<%=numOfOrgs%>&nbsp;ארגונים</td></tr>								
<%	
 End If
%>
</table>
</div>
</body>
</html>
<%
set con = nothing
%>