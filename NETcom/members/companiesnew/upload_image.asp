<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%ContactId = trim(Request.QueryString("ContactId"))
	  If trim(ContactId) <> "" Then
			sql="SELECT contact_image, contact_name  FROM dbo.Contacts WHERE contact_ID=" & ContactId
			Set listContact=con.GetRecordSet(sql)
			If not listContact.EOF Then 
				contact_image = trim(listContact("contact_image"))
				contact_name = trim(listContact("contact_name"))
			End If   
			Set  listContact = Nothing
	 End If
	 
	If Request.QueryString("delete") <> nil then
		If trim(contact_image) <> "" Then
			Set fs = Server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
			file_path="../../../download/pictures/" & trim(contact_image)
			If fs.FileExists(server.mappath(file_path)) Then
				Set fd = fs.GetFile(server.mappath(file_path))
				fd.delete 
				Set fd = Nothing
			End If
			Set fs = Nothing
			
			sqlUpd="UPDATE contacts Set contact_image = NULL WHERE contact_ID="& ContactId 
			'Response.Write(sqlUpd)
			'Response.End
			con.ExecuteQuery(sqlUpd)
			Set con = Nothing%>			
		<script language="javascript">
		<!--	
			opener.focus();
			opener.window.location.reload(true);
			self.close();
		//-->
		</script>			
	<%End If	
	
	ElseIf Request.QueryString("add") <> nil then
		Set upl=Server.CreateObject("SoftArtisans.FileUp")		
		If upl.TotalBytes > 6291456 Then%>
			<script language="javascript">
			<!--
				window.alert("ניתן לצרף מסמך עד 6 MB");
				window.history.back();
			//-->
			</script>
<%  Response.End
		Else
			str_mappath="../../../download/pictures"
			File_Name=Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1)
			extend = LCase(Mid(File_Name,InstrRev(File_Name,".")+1))
			name_without_extend = LCase(Mid(File_Name,1,Len(File_Name)-Len(extend)-1))
			NewFileName =  "contact_image_" & companyID
			new_name = NewFileName
			upload = true
			i = 0
			Set fs = server.CreateObject("Scripting.FileSystemObject") 
			do while fs.FileExists(Server.MapPath(str_mappath & "/"& new_name & "." & extend ))	
				i =  i + 1
				new_name = new_name & "_" & i
			loop
			newFileName = new_name & "." & extend
			Set fs = Nothing
		
			upl.Form("contact_image").SaveAs Server.Mappath("../../../download/pictures/") & "/" & NewFileName
			
			sqlUpd="UPDATE contacts Set contact_image = '" & sFix(NewFileName) & "' WHERE contact_ID="& ContactId 
			'Response.Write(sqlUpd)
			'Response.End
			con.ExecuteQuery(sqlUpd)
			Set upl = nothing				
			Set con = Nothing %>
		<script language="javascript">
		<!--	
			opener.focus();
			opener.window.location.reload(true);
			self.close();
		//-->
		</script><%	
		End If 
  End If	   %>
<!--#include file="../checkWorker.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<!-- #include file="../../title_meta_inc.asp" -->
	<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
	<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	<meta name=vs_defaultClientScript content="JavaScript">
	<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
	<script language="javascript">
	<!--		
		function checkfile()
		{
			if (document.Form1.contact_image.value=='')
			{
					window.alert('! נא לבחור קובץ');
					document.Form1.contact_image.focus();
					return false;
			}
			else	
			{
					var fname = new String();
					var fext = new String();
					var extfiles = new String();
					fname = document.Form1.contact_image.value;
					indexOfDot = fname.lastIndexOf('.')
					fext = fname.slice(indexOfDot+1,-1)		
					fext = fext.toUpperCase();
					extfiles='JPG,JPEG,GIF,BMP,PNG';		
					if(extfiles.indexOf(fext)>-1)
						return true;
					else
					{
						window.alert(':סיומת הקובץ - אחת מרשימה' + '\n' + extfiles);
						return false;
					}	
			}  
			return true;
		}

		function exit_()
		{
			this.close();	
			return false;
		}
	//-->
	</script>	
</head>
<body style="margin: 0px; background-color: #e6e6e6;" onload="self.focus()">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="100%" class="page_title" dir="ltr" colspan="2">העלאת תמונה&nbsp;&nbsp;<font color="#6F6DA6"><%=contact_name%></font>&nbsp;</td>
</tr> 
<tr><td height=20 nowrap></td></tr>        
	<tr>
		<td align="right" width="100%"> 
		<form name="Form1" action="upload_image.asp?add=1&ContactId=<%=ContactId%>" method="post" onsubmit="return checkfile();" target="_self" ENCTYPE="multipart/form-data" ID="Form1">
		<table border="0" cellpadding="1" cellspacing="3" width="100%" align="right" dir="ltr" >        
			<tr>
				<td align="right" width="100%" class="card">
				<input dir="ltr" type="file" name="contact_image" style="width: 340px" maxlength="100" ID="contact_image">
				</td>
				<td width="100" nowrap class="card" dir="rtl" align="center">קובץ תמונה</td>
			</tr>
			<%If trim(contact_image) <> "" Then%>
			<tr><td align="right" width="100%" colspan="2" class="card"><img src="../../../download/pictures/<%=contact_image%>" 
			border="0"></td></tr>
			<tr><td align="right" width="100%" colspan="2" class="card"><a class="button_delete_1" style="width: 100px"
			onclick="return window.confirm('? האם ברצונך למחוק את התמונה');"
			href="upload_image.asp?delete=1&ContactId=<%=ContactId%>">מחק תמונה</a></td></tr>
			<%End If%>
			<tr><td colspan="2" height="15" nowrap></td></tr>
			<tr>
            	<td colspan="2" align="center" nowrap dir="ltr" class="card">
				<INPUT type="button" class="but_menu" onclick="return exit_();" value="סגור" style="width:100px" ID="btnSubmit" 
				NAME="btnSubmit">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT 
				type="submit" class="but_menu" value="שמור" style="width:100px" ID="btnClose" NAME="btnClose">
				</td>
			</tr>
		</table>
	 </form>		
	</td></tr></table>
</body>
</html>
<%Set con = Nothing%>