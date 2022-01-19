<!--#include file="netcom/connect.asp"-->
<!--#include file="netcom/reverse.asp"-->


<!--#include file="upload/_clsUpload.asp" --><%
  Response.Buffer = false
  Server.ScriptTimeout = 600

	
%>
<html>
<head>
<title>test</title>
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<script LANGUAGE="javascript">

</script>
</head>
<%
       str_mappath="download/tasks_attachments"			
	  
		Dim uploadsDirVar
		uploadsDirVar = Server.MapPath(str_mappath)
		  
response.write "<br>uploadsDirVar=" & uploadsDirVar
		dim objUpload
		set objUpload = New clsUpload  
		objUpload.Upload uploadsDirVar 'Load http form data
		

'response.write "<br>textfield=" & objUpload.Form("textfield")
response.write "<br>contactID=" & objUpload.Form("contactID")
response.write "<br>quest_id=" & objUpload.Form("quest_id")
response.write "<br>testhidden=" & objUpload.Form("testhidden")
response.write "<br>field40104=" & objUpload.Form("field40104")
response.write "<br>objUpload.Files.Count=" & objUpload.Files.Count
if objUpload.Files.Count > 0 Then
			
				set objFile = objUpload.Files("file1")
response.write "<br>objFile.Size=" & objFile.Size
			dim objFile1
			FOR EACH item IN objUpload.Form.Files
				set objFile1 = objUpload.Files(item)
				exit for
			next
	  	'response.Write("<br>FileName=" & objFile.FileName)
  				If objFile1.Size > 0  Then		
			'response.Write(upl.UserFileName)
			'response.end
			
response.write "<br>objFile.FileName=" & objFile.FileName
					File_Name = objFile.FileName
					extend = LCase(Mid(File_Name,InstrRev(File_Name,".")+1))
					NewFileName =  "milatest" 
					new_name = NewFileName
					upload = true
					i = 0
					set fs = server.CreateObject("Scripting.FileSystemObject") 
					do while fs.FileExists(Server.MapPath(str_mappath & "/"& new_name & "." & extend ))	
						i =  i + 1
						new_name = NewFileName & "_" & i
					loop
					newFileName = new_name & "." & extend
					set fs = Nothing
				
response.write "<br>objFile.FileName=" & objFile.FileName
					'objUpload.Files("document_file").SaveAs Server.Mappath(str_mappath) & "/" & NewFileName
					objFile.FileName = NewFileName
					objFile.Save						
				End If
	end if	
%>
<body>
<table border="0" width="100%" align=center cellspacing="0" cellpadding="0" ID="Table1">
		 <!-- <form action="testPost.asp" id="form1" name="form1" method="post" ENCTYPE="multipart/form-data">	-->
			<form id="form1" name="form1" action="test.asp" method="post" ENCTYPE="multipart/form-data">		
			<tr>
				<td  style="font-size:16pt" align=center width=100% >
				<INPUT type="hidden" id="hiddenfield" name="hiddenfield" value="hiddenvalue">
				<INPUT type="text" id="textfield" name="textfield" value="textvalue">
				<INPUT type="file" id="file1" name="file1" >
				</td>
			</tr>
			<tr>
				<td  style="font-size:16pt" align=center width=100% >
				<INPUT type="submit" id="sub" name="sub" value="submit">
				</td>
			</tr>
			</form>
			</table>
</body>
</html>
