<%@ Language="VBScript" %>
<!--#include file="_cls2020Upload.asp"-->
<%
dim t1,t2
t1 = Timer()
' TIP: Delete all the comments to see the code
' flow more clearly.

' Define a variable to store the clsUpload object.
dim objUpload

' Instantiate the clsUpload object.
set objUpload = New clsUpload

' Turn on file extensions restriction.
    objUpload.RestrictFileExtentions = True

' Set a list of known "safe" file extensions.
    objUpload.SafeFileExtensions = "txt|jpg|gif|html|bmp"

' Call the Upload() method to retrieve data from
' the HTTP header. Prepare for failure.
' Just as example, I've set the path to "" in this
' case.  You could set the folder path here, or as
' in this example you can set the folder path for
' each file.
' (look for the objFile.UploadPath stuff later).
	IF NOT objUpload.Upload("") THEN
	%>
	<p><%=err.Number &": "& err.Description%>
	<%
	END IF

' As a courtesy, I'll write some data to the page.
' Note that Windows Server 2003 imposed a 200 kilobyte size restriction
' on HTTP headers called "AspMaxRequestEntityAllowed"
' This is documented here: http://msdn.microsoft.com/library/default.asp?url=/library/en-us/iissdk/html/99b8a8bd-f9e7-43a8-b4cf-1186e2b3e9e2.asp
response.write("<p>Total Bytes Uploaded: "& objUpload.TotalBytes &"</p>")

' Check if some files exists in the HTTP header.
	IF objUpload.Files.Count = 0 THEN
	%>
	<p>No files were uploaded or no files have &quot;Safe&quot; file extensions.</p>
	<%
' If there are files, then let's proceed.
	ELSE

' Create variables for later use.
	dim item
	dim objFile
	dim strFileName
	dim strFileExtension

' Loop through the file objects.
		FOR EACH item IN objUpload.Files

' Set the objFile variable to a reference to this item in the Files collection.
		set objFile = objUpload.Files(item)

' Check to see that the folder exists. (We can do this before "saving" the file.)
			IF NOT objFile.FolderExists THEN
			%>
			<p>Folder doesn't exist.  Will reset the folder to: 
			<%
' In this example, I KNEW that the folder wouldn't exist, so I'll reset that
' property here.
			objFile.UploadPath = Replace(server.MapPath(request.servervariables("SCRIPT_NAME")),Mid(request.servervariables("SCRIPT_NAME"),2),"")
' I've reset the "UploadPath" property, so now I'll explicitly call the
' "BuildPath()" method just as a courtesy.
			objFile.BuildPath
			%><%=objFile.UploadPath%></p>
			<%
			END IF

' Let's grab the file name.
		strFileName = objFile.FileName

' And write more stats to the page.
		response.write("<p>File: "& strFileName &"<br />Size: "& objFile.Size &"<br />Content-Type:"& objFile.ContentType)

' We can check to see if this is a duplicate file name.
			IF objFile.FileExists THEN
' Grab the file extension.
			strFileExtension = LCase(Right(strFileName,Len(strFileName) - InStrRev(strFileName,".")))
' Make a new file name.
			strFileName = objFile.Size &"-"& Replace(CDbl(Now()),".","") &"."& strFileExtension
' Reset the file name.
			objFile.FileName = strFileName
' More stats.
			response.write("<br />File Already Exists.  Name changed to: "& strFileName)
			END IF

' Now "Save" the file. (Prepare for failure.)
			IF NOT objFile.Save THEN
			%>
			<p>Cannot save file.</p>
			<%
			ELSE
' Great!
			%>
			<p>File Uploaded and Saved to <%=objFile.UploadPath%></p>
			<%
			END IF

		response.write("</p>")

' Cleanup the file object.
		set objFile = nothing

		NEXT

	END IF

' Check if some form elementss exist in the HTTP header.
	IF objUpload.Form.Count = 0 THEN
	%>
	<p>No form elements exist.</p>
	<%
' If so, then let's just write them to the page.
	ELSE

' Create a variable to store the dictionary item.
	dim element

' Then loop through them and write them to the page.
		FOR EACH element IN objUpload.Form
		%>
		<p>Name: <%=element%><br />Value: <%=objUpload.Form(element)%></p>
		<%
		NEXT

	END IF

' Cleanup
set objUpload = nothing

t2 = Timer()
response.write("Time: "& t2-t1)
%>