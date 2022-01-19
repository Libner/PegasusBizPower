<%SERVER.ScriptTimeout=3000%>
<html>
<head>
<title>Administration</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
</head>
<%
Function RandomPW(myLength)
	Dim X, Y, strPW
	For X = 1 To myLength
		'Randomize the type of this character
		Y = 1  '(1) Numeric, (2) Uppercase, (3) Lowercase
		Select Case Y
			Case 1
				'Numeric character
				Randomize
				strPW = strPW & CHR(Int((9 * Rnd) + 48))
			Case 2
				'Uppercase character
				Randomize
				strPW = strPW & CHR(Int((25 * Rnd) + 65))
			Case 3
				'Lowercase character
				Randomize
				strPW = strPW & CHR(Int((25 * Rnd) + 97))
		End Select
	Next
	RandomPW = strPW
End Function
serv_path = Request.ServerVariables("SERVER_NAME")
ins_mode = Request.QueryString("ins")
if ins_mode = "img" then
	str_mappath = "../../download/pictures"
elseif ins_mode = "file" then
	str_mappath = "../../download/files"
elseif ins_mode = "table" then
	str_mappath = "../../download/tbl_bgr_images"
elseif ins_mode = "flash" then
	str_mappath = "../../download/flash"	
end if

set fs = server.CreateObject("Scripting.FileSystemObject") 

set upl = Server.CreateObject("SoftArtisans.FileUp")
 
upl.path=server.mappath(str_mappath)
File_Name=Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1)
File_Name=replace(File_Name," ","_")
extend = LCase(Mid(File_Name,InstrRev(File_Name,".")+1))
name_without_extend = LCase(Mid(File_Name,1,Len(File_Name)-Len(extend)-1))
new_name = name_without_extend
upload = true
	do while fs.FileExists(Server.MapPath(str_mappath & "/"& new_name & "." & extend ))	
		new_name = name_without_extend & "_" & RandomPW(4)
	loop
	newFileName = new_name & "." & extend
	fileString= Server.MapPath(str_mappath & "/"& newFileName ) 
	%>
	<!--SCRIPT LANGUAGE=javascript>
	
		alert(" '<%=newFileName%>' הקובץ בשם\n        .כבר קיים\n\n   .אנא בחר שם אחר")
	
	</SCRIPT-->	
<%	
	upl.Form("UploadFile1").SaveAs newFileName
	if fs.FileExists(fileString) then
		if ins_mode = "file" then%>
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.parent.main_frame.txtFileName.newFileName = '<%=newFileName%>';
			//-->
			</SCRIPT>	
	<%	end if %>
	<SCRIPT LANGUAGE=javascript>
	<!--
		window.parent.main_frame.txtFileName.value = '<%=str_mappath%>/<%=newFileName%>';
		window.parent.main_frame.btnOKClick();
	//-->
	</SCRIPT>	
<%	else %>
	<SCRIPT LANGUAGE=javascript>
	<!--
		window.parent.main_frame.btnOKClick();
	//-->
	</SCRIPT>
<%	end if
set upl = nothing
set fs = nothing
	%>
</html>
