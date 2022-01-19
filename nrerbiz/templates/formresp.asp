<%SERVER.ScriptTimeout=3000%>
<!--#include file="..\..\netcom/includes/connect.asp"-->
<!--#include file="..\..\netcom/includes/reverse.asp"-->
<%
if session("admin.username") <> nil then	
	username=session("admin.username")
	password=session("admin.password")
else 
	Response.Redirect("../default.asp")	
end if
%>
<html>

<head>
<title>Document Upload</title>
<meta charset="windows-1255">

<%elemId=Request.QueryString("elemId")
  pageId=Request.QueryString("pageId")
  place=CInt(Request.QueryString("place"))
newPlace=place+1
isRightMenu=Request.querystring("isright")

'on error resume next
sqlstring="SELECT * from workers where loginName='" & userName & "' and  password='"& password &"'"
set worker=con.Execute(sqlstring)
if worker.EOF then%>
<p><center><font color="red" size="4"><b>You are not authorized to enter the BEZEQ staff zone !!!</b></font></center>
<%else
  session("admin.username")=username
  session("admin.password")=password
    if elemId=nil then
      sUpd="UPDATE Elements set Element_Num_On_Page=Element_Num_On_Page+1 WHERE Element_Num_On_Page>="&newPlace&" and Page_Id="&pageId&" "
	  con.Execute(sUpd)
	  sqlstring="INSERT INTO Elements(Page_ID,Element_Num_On_Page,Element_Type,Element_Link_Status,Element_Font_Size,Element_Font_Type,Element_Font_Color,Element_Font_Color_Name,Element_Align,Element_Text_Image_Align,Element_HomePage,Element_Connection,Element_MyVisible,Element_Space,Element_Hebrew_Size) values ("&pageId&","&newPlace&",3,'noLink',2,'STRONG','#0000ff','Blue','right','center',0,1,1,0,100) "
      con.Execute(sqlstring)
      set elem=con.execute("select * from Elements ORDER BY Element_ID DESC")
		elemId=CStr(elem("Element_ID"))
      elem.close
    else
		set fileName=con.Execute("SELECT Element_File_Name from Elements where Element_Id="&elemId ) 
		set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
			fileString= Server.MapPath("../../download_h/"&fileName("Element_File_name")) 'server.MapPath is the current path we're in
		if fs.FileExists(fileString) then
			set f=fs.GetFile(fileString) 
			f.Delete TRUE   'TRUE means: delete even if the file is Read-Only 
		end if
	end if
%>

<meta http-equiv="refresh" content="0;url=addFile.asp?elemId=<%=elemId%>&isright=<%=isRightMenu%>">

</head>
<body bgcolor="#ffffff">
<div align=right>
<%
	set upl=Server.CreateObject("SoftArtisans.FileUp")
	upl.path=server.mappath("../../download_h")
	NewFileName=Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1)
NewFileName=replace(NewFileName," ","_")	
'NewFileName=CStr(hour(now))+CStr(minute(now))+NewFileName
	upl.Form("File1").SaveAs NewFileName 

con.Execute "update Elements Set Element_File_Name='"&NewFileName&"' where Element_Id="&elemId&" "
con.close
end if %>
</body>
</html>
