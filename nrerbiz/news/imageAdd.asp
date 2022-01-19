<%SERVER.ScriptTimeout=3000%>
<%facultyId=Request.QueryString("catID")%>
<!--#include file="../../include/connect.asp"-->
<!--#include file="../../include/reverse.asp"-->
<!--#INCLUDE FILE="../checkAuWorker.asp"-->
<%
catID=Request("catID")
New_id=Request("New_id")

	if request.queryString("C")=0 then 'for deleting a file
	set pr=con.execute("SELECT New_Picture FROM news where New_id= "& New_id)
	    if not pr.eof then
			fileName1=pr("New_Picture")
			if fileName1<>"" then
				set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
				fileString= server.mappath("../../download/news") & "/" & fileName1  'deleting of existing file
				if fs.FileExists(fileString) then
					set f=fs.GetFile(fileString) 
					f.Delete TRUE   'TRUE means: delete even if the file is Read-Only 
		            con.Execute("UPDATE news SET New_Picture=null WHERE New_id="&New_id) 'delete	
				end if
				
				fileString= server.mappath("../../download/news") & "/small/" & fileName1  'deleting of existing file
				if fs.FileExists(fileString) then
					set f=fs.GetFile(fileString) 
					f.Delete TRUE   'TRUE means: delete even if the file is Read-Only 
				end if				
		    end if	
		end if
	end if	
	
	
	if request.queryString("C")=1 then 'upload a file
	set pr=con.execute("SELECT New_Picture FROM news where New_id= "& New_id)
	    if not pr.eof then
			fileName1=pr("New_Picture")
			if fileName1<>"" then
				set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
				fileString= server.mappath("../../download/news") & "/" & fileName1  'deleting of existing file
				if fs.FileExists(fileString) then
					set f=fs.GetFile(fileString) 
					f.Delete TRUE   'TRUE means: delete even if the file is Read-Only 
				end if
				
				fileString= server.mappath("../../download/news/small") & "/" & fileName1  'deleting of existing file
				if fs.FileExists(fileString) then
					set f=fs.GetFile(fileString) 
					f.Delete TRUE   'TRUE means: delete even if the file is Read-Only 
				end if
								
		    end if	
		end if
		set upl = Server.CreateObject("SoftArtisans.FileUp") 
		upl.path= server.mappath("../../download/news")
		newFileName=Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1)
		newFileName=New_id&"_"&newFileName
		newFileName=replace(newFileName," ","_")
		if fileName1<>"" then
			if Right(fileName1,Len(fileName1)-1)=newFileName then
				if left(newFileName,1)="a" then
					newFileName="b" & newFileName
				else
					newFileName="a" & newFileName
				end if
			else
				newFileName="a" & newFileName
			end if
		else
			newFileName="a" & newFileName
		end if
		upl.Form("UploadFile1").SaveAs newFileName 
		con.Execute("UPDATE news SET New_Picture='"&newFileName&"' WHERE New_id="&New_id)'update
    
		'//START OF SoftArtisans.ImageGen   
	   %>
	<!--METADATA TYPE="TypeLib" UUID="{50D94450-589A-409B-819A-4F88B151C134}" -->   
	   <%
			Dim saImg
			Set saImg = Server.CreateObject("SoftArtisans.ImageGen")

			'//START OF make small pics										
			saImg.LoadImage Server.MapPath("../../download/news/" & NewFileName)
			saImg.ResizeImage 85, 999999, saiBicubic,1
											
			saImg.JPEGProgressive = false
			Response.Clear
			path = Server.MapPath("../../download/news/small/" & NewFileName)
			saImg.SaveImage saiFile,saiJPG, path
			'//END OF make small pics  
			
			'//START OF make large pics										
			saImg.LoadImage Server.MapPath("../../download/news/" & NewFileName)
			saImg.ResizeImage 150, 999999, saiBicubic,1
											
			saImg.JPEGProgressive = false
			Response.Clear
			path = Server.MapPath("../../download/news/" & NewFileName)
			saImg.SaveImage saiFile,saiJPG, path
			'//END OF make large pics  		 
			
		'//END OF SoftArtisans.ImageGen      
    end if  
 con.Close%>
<html>
<head>
<title>Administration</title>
<meta http-equiv="refresh" content="0;url=Aparaadd.asp?catID=<%=catID%>&id=<%=New_id%>">
</head>
</html>
