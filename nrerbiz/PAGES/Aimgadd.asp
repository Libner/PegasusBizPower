<%SERVER.ScriptTimeout=3000%>
<!--#INCLUDE file="../../include/reverse.asp"-->
<!--#include file="../../include/connect.asp"-->
<%'on error resume next
   pageId=trim(Request.QueryString("pageId"))
   F=trim(Request.QueryString("F"))  
    
   if request.queryString("C")=0 then 'for deleting a file
   set pr=con.execute("SELECT Page_background FROM PagesTav where page_id= "& pageId)
	    if not pr.eof then
			fileName1=pr("Page_background")
			if fileName1<>"" then
				set fs=server.CreateObject("Scripting.FileSystemObject") ' open a page FILE object!
				fileString= server.mappath("../../download/page_backgrounds") & "/" & fileName1  'deleting of existing file
				if fs.FileExists(fileString) then
					set f=fs.GetFile(fileString) 
					f.Delete TRUE   'TRUE means: delete even if the file is Read-Only 
		            con.Execute("UPDATE PagesTav SET Page_background=null WHERE page_id="&pageId) 'delete	
				end if					
		    end if	
		end if
	end if	
	
	
	if request.queryString("C")=1 then 'upload a file
	set pr=con.execute("SELECT Page_background FROM PagesTav where page_id= "& pageId)
	    if not pr.eof then
			fileName1=pr("Page_background")
			if fileName1<>"" then
				set fs=server.CreateObject("Scripting.FileSystemObject") ' open a page FILE object!
				fileString= server.mappath("../../download/page_backgrounds") & "/" & fileName1  'deleting of existing file
				if fs.FileExists(fileString) then
					set f=fs.GetFile(fileString) 
					f.Delete TRUE   'TRUE means: delete even if the file is Read-Only 
				end if								
		    end if	
		end if
		set upl = Server.CreateObject("SoftArtisans.FileUp") 
		upl.path= server.mappath("../../download/page_backgrounds")
		pageFileName=Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1)
		pageFileName=pageId&"_"&pageFileName
		pageFileName=replace(pageFileName," ","_")
		if fileName1<>"" then
			if Right(fileName1,Len(fileName1)-1)=pageFileName then
				if left(pageFileName,1)="a" then
					pageFileName="b" & pageFileName
				else
					pageFileName="a" & pageFileName
				end if
			else
				pageFileName="a" & pageFileName
			end if
		else
			pageFileName="a" & pageFileName
		end if
		upl.Form("UploadFile2").SaveAs pageFileName 
		con.Execute("UPDATE PagesTav SET Page_background='"&sFix(pageFileName)&"' WHERE page_id="&pageId)'update
		'Change picture width
		Set objImageGen = Server.CreateObject("softartisans.ImageGen")
		objImageGen.LoadImage Server.MapPath("../../download/page_backgrounds" & "/" & pageFileName)
		if objImageGen.Width > 780 then			
			objImageGen.ResizeImage 780, objImageGen.Height, 2, 0								
			objImageGen.SaveImage saiFile, objImageGen.ImageFormat, Server.MapPath("../../download/page_backgrounds" & "/" & pageFileName)
		end if	
		if objImageGen.Height > 175 then			
			objImageGen.ResizeImage objImageGen.Width, 175, 2, 0								
			objImageGen.SaveImage saiFile, objImageGen.ImageFormat, Server.MapPath("../../download/page_backgrounds" & "/" & pageFileName)
		end if		   
		Set objImageGen = Nothing		
	End If
		
	Response.Redirect "EditPage.asp?pageID=" & pageId	
    
Set con=Nothing
%>
