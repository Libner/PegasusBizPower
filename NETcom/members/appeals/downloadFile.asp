<% If Request.QueryString("document_file")<>nil Then
     document_file=Request.QueryString("document_file")	 
	
     Response.Addheader "Content-Disposition", "Attachment; filename=" & document_file  
	 Response.ContentType = "application/x-msdownload"
	 Response.CacheControl = "public"
	 
	 Set download = Server.CreateObject("SoftArtisans.FileUp")
	 download.TransferFile Server.Mappath( "../../../download/documents/" & document_file)
	 Set download=Nothing
End If %>

