<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%

 
  set xmldoc = Server.CreateObject("Microsoft.XMLDOM")
	xmldoc.async=false
	xmldoc.load(request)

for each x in xmldoc.documentElement.childNodes
		if x.NodeName = "DepartureId" then DepartureId=x.text  
	next
	
	strTemp = ""
'	TourId=64
	
	If trim(DepartureId) <> "" Then    
		sqlstr = "SELECT Guide_Id  FROM dbo.Tours_Departures WHERE Departure_Id =" & DepartureId 
		set rs_Guides = conPegasus.Execute(sqlstr)	
		if rs_Guides.eof Then
			strTemp = ""
		else
			'separate rows using ";;" and columns using  ";"                         
			strTemp = rs_Guides("Guide_Id")
		end if
    Else
		strTemp = ""
    End If
    
    set rs_Guides = Nothing
    Set con = Nothing
    set xmldoc = Nothing
    
    Response.Write strTemp
  
    
%>