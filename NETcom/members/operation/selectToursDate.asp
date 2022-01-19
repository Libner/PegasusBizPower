<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%

 
  set xmldoc = Server.CreateObject("Microsoft.XMLDOM")
	xmldoc.async=false
	xmldoc.load(request)

for each x in xmldoc.documentElement.childNodes
		if x.NodeName = "SeriaId" then SeriaId=x.text  
	next
	
	strTemp = ""
'	TourId=64
	
	If trim(SeriaId) <> "" Then    
''		sqlstr = "SELECT Departure_Id, (IsNULL(Departure_Code, '') + ' ' + IsNULL(Date_Begin, '') " & _
''       " + ' - ' + IsNULL(Date_End, '')) as Departure_Name  FROM dbo.Tours_Departures TD " & _
''        " WHERE (Tour_Id =" & TourId &") ORDER BY Departure_Order"
                'AND (Departure_Vis = 1) 
                
                	sqlstr = "SELECT Departure_Id, (IsNULL(Departure_Code, '') + ' ' + IsNULL(Date_Begin, '') ) as Departure_Name  FROM dbo.Tours_Departures TD " & _
       " left join Tours T on T.Tour_Id=TD.Tour_Id " & _
        " WHERE (SeriasId =" & SeriaId &")   and  ( DATEDIFF(dd, GetDate(), TD.Departure_Date)>=-30 or DATEDIFF(dd, GetDate(), TD.Departure_Date) >=0) ORDER BY Departure_Order"
                
        
		set rs_Departures = conPegasus.Execute(sqlstr)	
		if rs_Departures.eof Then
			strTemp = ""
		else
			'separate rows using ";;" and columns using  ";"                         
			strTemp = rs_Departures.Getstring(,,",",";")
		end if
    Else
		strTemp = ""
    End If
    
    set rs_Departures = Nothing
    Set con = Nothing
    set xmldoc = Nothing
    
    Response.Write strTemp
  
    
%>