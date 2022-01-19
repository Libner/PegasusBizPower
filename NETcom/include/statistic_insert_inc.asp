<%  'Dim ServerVariable
	'For Each ServerVariable In Request.ServerVariables
		'Response.Write("<br>"& ServerVariable &":"& Request.ServerVariables(ServerVariable))
	'next
	'Response.End
	HTTP_REFERER  = Request.ServerVariables("HTTP_REFERER")
	MYreferer     = Request.ServerVariables("HTTP_HOST")
	  
	'Response.Write(chr(60) &"!-- HTTP_REFERER="& HTTP_REFERER &" --"& chr(62))
	'Response.Write(chr(60) &"!-- MYreferer="& MYreferer &" --"& chr(62))
	    
	arrPath=Split(HTTP_REFERER,"/")
	if UBound(arrPath) > 1 then
		referer=arrPath(2)
		Response.Write(chr(60) &"!--referer="& referer &"--"& chr(62))
	end if
  
	'response.Write referer
	'response.End
  
	if trim(referer)<>"" AND  Lcase(trim(referer))<>trim(MYreferer) then    
           sql="SELECT counter FROM Statistic WHERE LTRIM(RTRIM(referer))='" & sFix(referer) &_
           "' AND referer IS NOT NULL And DateDiff(d,'" & Date() & "',[Date]) = 0"
           set checkTable=con.getRecordSet(sql)
           if checkTable.EOF then
               sqlCount="INSERT INTO Statistic (Page_ID,Product_ID,counter,Date,referer) VALUES ('" &_
               pageId & "','" & quest_id & ", 1, getDate(), '" & sFix(referer) &"')"
           else
               sqlCount="UPDATE statistic_counter SET counter=counter+1 WHERE DateDiff(d,'" & Date() &_
               "',[Date]) = 0 AND LTRIM(RTRIM(referer))='"& sFix(referer) &"' AND referer IS NOT NULL"
               If Len(pageId) > 0 Then
					sqlCount = sqlCount & " And Page_ID = " & pageId
               End If
               If Len(quest_id) > 0 Then
					sqlCount = sqlCount & " And Product_ID = " & quest_id
               End If               
           end if
           con.getRecordSet(sqlCount)
	end if%>