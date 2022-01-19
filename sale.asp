<!--#include file="netcom/connect.asp" -->
<!--#include file="netcom/reverse.asp" -->
<%	'pageId = trim(Request("pageId"))
	'prodId = trim(Request("prodId"))
	'C = trim(Request("C"))
	'U = trim(Request("U"))	
	pageId = Decode(Request(Encode("pageId")))
    prodId = Decode(Request(Encode("prodId")))
    C = Decode(Request(Encode("C")))
    U = Decode(Request(Encode("U")))
    fromID = Decode(Request(Encode("fromID")))
    
	url = strlocal & "/netcom/members/pages/result.asp?pageId="&pageId&"&prodId="&prodId&"&C="&C&"&U="&U&"&fromID="&fromID	
	
	lang_id = 1
	If IsNumeric(pageId) And trim(pageId) <> "" Then
		sqlstr = "SELECT Lang_Id From dbo.ORGANIZATIONS O INNER JOIN dbo.Pages P " & _
		" ON O.ORGANIZATION_ID = P.ORGANIZATION_ID WHERE P.Page_Id = " & cLng(pageId)
		set rs_tmp=con.getRecordSet(sqlstr)
		if not rs_tmp.EOF then
			lang_id = cInt(rs_tmp("Lang_Id"))
		end if    
		rs_tmp.close
		set rs_tmp=Nothing		
	End If
	
    If IsNumeric(prodId) And IsNumeric(C) Then
        sqlstr = "UPDATE PRODUCT_CLIENT SET IS_OPENED = 1 WHERE PRODUCT_ID = " & prodId & " AND PEOPLE_ID = " & C
        con.ExecuteQuery(sqlstr)
    End If
                 
    If trim(fromID) <> "" Then
		sql="SELECT category_ID FROM statistic_from_banner WHERE isAction=1 AND category_ID = " & fromID & " AND PAGE_ID = " & pageID & " AND referer IS NOT NULL"
		set checkCat=con.getRecordSet(sql)
		if not checkCat.EOF then
			catID=checkCat("category_ID")
		end if    
		checkCat.close
		set checkCat=Nothing
    Else
	    catID = 0 		
    End If  
    
    If trim(catID)<>"" And trim(pageID) <> "" Then    
			sql="SELECT * FROM statistic_from_banner_counter WHERE category_ID= " & catID & " AND Page_ID = " & pageID
            set checkTable=con.getRecordSet(sql)
            if checkTable.EOF then
               sqlCount="INSERT INTO statistic_from_banner_counter (category_ID,counter,page_ID) VALUES (" & catID & ",1," & pageID & ")"
            else
               sqlCount="UPDATE statistic_from_banner_counter SET counter=counter+1 WHERE category_ID="& catID &" AND Page_ID = " & pageID
            end if
            'Response.Write ("<br>sqlCount="& sqlCount)
            'Response.End
            con.ExecuteQuery(sqlCount)
            checkTable.close
            set checkTable=Nothing
     End If%>
<html>
<head>
<!-- #include file="netcom/title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
</head>
<frameset framespacing="0" border="false" frameborder="0" rows="100%"  cols=100%>
  <frame name="top" target="middle" marginwidth="0" marginheight="0" 
  hspace="0" vspace="0" topmargin="0" leftmargin="0" rightmargin="0" src="<%=url%>" scrolling=yes> 
  <noframes>
  <body>
  <p>This page uses frames, but your browser doesn't support them.</p>
  </body>
  </noframes>
</frameset>
</html>